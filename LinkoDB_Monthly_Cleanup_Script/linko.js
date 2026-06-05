/* linko.js — browser/Node port of the Linko cleanup pipeline.
   Uses SheetJS (global `XLSX` in browser, or require in Node). */
(function (root) {
"use strict";

// ---------- config ----------
const REPORT_CONFIG = {
  permit_list: { name:"Permit List – Extractors", match:["permit"],
    fields:["Extractor ID","Extractor Type","Trap Size and Units","Cleaning Frequency"] },
  fse: { name:"FSE Inspections – Last 5 Years", match:["fse"],
    fields:["ReceivingPlant","ClassCode","SecondClass","TrunkLine"] },
  master: { name:"Master List with Additional Fields", match:["master"],
    fields:["MapCategory"] },
  ag: { name:"AG Extract Summary", match:["extract_summary","extract summary","ag_sys"],
    fields:["Extractor ID","Extractor Type"] },
  events: { name:"Inspection Events", match:["events_inspection","inspection_details","_events_"],
    fields:["EventTypeAbbrv"] },
};
const FIELD_DISPLAY = {
  "Extractor ID":"Extractor ID","Extractor Type":"Extractor Type",
  "Trap Size and Units":"Trap Size & Units","Cleaning Frequency":"Cleaning Frequency",
  "ReceivingPlant":"Receiving Plant","ClassCode":"Class Code","SecondClass":"Second Class",
  "TrunkLine":"Trunk Line","MapCategory":"Map Category","EventTypeAbbrv":"Event Type",
};
function matchReport(filename){
  const fn = filename.toLowerCase();
  for(const key in REPORT_CONFIG){
    if(REPORT_CONFIG[key].match.some(kw => fn.includes(kw))) return {key, cfg:REPORT_CONFIG[key]};
  }
  return {key:null, cfg:null};
}

// ---------- helpers ----------
const NEEDED_SHEETS = ['Key','Tables Extractor IDs, Type','6Other Codes_MapCategory_Events',
  'FS - Trap Size & Units','FS - Cleaning Frequency'];
const OTHER_CODES_COLUMNS = {
  "new class codes":"ClassCode","new secondclass":"SecondClass",
  "new mapcategory":"MapCategory","new eventtypeabbrv":"EventTypeAbbrv",
};
const OTHER_CODES_FIELDS = new Set(Object.values(OTHER_CODES_COLUMNS).concat(["ReceivingPlant","TrunkLine"]));
const TABLES_FIELDS = new Set(["Extractor ID","Extractor Type"]);

function S(v){ return (v===null||v===undefined) ? "" : String(v).trim(); }
function sheetToGrid(XLSX, ws){
  // 2D array of trimmed strings, preserving row/col layout
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:"", blankrows:true})
    .map(row => row.map(c => S(c)));
}
function norm(s){ return s.toLowerCase().replace(/[^a-z0-9]/g,""); }
function normSp(s){ return s.toLowerCase().replace(/[^a-z0-9 ]/g,"").trim(); }

// ---------- rubric parsing ----------
function parseRubric(XLSX, wb){
  const rubric = {column_mapping:{}, valid_values:{}, value_patterns:{}};
  const names = wb.SheetNames;

  // Key sheet → column mapping
  const keyName = names.find(n => n.toLowerCase().includes("key"));
  if(keyName){
    const g = sheetToGrid(XLSX, wb.Sheets[keyName]);
    for(const row of g){
      const vals = row.filter(v => v!=="" && v.toLowerCase()!=="nan");
      if(vals.length>=2){
        const messy=vals[0], correct=vals[1];
        if(["download data header","header","field"].includes(messy.toLowerCase())) continue;
        rubric.column_mapping[messy]=correct;
      }
    }
  }

  const ruleSheets = names.filter(n => NEEDED_SHEETS.includes(n));
  for(const sn of ruleSheets){
    if(sn.includes("Other Codes")) extractOtherCodes(XLSX, wb.Sheets[sn], rubric);
    else if(sn.includes("Tables Extractor")) extractTables(XLSX, wb.Sheets[sn], rubric);
    else extractValidValues(XLSX, wb.Sheets[sn], rubric);
  }
  buildPatterns(rubric);
  rubric.valid_values["Trap Size and Units"]=[];
  return rubric;
}

function extractTables(XLSX, ws, rubric){
  const g = sheetToGrid(XLSX, ws);
  let boundary = g.length;
  for(let r=0;r<g.length;r++){
    if(g[r].join(" ").toLowerCase().includes("current linko pick")){ boundary=r; break; }
  }
  const ids=new Set(), types=new Set();
  // Extractor IDs: any "EX###" cell above the boundary, wherever it sits
  for(let r=0;r<boundary;r++){
    for(let c=0;c<g[r].length;c++){
      if(/^EX\s*\d/i.test(g[r][c]||"")) ids.add(g[r][c]);
    }
  }
  // Extractor Types: find the "Extractor Type ... Pick List" header column, read below it
  let typeCol=-1, typeRow=-1;
  for(let r=0;r<boundary&&typeCol<0;r++){
    for(let c=0;c<g[r].length;c++){
      const cl=(g[r][c]||"").toLowerCase();
      if(cl.includes("extractor type")&&cl.includes("pick")){ typeCol=c; typeRow=r; break; }
    }
  }
  if(typeCol>=0){
    for(let r=typeRow+1;r<boundary;r++){
      const cell=g[r][typeCol]||"";
      if(cell===""||cell.toLowerCase()==="nan") continue;
      if(cell.length>30) continue;
      types.add(cell);
    }
  }
  if(ids.size) rubric.valid_values["Extractor ID"]=Array.from(ids).sort();
  if(types.size) rubric.valid_values["Extractor Type"]=Array.from(types).sort();
}

function extractOtherCodes(XLSX, ws, rubric){
  const g = sheetToGrid(XLSX, ws);
  // header row with the most "new ___" labels
  let headerRow=0, best=0;
  for(let r=0;r<Math.min(5,g.length);r++){
    let hits=0;
    for(const c of g[r]) if(OTHER_CODES_COLUMNS[c.toLowerCase()]) hits++;
    if(hits>best){best=hits;headerRow=r;}
  }
  const ncol = Math.max(...g.map(row=>row.length));
  for(let col=0; col<ncol; col++){
    const label = (g[headerRow][col]||"").toLowerCase();
    if(!OTHER_CODES_COLUMNS[label]) continue;
    const field = OTHER_CODES_COLUMNS[label];
    const vals=new Set();
    for(let r=headerRow+1;r<g.length;r++){
      const cell=(g[r][col]||"");
      if(cell===""||cell.toLowerCase()==="nan") continue;
      if(cell.length>40) continue;
      if(cell.toLowerCase().startsWith("delete entry")) continue;
      vals.add(cell);
    }
    if(vals.size) rubric.valid_values[field]=Array.from(vals).sort();
  }
  rubric.valid_values["TrunkLine"]=[];
  // receiving plant sub-header
  const rv=new Set();
  for(let col=0;col<ncol;col++){
    for(let r=0;r<g.length;r++){
      if((g[r][col]||"").toLowerCase()==="new receiving plant"){
        for(let rr=r+1; rr<Math.min(r+10,g.length); rr++){
          const v=(g[rr][col]||"");
          if(v===""||v.toLowerCase()==="nan") continue;
          if(v.toLowerCase().startsWith("delete entry")) continue;
          if(v.length>60) continue;
          rv.add(v);
        }
        break;
      }
    }
  }
  if(rv.size) rubric.valid_values["ReceivingPlant"]=Array.from(rv).sort();
}

function extractValidValues(XLSX, ws, rubric){
  const g = sheetToGrid(XLSX, ws);
  const knownFields = Object.values(rubric.column_mapping);
  for(let r=0;r<g.length;r++){
    for(let c=0;c<g[r].length;c++){
      const matched = matchFieldName(g[r][c], knownFields);
      if(matched && !OTHER_CODES_FIELDS.has(matched) && !TABLES_FIELDS.has(matched)){
        const vals = collectColumnValues(g, r, c);
        if(vals.length){
          const existing = rubric.valid_values[matched]||[];
          rubric.valid_values[matched]=Array.from(new Set(existing.concat(vals))).sort();
        }
      }
    }
  }
}
function matchFieldName(cell, knownFields){
  const cn=normSp(cell);
  for(const f of knownFields){ const fn=normSp(f); if(fn && cn.includes(fn)) return f; }
  return null;
}
function collectColumnValues(g, startRow, col){
  const out=[];
  const fieldNames=["extractor id","extractor type","trap size and units","cleaning frequency",
    "receivingplant","classcode","secondclass","eventtypeabbrv","mapcategory"];
  const junk=["total","rows","script","number","delete","pick list","description","comments",
    "current","future","download","header","series","interceptor","separator","trap -","shared",
    "grease","multiple","facilities","retired","removed","unknown","effluent","verified","inactive",
    "extractor ids","etc"];
  for(let r=startRow+1; r<Math.min(startRow+50, g.length); r++){
    const cell=(g[r][col]||"");
    if(cell===""||cell.toLowerCase()==="nan") continue;
    if(cell.length>40) continue;
    if(/^\d+$/.test(cell)) continue;
    const cl=cell.toLowerCase();
    if(fieldNames.includes(cl)) continue;
    if(junk.some(j=>cl.includes(j))) continue;
    out.push(cell);
  }
  return out;
}
function buildPatterns(rubric){
  for(const field in rubric.valid_values){
    const values=rubric.valid_values[field];
    if(!values||!values.length) continue;
    const p=inferPattern(field, values);
    if(p) rubric.value_patterns[field]=p;
  }
}
function escapeRe(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"); }
function inferPattern(field, values){
  const clean=values.map(S).filter(v=>v);
  if(!clean.length) return null;
  if(field.toLowerCase().includes("extractor id")) return "^EX\\s*[-]?\\s*\\d+";
  const esc=clean.filter(v=>v.length<=40).map(escapeRe);
  if(esc.length) return "^("+esc.join("|")+")$";
  return null;
}

// ---------- data parsing ----------
function detectFileType(g, rubric){
  const row0=(g[0]||[]).map(S), row1=(g[1]||[]).map(S);
  if(S(g[0]&&g[0][0]).includes("Fort Wayne") && row1.includes("txtPermitInfo")) return "inspection_events";
  let extractNameCount=0;
  for(const row of g) for(const v of row) if(S(v).includes("txtExtractName")) extractNameCount++;
  if(extractNameCount>3) return "extract_summary";
  const access=row0.filter(v=>v.startsWith("txt")||["CleaningFreq","ExtractorID","ExtractName"].includes(v)).length;
  if(access>=1) return "permit_list";
  const nonEmpty=row0.filter(v=>v!==""&&v!=="nan").length;
  if(nonEmpty>=5) return "flat_table";
  return "unknown";
}
function parseData(XLSX, wb, rubric){
  const ws=wb.Sheets[wb.SheetNames[0]];
  const g=sheetToGrid(XLSX, ws);
  const ftype=detectFileType(g, rubric);
  if(ftype==="permit_list") return parsePermitList(g, rubric);
  if(ftype==="extract_summary") return parseExtractSummary(g, rubric);
  if(ftype==="inspection_events") return parseInspectionEvents(g, rubric);
  return parseFlat(g, rubric);
}
function rereadWithHeader(g, hr){
  const headers=(g[hr]||[]).map(S);
  const rows=[];
  for(let r=hr+1;r<g.length;r++){
    const obj={};
    for(let c=0;c<headers.length;c++) obj[headers[c]]= (g[r][c]!==undefined? g[r][c] : "");
    rows.push(obj);
  }
  return {headers, rows};
}
function matchColumns(headers, sampleFn, rubric){
  const mapping={};
  for(const col of headers){
    const m=matchByName(col, rubric.column_mapping);
    if(m){ mapping[col]=m; continue; }
    const matched=matchByValues(sampleFn(col), rubric.value_patterns);
    if(matched){ mapping[col]=matched; }
  }
  return mapping;
}
function matchByName(col, cm){
  if(cm[col]) return cm[col];
  const cn=norm(col);
  for(const messy in cm){ if(cn===norm(messy)) return cm[messy]; }
  return null;
}
function matchByValues(values, patterns){
  const sample=values.map(S).filter(v=>v!==""&&v.toLowerCase()!=="nan").slice(0,20);
  if(!sample.length) return null;
  let best=null, bestScore=0;
  for(const field in patterns){
    const re=new RegExp(patterns[field],"i");
    const matches=sample.filter(v=>re.test(v)).length;
    const score=matches/sample.length;
    if(score>bestScore){bestScore=score;best=field;}
  }
  return bestScore>=0.60 ? best : null;
}
function applyMapping(rows, mapping){
  return rows.map(r=>{
    const o={};
    for(const k in r){ o[mapping[k]||k]=r[k]; }
    return o;
  });
}
function cleanRec(r){
  const o={};
  for(const k in r){ const v=S(r[k]); o[k]=(v===""||v==="nan"||v==="NaN")?null:r[k]; }
  return o;
}
function parsePermitList(g, rubric){
  const {headers, rows}=rereadWithHeader(g,0);
  const colVals={};
  headers.forEach((h,i)=>{ colVals[h]=rows.map(r=>r[h]); });
  const mapping=matchColumns(headers, h=>colVals[h]||[], rubric);
  let mapped=applyMapping(rows, mapping);
  // combine trap size + units
  mapped=mapped.map(r=>{
    if(("TrapSize" in r)&&("TrapSizeUnits" in r)){
      const size=S(r["TrapSize"]), units=S(r["TrapSizeUnits"]);
      if(size && size!=="nan") r["Trap Size and Units"]=(size+" "+units).trim();
    }
    return r;
  });
  // merge facility rows with extractor rows
  return mergeByFacility(mapped, ["txtPermittee","txtPermitNo"],
    ["Extractor ID","Extractor Type","Cleaning Frequency","Trap Size and Units"]);
}
function mergeByFacility(rows, facCols, dataCols){
  const out=[]; let cur={};
  for(const r of rows){
    const hasFac=facCols.some(c=>(c in r)&&S(r[c])!==""&&S(r[c])!=="nan");
    const hasData=dataCols.some(c=>(c in r)&&S(r[c])!==""&&S(r[c])!=="nan");
    if(hasFac){ cur={}; facCols.forEach(c=>{ if(c in r) cur[c]=r[c]; }); }
    if(hasData){
      const combined=Object.assign({}, cur);
      for(const k in r){ if(!facCols.includes(k)) combined[k]=r[k]; }
      out.push(cleanRec(combined));
    }
  }
  return out;
}
function parseFlat(g, rubric){
  const messy=Object.keys(rubric.column_mapping);
  let bestRow=0,bestScore=0;
  for(let r=0;r<Math.min(20,g.length);r++){
    const score=g[r].filter(v=>messy.includes(S(v))).length;
    if(score>bestScore){bestScore=score;bestRow=r;}
  }
  const {headers, rows}=rereadWithHeader(g,bestRow);
  const colVals={}; headers.forEach(h=>{ colVals[h]=rows.map(r=>r[h]); });
  const mapping=matchColumns(headers, h=>colVals[h]||[], rubric);
  return applyMapping(rows, mapping).map(cleanRec);
}
function parseExtractSummary(g, rubric){
  const records=[]; const messy=Object.keys(rubric.column_mapping);
  let curFac=null,curPermit=null,curOrder=null;
  for(let r=0;r<g.length;r++){
    const rv=g[r].map(S);
    if(!rv.some(v=>v!==""&&v!=="nan"&&v!=="NaN")) continue;
    const headerMatches=rv.filter(v=>messy.includes(v)).length;
    if(headerMatches>=2){ curOrder=rv; continue; }
    const col0=rv[0]||"", col1=rv[1]||"";
    const restEmpty=rv.slice(2).every(v=>v===""||v==="nan"||v==="NaN");
    const looksPermit=/^[A-Z0-9_][A-Z0-9_\-]+$/i.test(col1);
    if(col0!==""&&col0!=="nan"&&looksPermit&&restEmpty){ curFac=col0;curPermit=col1;continue; }
    if(curOrder){
      const dataVals=rv.slice(2), colNames=curOrder.slice(2);
      if(dataVals.some(v=>v!==""&&v!=="nan"&&v!=="NaN")){
        const rec={SiteCompany:curFac, PermitNo:curPermit};
        for(let i=0;i<colNames.length;i++){
          const cn=colNames[i]; if(cn===""||cn==="nan") continue;
          const correct=rubric.column_mapping[cn]||cn;
          const val=dataVals[i];
          rec[correct]=(val===""||val==="nan"||val==="NaN")?null:val;
        }
        const real=Object.keys(rec).filter(k=>k!=="SiteCompany"&&k!=="PermitNo"&&rec[k]!=null);
        if(real.length) records.push(rec);
      }
    }
  }
  return records;
}
function parseInspectionEvents(g, rubric){
  const {headers, rows}=rereadWithHeader(g,1);
  const colVals={}; headers.forEach(h=>{ colVals[h]=rows.map(r=>r[h]); });
  const mapping=matchColumns(headers, h=>colVals[h]||[], rubric);
  let mapped=applyMapping(rows, mapping);
  const firstCol=Object.keys(mapped[0]||{})[0];
  // EventTypeAbbrv from ContactType
  mapped=mapped.map(r=>{
    if("ContactType" in r){
      const v=S(r["ContactType"]);
      r["EventTypeAbbrv"]=(v===""||v==="nan"||v==="None")?null:v.split(" - ")[0].trim();
    }
    return r;
  });
  const records=[]; let curFacRaw=null;
  for(const r of mapped){
    const fv=S(r[firstCol]);
    if(/^\[\d+]/.test(fv)){ curFacRaw=fv; continue; }
    const hasData=Object.keys(r).some(k=>k!==firstCol && S(r[k])!==""&&S(r[k])!=="nan");
    if(hasData && curFacRaw){
      const facInfo=parseFacilityString(curFacRaw);
      const rec=Object.assign({}, facInfo);
      for(const k in r){ if(k!==firstCol){ const v=S(r[k]); rec[k]=(v==="nan"||v==="NaN")?null:r[k]; } }
      records.push(rec);
    }
  }
  return records;
}
function parseFacilityString(raw){
  const m=raw.match(/^\[(\d+)]\s*(.*)/);
  if(m){
    const rest=m[2].trim();
    const parts=rest.split("   -   ");
    return {PermitID:m[1], FacilityName:(parts[0]||rest).trim(), Address:parts.length>1?parts[1].trim():null};
  }
  return {FacilityInfo:raw};
}

// ---------- validation ----------
function findPartialMatch(value, valid){
  const vl=value.toLowerCase();
  for(const v of valid){ const vv=v.toLowerCase();
    if(vl.startsWith(vv)) return v;
    if(vv.startsWith(vl)) return v;
  }
  return null;
}
function checkTrapSize(value){
  const parts=value.trim().split(/\s+/);
  if(parts.length!==2) return {status:"flagged",cleaned_value:value,note:`expected format '<number> <unit>' (e.g. '35 gpm') — got '${value}'`};
  const sizeStr=parts[0], unit=parts[1];
  const size=parseFloat(sizeStr);
  if(isNaN(size)) return {status:"flagged",cleaned_value:value,note:`'${sizeStr}' is not a valid numeric trap size`};
  const correct = size<=99 ? "gpm":"gal";
  if(unit===correct) return {status:"pass",cleaned_value:value,note:"exact match"};
  if(unit.toLowerCase()===correct){ const f=`${sizeStr} ${correct}`; return {status:"fixed",cleaned_value:f,note:`fixed casing: '${value}' → '${f}'`}; }
  const f=`${sizeStr} ${correct}`;
  return {status:"flagged",cleaned_value:value,note:`unit should be '${correct}' for size ${sizeStr} (rule: ≤99 → gpm, ≥100 → gal) — suggested: '${f}'`};
}
function checkExtractorId(value, valid){
  const ranges=[];
  for(const v of valid){ const nums=(v.match(/\d+/g)||[]).map(Number);
    if(nums.length>=2) ranges.push([nums[0],nums[1]]); else if(nums.length===1) ranges.push([nums[0],nums[0]]); }
  const raw=value.trim();
  let candidate;
  if(/^\d+$/.test(raw)) candidate="EX"+raw;
  else if(/^EX\s*\d+$/i.test(raw)) candidate="EX"+raw.replace(/[^0-9]/g,"");
  else return {status:"flagged",cleaned_value:value,note:`'${value}' is not a standard extractor ID — review manually`};
  const num=parseInt(candidate.replace(/[^0-9]/g,""),10);
  const inRange=ranges.some(([lo,hi])=>num>=lo&&num<=hi);
  if(inRange){
    if(candidate===raw) return {status:"pass",cleaned_value:candidate,note:"valid"};
    return {status:"fixed",cleaned_value:candidate,note:`added EX prefix: '${value}' → '${candidate}'`};
  }
  if(num<100) return {status:"flagged",cleaned_value:value,note:`'${value}' is an old-scheme ID (${candidate}) — needs a new EX1xx–EX8xx ID (review manually)`};
  return {status:"flagged",cleaned_value:value,note:`'${value}' (${candidate}) is not in any valid range (EX100–EX830) — review manually`};
}
function checkValue(field, value, rubric){
  const valid=rubric.valid_values[field]||[];
  const pattern=rubric.value_patterns[field];
  if(field==="Trap Size and Units") return checkTrapSize(value);
  if(field==="Extractor ID") return checkExtractorId(value, valid);
  if(valid.length===0) return {status:"flagged",cleaned_value:value,note:`'${value}' should be deleted — leave this field blank per rubric`};
  if(valid.includes(value)) return {status:"pass",cleaned_value:value,note:"exact match"};
  const vl=value.toLowerCase();
  for(const v of valid){ if(v.toLowerCase()===vl) return {status:"fixed",cleaned_value:v,note:`fixed casing: '${value}' → '${v}'`}; }
  if(pattern && new RegExp(pattern,"i").test(value)) return {status:"pass",cleaned_value:value,note:"matched pattern"};
  const close=findPartialMatch(value, valid);
  if(close) return {status:"flagged",cleaned_value:value,note:`'${value}' is close to '${close}' — should it be changed to '${close}'?`};
  return {status:"flagged",cleaned_value:value,note:`'${value}' is not a valid value for '${field}' — review manually`};
}
function getFacility(rec){
  for(const f of ["txtPermittee","SiteCompany","FacilityName","Permittee","PermitteeAccount","AccountName","Name","FacilityInfo"]){
    const v=rec[f]; if(v && !["","nan","NaN","None"].includes(S(v))) return S(v);
  }
  return "Unknown";
}
function getPermit(rec){
  for(const f of ["txtPermitNo","PermitNo","PermitID","PermitNumber","Permit"]){
    const v=rec[f]; if(v && !["","nan","NaN","None"].includes(S(v))) return S(v);
  }
  return "Unknown";
}
function validateData(records, rubric, sourceKey, onlyFields){
  let checkable=Object.keys(rubric.valid_values);
  if(onlyFields) checkable=checkable.filter(f=>onlyFields.includes(f));
  const changes=[];
  for(const rec of records){
    for(const field in rec){
      const value=rec[field];
      if(value===null||["","nan","NaN"].includes(S(value))) continue;
      if(!checkable.includes(field)) continue;
      const res=checkValue(field, S(value), rubric);
      if(res.status!=="pass"){
        changes.push({source_file:sourceKey, facility:getFacility(rec), permit_no:getPermit(rec),
          field, original:S(value), cleaned_value:res.cleaned_value, status:res.status, note:res.note});
      }
    }
  }
  return changes;
}

// ---------- export to global / module ----------
const API={REPORT_CONFIG, FIELD_DISPLAY, matchReport, parseRubric, parseData, validateData,
  checkValue, checkExtractorId, sheetToGrid};
if(typeof module!=="undefined"&&module.exports) module.exports=API;
else root.Linko=API;

})(typeof window!=="undefined"?window:globalThis);