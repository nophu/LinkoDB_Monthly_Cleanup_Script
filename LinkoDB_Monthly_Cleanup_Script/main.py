import os
import rubric_parser
import data_parser
import data_validator

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')
    records = data_parser.parse_data('r_M_FG_sys_Iu_Permit_List_Extractors.xlsx', rubric)
    validated, changes = data_validator.validate_data( records,  rubric, 'r_M_FG_sys_Iu_Permit_List_Extractors.xlsx')
    print(f"\nTotal records validated: {len(validated)}")
    print(f"Total changes logged:    {len(changes)}")