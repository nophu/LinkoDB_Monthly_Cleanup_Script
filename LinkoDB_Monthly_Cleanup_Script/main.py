import os
import rubric_parser
import data_parser
import validator

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)

    # step 1 — read the rubric
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')

    # step 2 — parse a data file
    records = data_parser.parse_data('r_M_FG_sys_Iu_Permit_List_Extractors.xlsx', rubric)

    # step 3 — validate the records
    validated, changes = validator.validate_records(
        records,
        rubric,
        'r_M_FG_sys_Iu_Permit_List_Extractors.xlsx'
    )

    print(f"\nTotal records validated: {len(validated)}")
    print(f"Total changes logged:    {len(changes)}")