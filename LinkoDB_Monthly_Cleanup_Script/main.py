import os
import rubric_parser
import data_parser

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')

    records = data_parser.parse_data('r_M_FG_sys_Iu_Permit_List_Extractors.xlsx', rubric)
    records = data_parser.parse_data('r_M_AG_sys_Iu_Extract_Summary.xlsx', rubric)
    records = data_parser.parse_data('r_sys_Events_Inspection_Details.xlsx', rubric)

    print(f"\nFirst record preview:")
    print(records[0])