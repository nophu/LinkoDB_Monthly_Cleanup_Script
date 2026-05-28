import os
import rubric_parser
import data_parser

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')

    #records = data_parser.parse_data('r_M_FG_sys_Iu_Permit_List_Extractors.xlsx', rubric)
    #records = data_parser.parse_data('r_M_AG_sys_Iu_Extract_Summary.xlsx', rubric)
    #records = data_parser.parse_data('r_sys_Events_Inspection_Details.xlsx', rubric)
    records = data_parser.parse_data('2025-08-05_FSE_Inspections_last_5_years.xlsx', rubric)
    #records = data_parser.parse_data('Master_List_with_Additional_Fields__1_.xlsx', rubric)

    print(f"\nFirst record:")
    print(records[0])