import os
import rubric_parser
import data_parser
import data_validator

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')
    records = data_parser.parse_data('Master List with Additional Fields.xlsx', rubric)
    validated, changes = data_validator.validate_data( records,  rubric, 'Master List with Additional Fields.xlsx')
    print(f"\nTotal records validated: {len(validated)}")
    print(f"Total changes logged:    {len(changes)}")