import os
import rubric_parser

if __name__ == '__main__':
    os.makedirs("output")
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')

    print("\n--- COLUMN MAPPINGS ---")
    for messy, correct in rubric["column_mapping"].items(): print(f"   '{messy}'  =  '{correct}'")

    print("\n--- VALID VALUES ---")
    for field, values in rubric["valid_values"].items(): print(f"   '{field}': {values[:5]}")

    print("\n--- REGEX PATTERNS ---")
    for field, pattern in rubric["value_patterns"].items(): print(f"   '{field}': {pattern[:80]}")