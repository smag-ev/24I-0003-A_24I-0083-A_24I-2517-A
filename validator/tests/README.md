# Validator test fixtures

- `valid_A.xlsx` and `valid_B.xlsx` are byte-identical copies of the generated master workbook and must pass.
- `invalid_A_five_faults.xlsx` contains five independent faults: wrong header, out-of-range crate weight, duplicate serial, invalid road endpoint, and a road repeated across closure waves.
- `invalid_B_five_faults.xlsx` contains five independent faults: incorrect radio pattern count, duplicate tour site, inverted planted safety assignment for case 24-00, removed planted ration selection for case 24-00, and invalid location penalty.

Rebuild all four fixtures from the current master workbook with `python validator/make_tests.py`.
