# LIFELINE Phase 1 data generator

Requires Python 3.10+ and `openpyxl`.

From the project root, reproduce the workbook with this one command:

```bash
python generator/generate.py --seed 24 --output cases.xlsx
```

The seed is the group number (24). Each case uses a deterministic pseudo-random stream derived from the group seed and case number. The script writes all 20 sheets in the exact order and schema specified in the Phase 1 brief. Do not edit `cases.xlsx` manually; regenerate it with this command after changing the generator.
