#!/usr/bin/env python3
"""Rebuild valid copies and two independently corrupted five-fault workbooks."""
from pathlib import Path
from shutil import copy2
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parents[1]
TESTS=Path(__file__).resolve().parent/'tests'
MASTER=ROOT/'cases.xlsx'
TESTS.mkdir(exist_ok=True)
for name in ('valid_A.xlsx','valid_B.xlsx','invalid_A_five_faults.xlsx','invalid_B_five_faults.xlsx'):
    p=TESTS/name
    if p.exists():p.unlink()
copy2(MASTER,TESTS/'valid_A.xlsx');copy2(MASTER,TESTS/'valid_B.xlsx')
# A: header, crate range, duplicate serial, invalid road endpoint, duplicate closure.
w=load_workbook(MASTER)
w['cases'].cell(1,1).value='wrong_case_id'
w['crates'].cell(2,4).value=0
w['crates'].cell(3,7).value=w['crates'].cell(2,7).value
w['roads'].cell(2,4).value=999999
w['waves'].cell(3,3).value=w['waves'].cell(2,3).value
w.save(TESTS/'invalid_A_five_faults.xlsx');w.close()
# B: pattern count, tour permutation, safety assignment, ration selection, location penalty.
w=load_workbook(MASTER)
w['radio_patterns'].cell(2,4).value=(w['radio_patterns'].cell(2,4).value or 0)+1
w['tour'].cell(3,3).value=w['tour'].cell(2,3).value
for row in range(2,w['safety_planted'].max_row+1):
    if w['safety_planted'].cell(row,1).value=='24-00':w['safety_planted'].cell(row,3).value=1-w['safety_planted'].cell(row,3).value
for row in range(2,w['rations'].max_row+1):
    if w['rations'].cell(row,1).value=='24-00':w['rations'].cell(row,4).value=0
w['locations'].cell(3,4).value=3
w.save(TESTS/'invalid_B_five_faults.xlsx');w.close()
print('Created two valid copies and two corrupted workbooks (five independent faults each).')
