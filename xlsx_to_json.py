import json

from openpyxl import Workbook
from openpyxl import load_workbook


FG_GRAY = 'FFC0C0C0'
FG_WHITE = 'FFFFFFFF'

wb = load_workbook(filename = 'setorial.xlsx')
sheet = wb.worksheets[0]
b3_classification = dict()

setor = None
subsetor = None
segmento = None

for line in range(1, 999):
    if sheet[f'A{line}'].fill.fgColor.value == FG_GRAY:
        setor = sheet[f'A{line}'].value

    if sheet[f'B{line}'].fill.fgColor.value == FG_GRAY:
        subsetor = sheet[f'B{line}'].value

    if sheet[f'C{line}'].fill.fgColor.value == FG_GRAY:
        segmento = sheet[f'C{line}'].value

    codigo = sheet[f'D{line}'].value
    if codigo is not None and len(codigo) == 4:
        b3_classification[codigo] = {'segmento': segmento, 'subsetor': subsetor, 'setor': setor}

with open("setorial.json", "w") as outfile:
    json.dump(b3_classification, outfile, indent=4, sort_keys=True)
