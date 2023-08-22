#!/usr/bin/python3.11
import re
import sys
import pycel
import tempfile
import openpyxl
from tqdm.auto import tqdm
from dataclasses import dataclass

FORMULA_NAME = "DERIVATIVE"
DX = 1e-10

if len(sys.argv) != 3:
  print(f"Usage: {sys.argv[0]} [input_file.xlsx] [output_file.xlsx]", file=sys.stderr)
  exit(-1)

filename = sys.argv[1]
output_filename = sys.argv[2]

def get_cell_addr(address, fallback_sheet):
  address = address.replace("$", "")
  sheet = fallback_sheet
  if ("!" in address):
    sheet, address = address.split("!")
    sheet = sheet.replace("'", "")
  return (sheet, address)

def lookup_cell(wb, addr):
  return wb[addr[0]][addr[1]]

@dataclass
class DerivativeResultCell:
  dst: (str, str)
  y: (str, str)
  x: (str, str)
  orig_x_value: float
  orig_y_value: float

wb = openpyxl.open(filename)
wb_data = openpyxl.open(filename, data_only=True)

resultCells = []

for sheet in wb.sheetnames:
  ws = wb[sheet]
  for row in ws.iter_rows():
    for cell in row:
      formula = None
      if type(cell.value) == openpyxl.worksheet.formula.ArrayFormula and cell.value.text.startswith(f"={FORMULA_NAME}("):
        formula = cell.value.text
      elif type(cell.value) == str and cell.value.startswith(f"={FORMULA_NAME}("):
        formula = cell.value
      
      if formula is not None:
        parts = re.search(r"\(([^,]+),([^,]+),([^,]+)\)", formula)
        dst_ref, y_ref, x_ref = parts.groups()
        dst_addr = get_cell_addr(dst_ref, sheet)
        y_addr = get_cell_addr(y_ref, sheet)
        x_addr = get_cell_addr(x_ref, sheet)
        
        x_val = lookup_cell(wb_data, x_addr).value
        y_val = lookup_cell(wb_data, y_addr).value

        assert type(x_val) == float, type(x_val)
        assert type(y_val) == float, type(y_val)

        resultCells.append(DerivativeResultCell(dst_addr, y_addr, x_addr, x_val, y_val))

wb_data.close()

for resultCell in tqdm(resultCells):
  wb_clone = openpyxl.open(filename)
  
  lookup_cell(wb_clone, resultCell.x).value = resultCell.orig_x_value + DX
  with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
    wb_clone.save(tmp.name)
    wb_clone.close()

    compiler = pycel.ExcelCompiler(filename=tmp.name)
    new_y_value = compiler.evaluate(f"'{resultCell.y[0]}'!{resultCell.y[1]}")
  
  dy = new_y_value - resultCell.orig_y_value
  lookup_cell(wb, resultCell.dst).value = dy/DX

wb.save(output_filename)
wb.close()
