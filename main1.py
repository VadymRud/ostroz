import os
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
read_dir = os.path.join(BASE_DIR, 'read')
write_dir = os.path.join(BASE_DIR, 'write')
read_file = os.path.join(read_dir, 'НОВИЙ ШТАТ-резервисты.xlsx')
wb = load_workbook(filename=read_file)


ws = wb['штатка']

# for row in ws.rows:
#     for cell in row:
#         print(cell.value)
first_row = list(ws.rows)[:20]
print(first_row)
for row in first_row:
    for cell in row:
        if cell.font.sz == 14.0:
            print(cell.value, cell.font.sz,cell)

