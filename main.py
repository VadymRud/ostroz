import os
from xlrd import open_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
read_dir = os.path.join(BASE_DIR, 'read')
write_dir = os.path.join(BASE_DIR, 'write')
print(read_dir, write_dir)
read_file = os.path.join(read_dir, 'НОВИЙ ШТАТ-резервисты.xls')
book = open_workbook(read_file, on_demand=True)
sheet = book.sheet_by_index(0)

first_row = [] # Header
# for col in range(sheet.ncols):
#     first_row.append( sheet.cell_value(0,col) )
# # tronsform the workbook to a list of dictionnaries
# data =[]
# for row in range(1, sheet.nrows):
#     elm = {}
#     for col in range(sheet.ncols):
#         elm[first_row[col]]=sheet.cell_value(row,col)
#     data.append(elm)
# print (data)
# r1 = sheet.col_values(0)
# print(r1)
# print(sheet.nrows)
# print(sheet.ncols)


from xlrd.sheet import ctype_text
row = sheet.row(0)
print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))

# Print all values, iterating through rows and columns
#
# num_cols = sheet.ncols   # Number of columns
# for row_idx in range(0, sheet.nrows):    # Iterate through rows
#     print ('-'*40)
#     print ('Row: %s' % row_idx)   # Print row number
#     for col_idx in range(0, num_cols):  # Iterate through columns
#         cell_obj = sheet.cell(row_idx, col_idx)  # Get cell object by row, col
#         print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))
num_cols = sheet.ncols   # Number of columns
for row_idx in range(0, 20):    # Iterate through rows
    print ('-'*40)
    print ('Row: %s' % row_idx)   # Print row number
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        print ( cell_obj.xf_index)
