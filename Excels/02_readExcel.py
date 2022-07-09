#pip install xlrd
import xlrd

#open file
wb = xlrd.open_workbook('./Data/weight.xls')

#print the number of sheets
print(f'There are {wb.nsheets} piece of sheets')
#print sheet names
print(f'sheet names:{wb.sheet_names()}')

#select sheet
ws1 = wb.sheet_by_index(0)
#ws2 = wb.sheet_by_name('weight')

#fetching sepcific row or col
print(f'total {ws1.nrows} rows and {ws1.ncols} cols') 
print(f'value from row 1 col 2 {ws1.cell_value(0,2)}')
print(f'value from row 1 col 2 {ws1.cell(0,2).value}')
print(f'value from row 1 col 2 {ws1.row(0)[2].value}')

#all data same row or col
print(f'row 1 data: {ws1.row_values(0)}')
print(f'column 1 data: {ws1.col_values(0)}')

#loop through all
for r in range(ws1.nrows):
    for c in range(ws1.ncols):
        print(f'data from row {r} col {c} is: {ws1.cell_value(r,c)}')