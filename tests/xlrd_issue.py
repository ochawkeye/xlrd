import xlrd

wb = xlrd.open_workbook('xlrd_IndexError1.xlsx')
#wb = xlrd.open_workbook('xlrd_IndexError2.xlsx')

for sheet in wb.sheets():
    print 'Sheet:', sheet.name
    for row in range(sheet.nrows):
        print 'Row:', row
        for col in range(sheet.ncols):
            print '\tCol:', col, sheet.cell(row, col).value
    print
