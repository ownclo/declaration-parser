from xlrd import open_workbook, XL_CELL_TEXT
import os

dirname = "samples"
fileName = "svedeniya-o-dohodah-sotrudnikov-territorialnyih-organov-roskomnadzora-2015.xls"

fname = os.path.join(dirname, fileName)

book = open_workbook(fname)

sheet = book.sheet_by_index(0)

cell = sheet.cell(0, 0)
print cell
print cell.value
print cell.ctype == XL_CELL_TEXT

for i in range(sheet.nrows):
    #    print sheet.cell_type(1,i),sheet.cell_value(1,i)
    print sheet.cell_value(i, 0)
