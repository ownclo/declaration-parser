import json
import os

from xlrd import open_workbook

dirname = "samples"
fileName = "svedeniya-o-dohodah-sotrudnikov-territorialnyih-organov-roskomnadzora-2015.xls"
schemaName = "schema.json"

schema = json.load(open(schemaName))

fname = os.path.join(dirname, fileName)

book = open_workbook(fname)

sheet = book.sheet_by_index(0)


cell = sheet.cell(0, 0)

# for i in range(sheet.nrows):
#     #    print sheet.cell_type(1,i),sheet.cell_value(1,i)
#     print sheet.cell_value(i, 0)

columnName = schema['columns'][0]['name']
#print columnName

columnFound = False

print sheet.ncols
print sheet.nrows

def findColumn(sheet, columnName):
	for i in xrange(0, sheet.nrows):
		for j in xrange(0, sheet.ncols):
			cell = sheet.cell_value(i, j)
			if columnName in cell:
				return (True, i, j)
	return (False, sheet.nrows, sheet.ncols)

def readColumnData (sheet, rowIndex, columnIndex):
	result = []

	for i in xrange(rowIndex + 1, sheet.nrows):
		cellValue = sheet.cell_value(i, columnIndex)
		if cellValue != "":
			result.append((i, cellValue))
	return result


(isFound, row, column) = findColumn(sheet, columnName)

if isFound:
	column = readColumnData(sheet, row, column)
	print column
