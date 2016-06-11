import json
import os

from xlrd import open_workbook

dirname = "samples"
fileName = "svedeniya-o-dohodah-sotrudnikov-territorialnyih-organov-roskomnadzora-2015.xls"
schemaName = "schema.json"

schema = json.load(open(schemaName))

fname = os.path.join(dirname, fileName)

book = open_workbook(fname, formatting_info=True)

sheet = book.sheet_by_index(0)


cell = sheet.cell(0, 0)

# for i in range(sheet.nrows):
#     #    print sheet.cell_type(1,i),sheet.cell_value(1,i)
#     print sheet.cell_value(i, 0)

columnName = schema['columns'][0]['name']
#print columnName

print sheet.ncols
print sheet.nrows

def findColumn(sheet, columnName):
	for i in xrange(0, sheet.nrows):
		for j in xrange(0, sheet.ncols):
			cell = sheet.cell_value(i, j)
			if columnName in cell:
				return (True, i, j)
	return (False, sheet.nrows, sheet.ncols)

def readUnmergedColumnData (sheet, rowIndex, columnIndex):
	result = []

	for i in xrange(rowIndex + 1, sheet.nrows):
		cellValue = sheet.cell_value(i, columnIndex)
		if cellValue != "":
			result.append((columnIndex, i, cellValue))

	return result

#lists all the merged cells in the xls
def parseMergedCells (sheet):
	xLow = yLow = yHigh = 0
	result = []

	for crange in sheet.merged_cells:
		yLow, yHigh, xLow, xHigh = crange
		result.append (((xLow, yLow, xHigh, yHigh), sheet.cell_value(yLow, xLow)))

	result = sorted (sorted (result, key=lambda tup: tup[0][0]), key=lambda tup: tup[0][1])

	return result

#merges all the records from merged and non-merged cell lists
def MergeColumnDataOnSheet (startingRowIndex, parsedUnmergedColumnData, parsedMergedCellsData):
	result = []
	length = len(parsedUnmergedColumnData)

	for i in xrange (startingRowIndex, length):
		(columnIndex, rowIndex, cellValue) = parsedUnmergedColumnData[i]
		if cellValue != "":
			result.append(((columnIndex, rowIndex, 1, 1), cellValue))

	for i in xrange (startingRowIndex, length):
		result.append(parsedMergedCellsData[i])

	result = sorted (sorted (result, key=lambda tup: tup[0][0]), key=lambda tup: tup[0][1])
	return result

#chooses records from the merged records list that have the specified column index
def SelectColumnFromMergedSheetData (MergedColumnDataOnSheet, columnIndex):
	result = []
	#cycle prefinish flag
	lastRecord = False
	#if any records were appended to the resulting list
	appendingOccured = False

	i = 0
	while not lastRecord and i < len(MergedColumnDataOnSheet):
		if MergedColumnDataOnSheet[i][0] == columnIndex:
			result.append(MergedColumnDataOnSheet)
			appendingOccured = True
		else:
			if appendingOccured:
				lastRecord = True
		i += 1

	return result



(isFound, row, column) = findColumn(sheet, columnName)

if isFound:
	columnData = readUnmergedColumnData(sheet, row, column)

#print readUnmergedColumnData(sheet, row, column)
#print parseMergedCells(sheet)

r = MergeColumnDataOnSheet(row, readUnmergedColumnData(sheet, row, column), parseMergedCells(sheet))

for ((xLow, xHigh, yLow, yHigh), s) in r:
	print yLow, xLow, xHigh, yHigh, s