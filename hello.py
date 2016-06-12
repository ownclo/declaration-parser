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

#returns TRUE and the coordinates of the cell that contains the specified text if the search succeeded.
#if the cell was not found, then FALSE and the size of the entire table is returned.
def findCellWithText(sheet, TextToFind):
	for i in xrange(0, sheet.nrows):
		for j in xrange(0, sheet.ncols):
			cell = sheet.cell_value(i, j)
			if TextToFind in cell:
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
			result.append(((columnIndex, rowIndex, columnIndex + 1, rowIndex + 1), cellValue))

	for i in xrange (len(parsedMergedCellsData)):
		result.append(parsedMergedCellsData[i])

	result = sorted (sorted (result, key=lambda tup: tup[0][0]), key=lambda tup: tup[0][1])
	return result

def FillSchemeDataByScheme(InputData, Scheme):
	result = []
	SchemeElement = []
	SchemeLength = len(Scheme)
	InputDataLength = len(InputData)

	for i in xrange (SchemeLength):
		result.append([])
		SchemeElement = Scheme["columns"][i]
		for j in xrange (InputDataLength):
			CellValue = InputData[j][1]
			if SchemeElement["name"] in CellValue:
				result[i].append(InputData[j])
			else:
				for k in xrange(len(SchemeElement["aliases"])):
					if SchemeElement["aliases"][k] in CellValue:
						result[i].append(InputData[j])
	if len(result) == SchemeElement:
		return (True, result)
	else:
		return (False, [])

#remerges the list in multiple logical columns according to the initial data;
#builds a table containing logical column coordinates
#def BuildLogicalColumnsFromMergedDataset (InputData, ParsedScheme):
#	result = []
#	InputLength = len (InputData)
#	SchemeLength = len(ParsedScheme)
#
#	for i in xrange(SchemeLength):
#		CellValue = ParsedScheme[i][4]
#		result.append(())
#		(IsFound, x, y) = findCellWithText(CellValue)
#		if IsFound:
#			j = 0
#			Found = False
#			while j < InputLength and not Found:
#				if CellValue == InputData[j][4]:
#					result[i].append()
#					Found = True
#					j += 1
#
#
#	return result

#chooses records from the merged records list that have the specified column index
def SelectColumnFromMergedSheetData (MergedColumnDataOnSheet, columnIndex):
	return filter (lambda x: x[0] == columnIndex, MergedColumnDataOnSheet)



(isFound, row, column) = findCellWithText(sheet, columnName)

if isFound:
	columnData = readUnmergedColumnData(sheet, row, column)

#print readUnmergedColumnData(sheet, row, column)
#print parseMergedCells(sheet)

MergedCells = MergeColumnDataOnSheet(row, readUnmergedColumnData(sheet, row, column), parseMergedCells(sheet))

r = FillSchemeDataByScheme(MergedCells, schema)

for ((xLow, yLow, xHigh, yHigh), s) in r:
	print xLow, yLow, xHigh, yHigh, s