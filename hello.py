import json
import types
import os

from xlrd import open_workbook

def parseUnmergedCells(sheet):
	result = []

	for i in xrange(sheet.nrows):
		for j in xrange(sheet.ncols):
			cellValue = sheet.cell_value(i, j)
			if cellValue != "":
				result.append((j, i, cellValue))

	return result

#lists all the merged cells in the xls
def parseMergedCells(sheet):
	xLow = yLow = yHigh = 0
	result = []

	for crange in sheet.merged_cells:
		yLow, yHigh, xLow, xHigh = crange
		result.append (((xLow, yLow, xHigh, yHigh), sheet.cell_value(yLow, xLow)))

	result = sorted (sorted (result, key=lambda tup: tup[0][0]), key=lambda tup: tup[0][1])

	return result

#merges all the records from merged and non-merged cell lists
def mergeColumnDataOnSheet(parsedUnmergedColumnData, parsedMergedCellsData):
	result = []
	length = len(parsedUnmergedColumnData)

	for i in xrange(length):
		(columnIndex, rowIndex, cellValue) = parsedUnmergedColumnData[i]
		if cellValue != "":
			result.append(((columnIndex, rowIndex, columnIndex + 1, rowIndex + 1), cellValue))

	for i in xrange (len(parsedMergedCellsData)):
		result.append(parsedMergedCellsData[i])

	result = sorted (sorted (result, key=lambda tup: tup[0][0]), key=lambda tup: tup[0][1])
	return result

def fillSchemeDataByScheme(inputData, scheme):
	result = {}
	columns = scheme["columns"]

	for column in columns:
		toName = column["toName"]

		for inp in inputData:
			value = inp[1]
			xLow, xHigh = xRange(inp[0])

			if isMatchedName(column, value) and toName not in result:
				result[toName] = (xLow, xHigh), typeMap()[column["type"]]

	return result

def isMatchedName(schemeElem, value):
	return type(value) is types.UnicodeType and matchedName(schemeElem, value)

def matchedName(schemeElem, value):
	aliases = [schemeElem["name"]] + schemeElem["aliases"]
	return any([ name in value for name in aliases])

#chooses records from the merged records list that have the specified column index
def selectColumn(columnDesc, data):
	return map(lambda datum: convertType(columnDesc, datum),
			filter(lambda datum: isDatumInColumn(columnDesc, datum), data))

def selectRow(rowLoc, data):
	return filter(lambda datum: inRange(yRange(datum[0]), yRange(rowLoc)), data)

def isDatumInColumn(columnDesc, datum):
	dataRange = datum[0]
	dataValue = datum[1]
	descRange = columnDesc[0]
	descType  = columnDesc[1]

	return inRange(xRange(dataRange), descRange) and typeMatches(dataValue, descType)

def xRange((xLow, _yLow, xHigh, _yHigh)):
	return (xLow, xHigh)

def yRange((_xLow, yLow, _xHigh, yHigh)):
	return (yLow, yHigh)

def inRange((datumLow, datumHigh), (rangeLow, rangeHigh)):
	return datumLow >= rangeLow and datumHigh <= rangeHigh

def typeMatches(dataValue, descType):
	dataType = type(dataValue)
	if dataType is descType:
		return True

	if dataType is float and descType is int:
		return dataValue.is_integer()

def convertType(columnDesc, datum):
	descType = columnDesc[1]
	coord, datumValue = datum
	if descType is int:
		return coord, int(datumValue)
	else: return datum

def typeMap():
	TYPES = {"string" : types.UnicodeType,
			 "int"    : int}
	return TYPES

def main():
	dirname = "samples"
	fileName = "svedeniya-o-dohodah-sotrudnikov-territorialnyih-organov-roskomnadzora-2015.xls"
	schemaName = "schema.json"

	schema = json.load(open(schemaName))
	fname = os.path.join(dirname, fileName)
	book = open_workbook(fname, formatting_info=True)
	sheet = book.sheet_by_index(0)

	unmerged = parseUnmergedCells(sheet)
	merged = parseMergedCells(sheet)

	data = mergeColumnDataOnSheet(unmerged, merged)
	columnDict = fillSchemeDataByScheme(data, schema)

	columnDesc = columnDict['id']

	idColumn = selectColumn(columnDesc, data)
	firstIdLoc = idColumn[2][0]
	firstRow = selectRow(firstIdLoc, data)

	for (arange, elem) in firstRow:
		print arange, elem


if __name__ == "__main__":
	main()