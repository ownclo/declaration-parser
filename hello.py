import json
import os
import types

from xlrd import open_workbook


def isMatchedName2(nameAsString, value):
		return type(value) is types.UnicodeType and nameAsString in value

def findCellByName (mergedData, columnName):
	for record in mergedData:
		if isMatchedName2(columnName, record[1]):
			return record

def parseUnmergedCells(sheet):
	result = []

	for i in xrange(sheet.nrows):
		for j in xrange(sheet.ncols):
			cellValue = sheet.cell_value(i, j)
			if cellValue != "":
				result.append(((j, i, j+1, i+1), cellValue))

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

def getLeftTopCornerCoordinates (record):
	return (record[0][0], record[0][1])

def mergedAlready (record, list):
	return getLeftTopCornerCoordinates (record) in list

#merges all the records from merged and non-merged cell lists
def mergeColumnDataOnSheet(parsedUnmergedColumnData, parsedMergedCellsData):
	result = []
	LeftCornerCoordinates = set()

	for i in xrange (len(parsedMergedCellsData)):
		result.append(parsedMergedCellsData[i])
		LeftCornerCoordinates.add(getLeftTopCornerCoordinates(parsedMergedCellsData[i]))

	result.extend(filter(lambda datum: datum[1] != "" and not mergedAlready(datum, LeftCornerCoordinates), parsedUnmergedColumnData))

	result = sorted (sorted (result, key=lambda tup: tup[0][0]), key=lambda tup: tup[0][1])
	return result

def selectIntersectionOf (descRange, data):
	return filter (lambda datum: inRange(xRange(datum[0]), descRange), data)

def fillSchemeDataByRange(inputData, pathToColumn, column, (xLow, xHigh)):
	if pathToColumn == []:
		return (xLow, xHigh)

	head, tail = pathToColumn[0], pathToColumn[1:]

	subData = selectIntersectionOf((xLow, xHigh), inputData)
	header = findCellByName(subData, head)
	return fillSchemeDataByRange(subData, tail, column, xRange(header[0]))


def fillSchemeDataByScheme(inputData, scheme, tableColumnCount):
	result = {}
	columns = scheme["columns"]
	initialRange = (0, tableColumnCount - 1)

	for column in columns:
		toName = column["toName"]
		Range = fillSchemeDataByRange(inputData, column["name"], column, initialRange)
		result[toName] = Range, typeMap()[column["type"]]

	return result

def isMatchedName(schemeElem, value):
		return type(value) is types.UnicodeType and matchedName(schemeElem, value)

def matchedName(schemeElem, value):
	aliases = schemeElem["name"] + schemeElem["aliases"]
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
			 "int"    : int,
			 "float"  : float}
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
	columnDict = fillSchemeDataByScheme(data, schema, sheet.ncols)

	columnDesc = columnDict['id']

	idColumn = selectColumn(columnDesc, data)
	firstIdLoc = idColumn[0][0]
	firstRow = selectRow(firstIdLoc, data)

	l = selectColumn(columnDict["ownershipObjectType"], firstRow)

	for (arange, elem) in l:
		print arange, elem

if __name__ == "__main__":
	main()