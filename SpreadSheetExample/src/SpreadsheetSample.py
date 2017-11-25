#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.table import BorderLine
from com.sun.star.table import TableBorder
from com.sun.star.awt import FontWeight
from com.sun.star.text import ControlCharacter
from com.sun.star.sheet.GeneralFunction import AVERAGE
from com.sun.star.lang import Locale
from com.sun.star.util import NumberFormat
from com.sun.star.sheet.FillDirection import TO_RIGHT, TO_LEFT, TO_TOP, TO_BOTTOM
from com.sun.star.sheet.FillMode import LINEAR, DATE, AUTO, GROWTH
from com.sun.star.sheet.FillDateMode import FILL_DATE_DAY, FILL_DATE_MONTH
from com.sun.star.sheet.TableOperationMode import BOTH, COLUMN
from com.sun.star.sheet import CellFlags
from com.sun.star.table import CellRangeAddress
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	doCellSamples(doc)
	doCellRangeSamples(doc)
	doCellRangesSamples(doc)
	doCellCursorSamples()
	doFormattingSamples()
	doDocumentSamples()
	doDatabaseSamples()
	doDataPilotSamples()
	doNamedRangesSamples()
	doFunctionAccessSamples()
	doApplicationSettingsSamples()
	
def	doCellSamples(doc):
	print("\n*** Samples for service sheet.SheetCell ***\n")
	sheets = doc.getSheets()
	sheet = sheets[0]	
	prepareRange(sheet, "A1:C7", "Cells and Cell Ranges")
	# --- Get cell B3 by position - (row, column) ---
	cell = sheet[2, 1]
	# --- Insert two text paragraphs into the cell. ---
	textcursor = cell.createTextCursor()
	cell.insertString(textcursor, "Text in first line.", False)
	cell.insertControlCharacter(textcursor, ControlCharacter.PARAGRAPH_BREAK, False)
	cell.insertString(textcursor, "And a ", False)
	# create a hyperlink
	hyperlink = doc.createInstance("com.sun.star.text.TextField.URL")
	hyperlink.setPropertyValue("URL", "https://p--q.blogspot.jp/")  # setPropertyValuesは使えない。
	hyperlink.setPropertyValue("Representation", "hyperlink")
	# ... and insert
	cell.insertTextContent(textcursor, hyperlink, False)
	# --- Query the separate paragraphs. ---
	paraenum = cell.createEnumeration()
	# Go through the paragraphs
	for portion in paraenum:
		portionenum = portion.createEnumeration()
		txt = ""
		# Go through all text portions of a paragraph and construct string.
		for rng in portionenum:
			txt += rng.getString()
		print("Paragraph text: {}".format(txt))
	# --- Change cell properties. ---
	# from styles.CharacterProperties
	cell.setPropertyValues(("CharColor", "CharHeight"), (0x003399, 20.0))
	# from styles.ParagraphProperties
	cell.setPropertyValue("ParaLeftMargin", 500)
	# from table.CellProperties
	cell.setPropertyValues(("IsCellBackgroundTransparent", "CellBackColor"), (False, 0x99CCFF))
	# --- Get cell address. ---
	address = cell.getCellAddress()
	txt = "Address of this cell:  Column={}".format(address.Column)
	txt += ";  Row={}".format(address.Row)
	txt += ";  Sheet={}".format(address.Sheet)
	print(txt)
	# --- Insert an annotation ---
	annotations = sheet.getAnnotations()
	annotations.insertNew(address, "This is an annotation")
	annotation = cell.getAnnotation()
	annotation.setIsVisible(True)
# ** All samples regarding the service com.sun.star.sheet.SheetCellRange. *
def	doCellRangeSamples(doc):
	print("\n*** Samples for service sheet.SheetCellRange ***\n")
	sheets = doc.getSheets()
	sheet = sheets[0]	
	# Preparation
	sheet["B5"].setFormula("First cell")
	sheet["B6"].setFormula("Second cell")
	# Get cell range B5:B6 by position - (column, row, column, row)
	cellrng = sheet[4:6, 1]
	# --- Change cell range properties. ---
	# from com.sun.star.styles.CharacterProperties
	cellrng.setPropertyValues(("CharColor", "CharHeight"), (0x003399, 20.0))
	# from com.sun.star.styles.ParagraphProperties
	cellrng.setPropertyValue("ParaLeftMargin", 500)
	# from com.sun.star.table.CellProperties
	cellrng.setPropertyValues(("IsCellBackgroundTransparent", "CellBackColor"), (False, 0x99CCFF))
	# --- Replace text in all cells. ---
	replacedesc = cellrng.createReplaceDescriptor()
	replacedesc.setSearchString("cell") 
	replacedesc.setReplaceString("text")
	# property SearchWords searches for whole cells!
	replacedesc.setPropertyValue("SearchWords", False)
	c = cellrng.replaceAll(replacedesc)
	print("Search text replaced {} times.".format(c))
	# --- Merge cells. ---
	cellrng = sheet["F3:G6"]
	prepareRange(sheet, "E1:H7", "XMergeable")
	cellrng.merge(True)
	# --- Change indentation. ---
	# does not work (bug in XIndent implementation)
# 	prepareRange(sheet, "I20:I23", "XIndent" )
# 	sheet["I21"].setValue(1)
# 	sheet["I22"].setValue(1)
# 	sheet["I23"].setValue(1)
# 	cellrange = sheet["I21:I22"]
# 	cellrange.incrementIndent()
# 	cellrange = sheet["I21:I23"]
# 	cellrange.incrementIndent()
	# --- Column properties. ---
	cellrange = sheet["B1"]
	columns = cellrange.getColumns()
	column = columns[0]
	column.setPropertyValue("Width", 6000)
	print("The name of the wide column is {}.".format(column.getName()))
	# --- Cell range data ---
	prepareRange(sheet, "A9:C30", "XCellRangeData")
	cellrange = sheet["A10:C30"]
	vals = 	("Name",   "Fruit",	"Quantity"),\
			("Alice",  "Apples",  3.0),\
			("Alice",  "Oranges",7.0 ),\
			("Bob",	"Apples",  3.0),\
			("Alice",  "Apples",  9.0),\
			("Bob",	"Apples",  5.0),\
			("Bob",	"Oranges", 6.0),\
			("Alice",  "Oranges", 3.0),\
			("Alice",  "Apples",  8.0),\
			("Alice",  "Oranges", 1.0),\
			("Bob",	"Oranges", 2.0),\
			("Bob",	"Oranges", 7.0),\
			("Bob",	"Apples",  1.0),\
			("Alice",  "Apples",  8.0),\
			("Alice",  "Oranges", 8.0),\
			("Alice",  "Apples",  7.0),\
			("Bob",	"Apples",  1.0),\
			("Bob",	"Oranges", 9.0),\
			("Bob",	"Oranges", 3.0),\
			("Alice",  "Oranges", 4.0),\
			("Alice",  "Apples",  9.0)	
	cellrange.setDataArray(vals)
	# --- Get cell range address. ---
	rangeaddress = cellrange.getRangeAddress()
	print("Address of this range:  Sheet={}".format(rangeaddress.Sheet))
	print("Start column={};  Start row={}".format(rangeaddress.StartColumn, rangeaddress.StartRow))
	print("End column={};  End row={}".format(rangeaddress.EndColumn, rangeaddress.EndRow))
	# --- Sheet operation. ---
	# uses the range filled with XCellRangeData
	result = cellrange.computeFunction(AVERAGE)
	print("Average value of the data table A10:C30: {}".format(result))
	# --- Fill series ---
	# Prepare the example
	formattypes = doc.getNumberFormats()
	dateformat = formattypes.getStandardFormat(NumberFormat.DATE, Locale())
	sheet["E10"].setValue(1)
	sheet["E11"].setValue(4)
	sheet["E12"].setFormula("2002-1-30")  # 年-月-日 または 月/日/年 にしないといけないらしい。
	sheet["E12"].setPropertyValue("NumberFormat", dateformat)
	sheet["I13"].setFormula("Text 10")
	sheet["E14"].setFormula("Jan")
	sheet["K14"].setValue(10)
	sheet["E16"].setValue(1)
	sheet["F16"].setValue(2)
	sheet["E17"].setFormula("2002-2-28") 
	sheet["E17"].setPropertyValue("NumberFormat", dateformat)
	sheet["F17"].setFormula("2002-1-28") 
	sheet["F17"].setPropertyValue("NumberFormat", dateformat)
	sheet["E18"].setValue(6)
	sheet["F18"].setValue(4)
	# Fill 2 rows linear with end value -> 2nd series is not filled completely
	sheet["E10:I11"].fillSeries(TO_RIGHT, LINEAR, FILL_DATE_DAY, 2, 9)
	# Add months to a date
	sheet["E12:I12"].fillSeries(TO_RIGHT, DATE, FILL_DATE_MONTH, 1, 0x7FFFFFFF)
	# Fill right to left with a text containing a value
	sheet["E13:I13"].fillSeries(TO_LEFT, LINEAR, FILL_DATE_DAY, 10, 0x7FFFFFFF)
	# Fill with an user defined list]
	sheet["E14:I14"].fillSeries(TO_RIGHT, AUTO, FILL_DATE_DAY, 10, 0x7FFFFFFF)
	# Fill bottom to top with a geometric series
	sheet["K10:K14"].fillSeries(TO_TOP, GROWTH, FILL_DATE_DAY, 2, 0x7FFFFFFF)
	# Auto fill
	sheet["E16:K18"].fillAuto(TO_RIGHT, 2)
	# Fill series copies cell formats -> draw border here
	prepareRange(sheet, "E9:K18", "XCellSeries")
	# --- Array formulas ---
	prepareRange(sheet, "E20:G23", "XArrayFormulaRange")
	# Insert a 3x3 unit matrix.
	arrayformula = sheet["E21:G23"]
	arrayformula.setArrayFormula("=A10:C12")
	print("Array formula is: {}".format(arrayformula.getArrayFormula()))
	#  --- Multiple operations ---
	sheet["E26"].setFormula("=E27^F26")
	sheet["E27"].setValue(1)
	sheet["F26"].setValue(1)
	sheet["E27:E31"].fillAuto(TO_BOTTOM, 1)
	sheet["F26:J26"].fillAuto(TO_RIGHT, 1)
	sheet["F33"].setFormula("=SIN(E33)")
	sheet["G33"].setFormula("=COS(E33)")
	sheet["H33"].setFormula("=TAN(E33)")
	sheet["E34"].setValue(0)
	sheet["E35"].setValue(0.2)
	sheet["E34:E38"].fillAuto(TO_BOTTOM, 2)
	prepareRange(sheet, "E25:J38", "XMultipleOperation")
	formularange = sheet["E26"].getRangeAddress()
	colcell = sheet["E27"].getCellAddress()
	rowcell = sheet["F26"].getCellAddress()
	sheet["E26:J31"].setTableOperation(formularange, BOTH, colcell, rowcell)
	formularange = sheet["F33:H33"].getRangeAddress()
	colcell = sheet["E33"].getCellAddress()
	# Row cell not needed
	sheet["E34:H38"].setTableOperation(formularange, COLUMN, colcell, rowcell)
	# --- Cell Ranges Query ---
	cellranges = sheet["A10:C30"].queryContentCells(CellFlags.STRING)
	print("Cells in A10:C30 containing text: {}".format(cellranges.getRangeAddressesAsString()))
def	doCellRangesSamples(doc):
	print("\n*** Samples for cell range collections ***\n")
	# Create a new cell range container
	rangecont = doc.createInstance("com.sun.star.sheet.SheetCellRanges")
	# --- Insert ranges ---
	address = CellRangeAddress(Sheet=0, StartColumn=0, StartRow=0, EndColumn=0, EndRow=0)
	rangecont.addRangeAddress(address, False)
# 	print("Inserting {} {} merge,{} resulting list: {}".format())
	
	conv = doc.createInstance("com.sun.star.table.CellAddressConversion")
	tcu.wtree(doc.createInstance("com.sun.star.table.CellAddressConversion"))
	
# 		insertRange( xRangeCont, 0, 0, 0, 0, 0, false );	// A1:A1
# 		insertRange( xRangeCont, 0, 0, 1, 0, 2, true );	 // A2:A3
# 		insertRange( xRangeCont, 0, 1, 0, 1, 2, false );	// B1:B3	
	


def getCellRangeAddressString():
	
	
	return "{}:{}".format()

def getCellAddressString(column, row):
	txt = ""
	if column>25:
		txt += "A{}".format(column/26-1)
	txt += "A{}".format(column%26)
	txt += row + 1
	return txt
def	doCellCursorSamples():
	pass
def	doFormattingSamples():
	pass
def	doDocumentSamples():
	pass
def	doDatabaseSamples():
	pass
def	doDataPilotSamples():
	pass
def	doNamedRangesSamples():
	pass
def	doFunctionAccessSamples():
	pass
def	doApplicationSettingsSamples():
	pass
# ** Draws a colored border around the range and writes the headline in the first cell.
def prepareRange(sheet, rng, headline):
	# draw border
	cellrange = sheet[rng]
	borderline = BorderLine(Color=0x99CCFF, InnerLineWidth=0, LineDistance=0, OuterLineWidth=100)
	tableborder = TableBorder(TopLine=borderline, BottomLine=borderline, LeftLine=borderline, RightLine=borderline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
	cellrange.setPropertyValue("TableBorder", tableborder)  # Pythonのオートメーションで実行すると、以後LibreOfficeを終了してJavaの例を実行しないと2列目以降のすべてのセルに上下の枠線が表示されてしまう。
	# draw headline
	addr = cellrange.getRangeAddress()
	sheet[addr.StartRow, addr.StartColumn:addr.EndColumn+1].setPropertyValue("CellBackColor", 0x99CCFF)
	# write headline
	cell = cellrange[0, 0]
	cell.setFormula(headline)
	cell.setPropertyValues(("CharColor", "CharWeight"), (0x003399, FontWeight.BOLD))


g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
	import sys
	from com.sun.star.beans import PropertyValue
	from com.sun.star.script.provider import XScriptContext  
	def connectOffice(func):  # funcの前後でOffice接続の処理
		@wraps(func)
		def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
			try:
				ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
			except:
				print("Could not establish a connection with a running office.", file=sys.stderr)
				sys.exit()
			print("Connected to a running office ...")
			smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
			print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
			return func(ctx, smgr)  # 引数の関数の実行。
		def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
			cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
			node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
			ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
			return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
		return wrapper
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
		class ScriptContext(unohelper.Base, XScriptContext):
			def __init__(self, ctx):
				self.ctx = ctx
			def getComponentContext(self):
				return self.ctx
			def getDesktop(self):
				return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
			def getDocument(self):
				return self.getDesktop().getCurrentComponent()
		return ScriptContext(ctx)  
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
# 	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
	if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
	flg = True
	while flg:
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		if doc is not None:
			flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
	macro()