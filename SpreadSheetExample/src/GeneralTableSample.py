#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.awt import Rectangle
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	sheets = doc.getSheets()
	sheet = sheets[0]	
	# *** Access and modify a VALUE CELL ***
	cell = sheet[0, 0]
	# Set cell value.
	cell.setValue(1234)
	# Get cell value.
	val = cell.getValue()*2
	sheet[1, 0].setValue(val)
	# *** Create a FORMULA CELL and query error type ***
	cell = sheet[2, 0]
	# Set formula string.
	cell.setFormula("=1/0")
	# Get error type.
	flag = (cell.getError()==0)  # cell.getError() return 532
	# Get formula string.
	txt = "The formula {} is ".format(cell.getFormula())
	txt += "valid." if flag else "erroneous."
	# *** Insert a TEXT CELL using the XText interface ***
	cell = sheet[3,0]
	textcursor = cell.createTextCursor()
	cell.insertString(textcursor, txt, False)
	# *** Change cell properties ***
	color = 0x00FF00 if flag else 0xFF4040
	cell.setPropertyValue("CellBackColor", color)
	# *** Accessing a CELL RANGE ***
	# Accessing a cell range over its position.
	cellrange = sheet[:2,2:4]
	# Change properties of the range.
	cellrange.setPropertyValue("CellBackColor", 0x8080FF)
	# Accessing a cell range over its name.
	cellrange = sheet["C4:D5"] 
	# Change properties of the range.
	cellrange.setPropertyValue("CellBackColor", 0xFFFF80)
	# *** Using the CELL CURSOR to add some data below of the filled area ***
	cell = sheet["A1"]
	cursor = sheet.createCursorByRange(cell)
	# Move to the last filled cell.
	cursor.gotoEnd()
	# Move one row down.
	cursor.gotoOffset(0, 1) # (ColumnOffset, RowOffset)
	cursor[0, 0].setFormula("Beyond of the last filled cell.")
	# *** Modifying COLUMNS and ROWS ***
	columns = sheet.getColumns()
	rows = sheet.getRows()
	# Get column C by index (interface XIndexAccess).
	column = columns[2]
	column.setPropertyValue("Width", 5000)
	# Get the name of the column.
	txt = "The name of this column is {}.".format(column.getName())
	sheet[2,2].setFormula(txt)
	# Get column D by name (interface XNameAccess).
	column = columns["D"]
	column.setPropertyValue("IsVisible", False)
	# Get row 7 by index (interface XIndexAccess)
	row = rows[6]
	row.setPropertyValue("Height", 5000)
	sheet[6, 2].setFormula("What a big cell.")
	# Create a cell series with the values 1 ... 7.
	sheet[8:15, 0].setDataArray([[i] for i in range(1, 8)])
	# Insert a row between 1 and 2
	rows.insertByIndex(9, 1)
	# Delete the rows with the values 3 and 4.
	rows.removeByIndex(11, 2)
	# *** Inserting CHARTS ***
	charts = sheet.getCharts()
	# The chart will base on the last cell series, initializing all values.
	chartname = "newChart"
	rectangle = Rectangle(X=10000, Y=3000, Width=5000, Height = 5000)
	rng = sheet["A9:A14"].getRangeAddress()
	# Create the chart.
	charts.addNewByName(chartname, rectangle, (rng,), False, False)
	# Get the chart by name.
	chart = charts.getByName(chartname)
	# Query the state of row and column headers.
	txt = "Chart has column headers: "
	txt += "yes" if chart.getHasColumnHeaders() else "no"
	sheet[8, 2].setFormula(txt)
	txt = "Chart has row headers: "
	txt += "yes" if chart.getHasRowHeaders() else "no"
	sheet[9, 2].setFormula(txt)
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