#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
	textoutputstream = smgr.createInstanceWithContext("com.sun.star.io.TextOutputStream", ctx)
	tcu.wtree(textoutputstream)

# 	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
# 	doc = XSCRIPTCONTEXT.getDocument()  # Calcドキュメント。
# 	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
# 	tcu.wcompare(desktop, docframe)
	
# 	tcu.wtree(desktop)

# 	filepicker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", ctx)
# 	tcu.wtree(filepicker)
	


# 	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
# 	selection = doc.getCurrentSelection()  # 選択しているオブジェクトを取得。
# 	cellrangeaddress = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
# 	tcu.wtree(cellrangeaddress)
	
# 	doc = XSCRIPTCONTEXT.getDocument()  # Calcドキュメント。
# 	controller = doc.getCurrentController()  # コントローラの取得。
# 	sheet = controller.getActiveSheet()  # アクティブなシートを取得。
# 	cell = sheet["C1"]
# 	address = cell.getCellAddress()
# 	annotations = sheet.getAnnotations()
# 	annotations.insertNew(address, "This is an annotation")
# 	
# 	
# 	annotation = cell.getAnnotation()
# 	annotation.setString("セル注釈")
# 	annotation.setIsVisible(False)
	
# 	tcu.wtree(annotation.getParent())
	
	
# 	annotations = sheet.getAnnotations()
# 	tcu.wtree(annotations)
	
# 	shape = annotation.getAnnotationShape()
# 	tcu.wtree(shape)
	
	
# 	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
# 	
# 	
# 	sheets = doc.getSheets()  # シートコレクション。
# 	sheet = sheets[0]  # 最初のシート。
# 	cell = sheet[0, 0]  # 行インデックス0、列インデックス0、のセル(つまりA1セル)。
# 	cells = sheet[2:5, 3:6]  # 行インデックス2以上5未満、列インデックス3以上6未満(つまりD3:F5と同じ)のセル範囲。 
# 	textcursor = cell.createTextCursor()  # A1セル内のテキストカーサー。
# 	cellcursor = sheet.createCursor()  # 最初のシートすべてのセルカーサー。
# 	rows = sheet.getRows()  # 行アクセスオブジェクト。
# 	row = rows[0]  # 1行目。
# 	columns = sheet.getColumns()  # 行アクセスオブジェクト。
# 	column = columns[0]  # 1行目。
# 	charts = sheet.getCharts()  # チャートコレクション
# 	controller = doc.getCurrentController()  # コントローラーの取得。
	
	
	
	
	
	
	
# 	cellrangeaddress = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
# 	celladdressconversion = doc.createInstance("com.sun.star.table.CellAddressConversion")
# 	cellrangeaddressconversion = doc.createInstance("com.sun.star.table.CellRangeAddressConversion")

# 	filteroptiondialog = smgr.createInstanceWithContext("com.sun.star.comp.GraphicExportDialog", ctx)  # UIコンポーネントをインスタンス化。
# 	filteroptiondialog = smgr.createInstanceWithContext("com.sun.star.comp.Calc.FilterOptionsDialog", ctx)  # UIコンポーネントをインスタンス化。
# 	filteroptiondialog = smgr.createInstanceWithContext("com.sun.star.comp.PDF.PDFDialog", ctx)  # UIコンポーネントをインスタンス化。
# 	tcu.wtree(filteroptiondialog)


	# 各1行しか実行できない。
# 	tcu.wtree(doc)  # Calcドキュメント。
# 	tcu.wtree(sheets)  # シートのコレクション。
# 	tcu.wtree(sheet)  # シート。
# 	tcu.wtree(cell)  # A1セル。
# 	tcu.wtree(cells)  # D3:F5のセル範囲。
# 	tcu.wtree(textcursor)  # A1セル内のテキストカーサー。
# 	tcu.wtree(cellcursor)  # 最初のシートすべてのセルカーサー。
# 	tcu.wtree(rows)  # 行アクセスオブジェクト。
# 	tcu.wtree(row)  # 行
# 	tcu.wtree(columns)  # 列アクセスオブジェクト
# 	tcu.wtree(column)  # 列
# 	tcu.wtree(charts)  # チャートコレクション
# 	tcu.wtree(controller)  # コントローラー
# 	tcu.wtree(celladdressconversion)
# 	tcu.wtree(cellrangeaddressconversion)
# 	tcu.wtree(docframe)


# 	tcu.wtree(sheet["A1"])
# 	print(sheet["A1"])


# 	prop = PropertyValue(Name="Hidden",Value=True)
# 	wdoc = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))
# 	tcu.wcompare(doc, wdoc)
# 	tcu.wcompare(cell, cells)
# 	tcu.wcompare(cell, sheet)
# 	tcu.wcompare(cells, cellcursor)
# 	tcu.wcompare(cellcursor, textcursor)
# 	tcu.wcompare(row, column)
# 	tcu.wcompare(cells, row)
# 	tcu.wcompare(doc, sheet)
# 	wcontroller = wdoc.getCurrentController()  # Writerドキュメントのコントローラーを取得。
# 	tcu.wcompare(controller, wcontroller)
# 	tcu.wcompare(cells, cellrangeaddress)

# 	tcu.wcompare(sheet["A1"], sheet[0, 0])  # どちらもセル。
# 	tcu.wcompare(sheet["A1"], sheet[0:1, 0])  # セルとセル範囲
# 	tcu.wcompare(sheet[0, 0], sheet[0:1, 0])  # セルとセル範囲

# 	print(sheet["A1"])
# 	print(sheet[0, 0])
# 	print(sheet[0:1, 0])



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