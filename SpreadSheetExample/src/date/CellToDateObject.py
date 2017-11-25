#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import date
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	createFormatKey = formatkeyCreator(doc)
	sheets = doc.getSheets()  # ドキュメントのシートコレクションを取得。。
	sheet = sheets[0]  # シートコレクションのインデックス0のシートを取得。
	sheet.clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA+CellFlags.HARDATTR+CellFlags.STYLES)  # セルの内容を削除。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)
	todayvalue = functionaccess.callFunction("TODAY", ())  # スプレッドシート関数で今日の日付のシリアル値を取得。
	sheet["A1"].setValue(todayvalue)  # セルに日付時間シリアル値を入力。
	sheet["A1"].setPropertyValue("NumberFormat", createFormatKey("GE.MM.DD"))  # セルの書式を設定。
	datetimevalue = sheet["A1"].getValue()  # セルから日付時間シリアル値を取得。
	year = int(functionaccess.callFunction("YEAR", (datetimevalue,)))  # シリアル値から年を取得。floatで返ってくるので整数にする。
	month = int(functionaccess.callFunction("MONTH", (datetimevalue,)))  # シリアル値から月を取得。floatで返ってくるので整数にする。
	day = int(functionaccess.callFunction("DAY", (datetimevalue,)))  # シリアル値から日を取得。floatで返ってくるので整数にする。
	celldate = date(year, month, day)  # Pythonのdateオブジェクトにする。
	sheet["A2"].setString(celldate.isoformat())  # 文字列でセルに書き出す。
	

	
	
# 	headers = "Sheet Function", "Return Type", "Return Value", "Format or Formula", "Formatted Value"
# 	for i, header in enumerate(headers):
# 		sheet[0, i].setString(header)
# 	today = functionaccess.callFunction("TODAY", ())  # 引数のないスプレッドシート関数。
# 	txts = "TODAY()", type(today).__name__, str(today), "YYYY-MM-DD"
# 	for i, t in enumerate(txts):
# 		sheet[1, i].setString(t)
# 	cell = castToXCellRange(sheet[1, i+1])	 # 次にcallFunction()の引数にいれるために、com.sun.star.table.XCellRange型でセルを取得する。
# 	cell.setValue(today)
# 	cell.setPropertyValue("NumberFormat", createFormatKey(t))  # セルの書式を設定。
# 	year = functionaccess.callFunction("YEAR", (cell,))  # 引数のあるスプレッドシート関数。タプルの入れ子で返ってくる。
# 	txts = 'year = YEAR("C2")', type(year).__name__, str(year), "year[0][0]"
# 	for i, t in enumerate(txts):
# 		sheet[2, i].setString(t)	
# 	sheet[2, i+1].setValue(year[0][0])
# 	now = functionaccess.callFunction("NOW", ())  # 引数のない関数の例。
# 	txts = "NOW()", type(now).__name__, str(now), "YYYY/M/D H:MM:SS"
# 	for i, t in enumerate(txts):
# 		sheet[3, i].setString(t)	
# 	sheet[3, i+1].setValue(now)	
# 	sheet[3, i+1].setPropertyValue("NumberFormat", createFormatKey(t))  # セルの書式を設定。
# 	sheet["A:E"].getColumns().setPropertyValue("OptimalWidth", True)  # 列幅を最適化する。
def formatkeyCreator(doc):  # ドキュメントを引数にする。
	def createFormatKey(formatstring):  # formatstringの書式はLocalによって異なる。	
		numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
		locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。。	
		formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。	
		if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
			formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。	
		return formatkey
	return createFormatKey
# def castToXCellRange(cell):  # セルをcom.sun.star.table.XCell型からcom.sun.star.table.XCellRange型に変換する。
# 	if cell.supportsService("com.sun.star.sheet.SheetCell"):  # 引数がセルのとき
# 		absolutename = cell.getPropertyValue("AbsoluteName")  # AbsoluteNameを取得。
# 		stringaddress = absolutename.split(".")[-1].replace("$", "")  # シート名を削除後$も削除して、セルの文字列アドレスを取得。
# 		sheet = cell.getSpreadsheet()  # セルのシートを取得。
# 		return sheet[stringaddress]  # com.sun.star.table.XCellRange型のセルを返す。
# 	else:
# 		raise RuntimeError("The argument of castToXCellRange() must be a cell.")
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
if __name__ == "__main__":  # オートメーションで実行するとき
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
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
		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
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
		XSCRIPTCONTEXT = createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
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
		return XSCRIPTCONTEXT
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
	macro()  # マクロの実行。
