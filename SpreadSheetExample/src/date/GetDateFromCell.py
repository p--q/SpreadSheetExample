#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.util import NumberFormat  # 定数
global XSCRIPTCONTEXT
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	createFormatKey = formatkeyCreator(doc)
	sheets = doc.getSheets()  # ドキュメントのシートコレクションを取得。。
	sheet = sheets[0]  # シートコレクションのインデックス0のシートを取得。
	sheet.clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA+CellFlags.HARDATTR+CellFlags.STYLES)  # セルの内容を削除。
	methods = "getFormula", "getValue", "getString", "getType"
	for i, method in enumerate(methods, start=2):
		sheet[0, i].setString("{}()".format(method))	
	datestring = "2017-10-25"  # 2017-10-25 or 10/25/2017
	sheet["A2"].setString('sheet["B1"].setFormula("{}")'.format(datestring))
	cell = sheet["B2"]
	cell.setFormula(datestring)  # 式で日付を入力する。
	for i, method in enumerate(methods, start=2):
		sheet[1, i].setString(str(getattr(cell, method)()))
	formatstring = "YYYY-MM-DD"
	sheet["A3"].setString(formatstring)
	sheet["B3"].setFormula(datestring)
	sheet["B3"].setPropertyValue("NumberFormat", createFormatKey(formatstring))  # セルの書式を設定。	
	cell = sheet["B3"]
	cell.setFormula(datestring)  # 式で日付を入力する。
	for i, method in enumerate(methods, start=2):
		sheet[2, i].setString(str(getattr(cell, method)()))	
	formatstring = "GE.MM.DD"
	sheet["A4"].setString(formatstring)
	sheet["B4"].setFormula(datestring)
	sheet["B4"].setPropertyValue("NumberFormat", createFormatKey(formatstring))  # セルの書式を設定。	
	cell = sheet["B4"]
	cell.setFormula(datestring)  # 式で日付を入力する。
	for i, method in enumerate(methods, start=2):
		sheet[3, i].setString(str(getattr(cell, method)()))		
	sheet["A5"].setString("Standard Date Format")
	sheet["B5"].setFormula(datestring)
	numberformats = doc.getNumberFormats()
	formatkey =numberformats.getStandardFormat(NumberFormat.DATE, Locale())
	sheet["B5"].setPropertyValue("NumberFormat", formatkey)  # セルの書式を設定。	
	cell = sheet["B5"]
	cell.setFormula(datestring)  # 式で日付を入力する。
	for i, method in enumerate(methods, start=2):
		sheet[4, i].setString(str(getattr(cell, method)()))	
	sheet["A:F"].getColumns().setPropertyValue("OptimalWidth", True)  # 列幅を最適化する。
def formatkeyCreator(doc):  # ドキュメントを引数にする。
	def createFormatKey(formatstring):  # formatstringの書式はLocalによって異なる。	
		numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
		locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。。	
		formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。	
		if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
			formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。	
		return formatkey
	return createFormatKey
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
