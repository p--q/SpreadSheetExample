#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from itertools import zip_longest
# from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.sheet import CellFlags as cf # 定数
def macro(documentevent=None):  # 引数は文書のイベント駆動用。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得
	sheets = doc.getSheets()
	sheet = sheets[0]
	conditionalformat = sheet["C35"].getPropertyValue("ConditionalFormat")
	outputs = []  # 出力行のリスト。
# 	for conditionalformat in conditionalformats.getConditionalFormats():
# 		propertysetinfo = conditionalformat.getPropertySetInfo()
# 		properties = propertysetinfo.getProperties()
# 		expandPropertyValueStructs(outputs, properties, 0)  # PropertyValue Structを展開する。
# 		header = "ConditionalFormat",
# 		datarows = [header]
# 		datarows.extend(outputs)
# 		sheet.getRows().insertByIndex(r+1, len(datarows)+1)  # 行インデックスrに挿入するデータの行数の行を挿入する。
# 		rowsToSheet(sheet[r+1, 0], datarows)	
	
def expandPropertyValueStructs(outputs, structs, h):  # outputs: 出力行のリスト、structs: PropertyValuesのタプル、h: 出力階層。
	flg = True  # 先頭行のフラグ。
	for struct in structs:
		ind = [""]*h
		if flg and h>0:  # 出力階層1以上で先頭行のとき。
			ind[-1] = "Value"
			flg = False
		name = "Name", struct.Name
		ind.extend(name)  # 戻り値はない。
		outputs.append(ind)
		if hasattr(struct, "Value"):
			v = struct.Value
			if isinstance(v, tuple):  # Valueがタプルの時は入れ子になっているので再帰呼出し。
				expandPropertyValueStructs(outputs, v, h+1)
			else:
				if isinstance(v, bool) or not isinstance(v, int):  # vがint以外の時。bool型はintに含まれるので注意。
					v = str(v)  # setDataArrayではint以外はStringしか代入できないのでStringに変換する。
				ind = [""]*h
				value = "Value", v
				ind.extend(value)  # 戻り値はない。
				outputs.append(ind)
def rowsToSheet(cellrange, datarows):  # 引数のセル範囲を左上端にして一括書き込みして列幅を最適化する。datarowsはタプルのタプル。
	datarows = tuple(zip(*zip_longest(*datarows, fillvalue="")))  # 一番長い行の長さに合わせて空文字を代入。
	sheet = cellrange.getSpreadsheet()  # セル範囲のあるシートを取得。
	cellcursor = sheet.createCursorByRange(cellrange)  # セル範囲のセルカーサーを取得。
	cellcursor.collapseToSize(len(datarows[0]), len(datarows))  # (列、行)で指定。セルカーサーの範囲をdatarowsに合せる。
	cellcursor.setDataArray(datarows)  # セルカーサーにdatarowsを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。  
def getNewSheet(doc, sheetname):  # docに名前sheetnameのシートを返す。sheetnameがすでにあれば連番名を使う。
	cellflags = cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES
	sheets = doc.getSheets()  # シートコレクションを取得。
	c = 1  # 連番名の最初の番号。
	newname = sheetname
	while newname in sheets: # 同名のシートがあるとき。sheets[newname]ではFalseのときKeyErrorになる。
		if not sheets[newname].queryContentCells(cellflags):  # シートが未使用のとき
			return sheets[newname]  # 未使用の同名シートを返す。
		newname = "{}{}".format(sheetname, c)  # 連番名を作成。
		c += 1 
	sheets.insertNewByName(newname, len(sheets))   # 新しいシートを挿入。同名のシートがあるとRuntimeExceptionがでる。
	if "Sheet1" in sheets:  # デフォルトシートがあるとき。
		if not sheets["Sheet1"].queryContentCells(cellflags):  # シートが未使用のとき
			del sheets["Sheet1"]  # シートを削除する。
	return sheets[newname]
def createConfigReader(ctx, smgr):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	def configReader(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return configReader
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
