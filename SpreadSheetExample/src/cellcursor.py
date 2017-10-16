#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	sheets = doc.getSheets()  # シートコレクション。
	sheet = sheets[0]  # 最初のシート。
	sheet.clearContents(511)  # シートのすべてを削除。
	cursor = sheet.createCursorByRange(sheet["B5:D4"]) 
	cursor.setPropertyValue("CellBackColor", 0x8080FF) 
	sheet[0, 0].setString("Start: {}".format(getRangeAddressesAsString(cursor)))
# 	methods = "gotoNext", "gotoEnd", "gotoPrevious", "gotoStart", "gotoEnd"
# 	for i, method in enumerate(methods, start=1):
# 		getattr(cursor, method)()
# 		sheet[i, 0].setString("{}: {}".format(method, getRangeAddressesAsString(cursor)))
# 	cursor = sheet.createCursorByRange(sheet["B5:D4"]) 
	methods = "gotoNext", "gotoEnd", "gotoStart", "gotoEnd", "gotoPrevious"
	for i, method in enumerate(methods, start=1):
		getattr(cursor, method)()
		sheet[i, 0].setString("{}: {}".format(method, getRangeAddressesAsString(cursor)))	
	
# 	i = 0
# 	sheet[i, 0].setString("Start: {}".format(getRangeAddressesAsString(cursor)))
# 	
# 	i += 1
# 	cursor.gotoEnd()
# 	sheet[i, 0].setString("gotoEnd: {}".format(getRangeAddressesAsString(cursor)))
# 						
# 	i += 1
# 	cursor.gotoPrevious()	
# 	sheet[i, 0].setString("gotoPrevious: {}".format(getRangeAddressesAsString(cursor)))
# 	
# 	i += 1
# 	cursor.gotoNext()	
# 	sheet[i, 0].setString("gotoPrevious: {}".format(getRangeAddressesAsString(cursor)))	
	
	
	
	
	sheet[:, 0].getColumns().setPropertyValue("OptimalWidth", True)  # 列幅を最適化する。
def getRangeAddressesAsString(rng):  # 文字列アドレスを返す。
	absolutename = rng.getPropertyValue("AbsoluteName") # セル範囲コレックションは$Sheet1.$A$4:$A$6,$Sheet1.$B$4という形式で返る。
	names = absolutename.replace("$", "").split(",")  # $を削除してセル範囲のリストにする。
	addresses = []
	for name in names:
		addresses.append(name.split(".")[-1])  # シート名を削除する。
	return ", ".join(addresses)
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