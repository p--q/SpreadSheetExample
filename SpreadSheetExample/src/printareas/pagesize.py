#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
def macro():
	doc = XSCRIPTCONTEXT.getDocument()
	defaultpagestyle = doc.getStyleFamilies()["PageStyles"]["Default"]
	properties = "Height", "TopMargin", "BottomMargin", "HeaderIsOn", "HeaderHeight", "FooterIsOn", "FooterHeight"
	height, topmargin, bottommargin, headerison, headerheight, footerison, footerheight = defaultpagestyle.getPropertyValues(properties)
	pageheight = height - topmargin - bottommargin  # 印刷高さを1/100mmで取得。
	if headerison:  # ヘッダーがあるときヘッダーの高さを除く。
		pageheight -= headerheight
	if footerison:  # フッターがあるときフッターの高さを除く。
		pageheight -= footerheight
	controller = doc.getCurrentController()
	sheet = controller.getActiveSheet()
	cursor = sheet.createCursorByRange(sheet["A1"])  # A1セルのセルカーサーを取得。
	cursor.gotoEndOfUsedArea(True)  # 使用範囲の右下のセルにセル範囲を拡大する。
	rows = cursor.getRows()  # 使用範囲の行コレクションを取得。
	rows.setPropertyValue("IsStartOfNewPage", False)  # すでにある改ページを消去。
	h = 0  # 行の高さの合計。
	for i in range(len(rows)):  # 行インデックスをイテレート。
		rowheight = rows[i].getPropertyValue("Height")  # 行の高さを取得。
		h += rowheight  # 行の高さを加算する。
		if h>pageheight:  # 1ページあたりの高さを越えた時。
			if i%2:  # 行インデックスが奇数の時。2で割り切れると0になるのでFalseになる。
				rows[i-1].setPropertyValue("IsStartOfNewPage", True)  # 一つ上の行に改ページを挿入。
				h = rows[i-1].getPropertyValue("Height") + rowheight  # 行の高さをリセット。		
			else:  # 行インデックスが偶数の時。
				rows[i].setPropertyValue("IsStartOfNewPage", True)  # 改ページを挿入。
				h = rowheight   # 行の高さをリセット。				
			
		
		

		
		
	
# 	rowheight = rows.getPropertyValue("Height")  # 1行の高さを取得。
# 	rowsperpage = pageheight//rowheight  # 1ページあたりの行数を取得。
	
# 	rows.setPropertyValue("IsStartOfNewPage", True)  #
	
	
	
# 	for i in range(rowsperpage, len(rows), rowsperpage):  # 1ページあたりの行数毎に改ページを設定。
# 		rows[i].setPropertyValue("IsStartOfNewPage", True)	
	
	
# 	for i in rows:
# 		i.setPropertyValue("IsStartOfNewPage", False)
	

		
# 	for i in range(rowsperpage, len(rows)):
# 		rows[i].setPropertyValue("IsStartOfNewPage", True)		
	
	
# 	rows = sheet.getRows()  # 行アクセスオブジェクト。
	
	
	
	
	print(pageheight)
	
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 
# 	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
# 	tcu.wtree(defaultpagestyle)
	
# 	unocontrolnumericfield = smgr.createInstanceWithContext("UnoControlNumericField", ctx)
# 	tcu.wtree(unocontrolnumericfield)
	
	




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
# 		doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
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
	