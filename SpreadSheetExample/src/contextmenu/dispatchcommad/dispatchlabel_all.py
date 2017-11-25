#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.uno import RuntimeException
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet import CellFlags as cf # 定数
def macro(documentevent=None):  # 引数はドキュメントイベント駆動用。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。	
	props = "Label", "ContextLabel", "PopupLabel"  # 取得するプロパティのタプル。
	compornents = "ReportCommands", "CalcCommands", "MathCommands", "GenericCommands", "DrawImpressCommands", \
					"StartModuleCommands", "ChartCommands", "BasicIDECommands", "DbuCommands", "BibliographyCommands", "WriterCommands"
	subnodes = "Commands", "Popups"
	datarows = []  # 行のリストのリスト。タプルでも良いが、すべての要素のリストの要素数は一致させる必要がある。
	setIndicator, endIndicator = createStatusIndicator(doc)
	c = 0  # カウンター。
	setIndicator(c)
	for compornent in compornents:
		for subnode in subnodes:
			rootpath = "/org.openoffice.Office.UI.{}/UserInterface/{}".format(compornent, subnode)
			root = None
			try:
				root = configreader(rootpath)	
			except RuntimeException:
				continue
			if root is not None:
				outputs = []
				for childname in root.getElementNames():  # 子ノードの名前のタプルを取得。ノードオブジェクトの直接取得はできない模様。
					node = root.getByName(childname)  # ノードオブジェクトを取得。
					propvalues = ["" if i is None else i for i in node.getPropertyValues(props)]  # 設定されていないプロパティはNoneが入るので""に置換する。
					datarow = [childname]
					datarow.extend(propvalues)
					outputs.append(datarow)
				if outputs:
					c += len(outputs)  # カウンターを増やす。
					setIndicator(c)
					outputs.sort(key=lambda r: r[0])  # ディスパッチコマンドでソートする。
					outputs.append([""]*(len(props)+1))  # 最後に空行を入れる。
					datarow = ["DispatchCommand"]  # 先頭行のセルのリスト。
					datarow.extend(props)  # propsを先頭行に追加。
					rootpathrow = [rootpath]
					rootpathrow.extend([""]*len(props))
					datarows.extend((rootpathrow, datarow))  # 行のセルのリストのリスト。			
					datarows.extend(outputs)
	endIndicator()
	controller = doc.getCurrentController()  # コントローラーを取得。
	sheet = controller.getActiveSheet()  # アクティブなシートを取得。
	sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。cf.HARDATTR+cf.STYLESでセル結合も解除。
	sheet[:len(datarows), :len(datarows[0])].setDataArray(datarows)  # シートに結果を出力する。
	cellcursor = sheet.createCursor()  # シート全体のセルカーサーを取得。
	cellcursor.gotoEndOfUsedArea(True)  # 使用範囲の右下のセルまでにセルカーサーのセル範囲を変更する。
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。
	sheet[0, 1].setString("Got {} Labels".format(c))
def createConfigReader(ctx, smgr):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	def configReader(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return configReader
def createStatusIndicator(doc):  # プログレスバーの表示。
	frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。  
	indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
	maxrange = 2254  # ステータスバーに表示するプログレスバーの目盛りの最大値。
	indicator.start("Collecting labels for dispatch commands.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
	def setIndicator(c):
		indicator.setValue(c)  # プレグレスバーの位置を設定。
	def endIndicator():
		indicator.end()  # reset()の前にend()しておかないと元に戻らない。
		indicator.reset()
	return setIndicator, endIndicator
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