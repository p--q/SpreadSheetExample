#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.sheet import CellFlags as cf # 定数
from xml.etree import ElementTree
def macro(documentevent=None):  # 引数はイベント駆動用。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	filepath = "/opt/libreoffice5.2/share/registry/res/registry_ja.xcd"  # xmlファイルへのパス。
	xpath = './/node[@oor:name=".uno:FormatCellDialog"]'  # XPath。1つのノードだけ選択する条件にしないといけない。
	namespaces = {"oor": "{http://openoffice.org/2001/registry}",\
				"xs": "{http://www.w3.org/2001/XMLSchema}",\
				"xsi": "{http://www.w3.org/2001/XMLSchema-instance}"}  # 名前空間の辞書。replace()で置換するのに使う。
	queryNodesWithXPath(filepath, xpath, namespaces, doc)
def queryNodesWithXPath(filepath, xpath, namespaces, doc):	
	replaceWithValue, replaceWithKey = createReplaceFunc(namespaces)  # 名前空間を置換する関数を取得。
	xpath = replaceWithValue(xpath)  # 名前空間の辞書のキーを値に変換。
	tree = ElementTree.parse(filepath)
	nodes = tree.findall(xpath)  # XPathで抽出されるノードのリストを取得する。
	datarows = [(replaceWithKey(formatNode(node)),) for node in nodes]  # Calcに出力するために行のリストにする。
	controller = doc.getCurrentController()  # コントローラーを取得。
	sheet = controller.getActiveSheet()  # アクティブなシートを取得。
	sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。cf.HARDATTR+cf.STYLESでセル結合も解除。
	sheet[:len(datarows), :len(datarows[0])].setDataArray(datarows)  # シートに結果を出力する。
	cellcursor = sheet.createCursor()  # シート全体のセルカーサーを取得。
	cellcursor.gotoEndOfUsedArea(True)  # 使用範囲の右下のセルまでにセルカーサーのセル範囲を変更する。
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。
def formatNode(node):  # 引数はElement オブジェクト。タグ名と属性を出力する。属性の順番は保障されない。
	tag = node.tag  # タグ名を取得。
	attribs = []  # 属性をいれるリスト。
	for key, val in node.items():  # ノードの各属性について。
		attribs.append('{}="{}"'.format(key, val))  # =で結合。
	attrib = " ".join(attribs)  # すべての属性を結合。
	n = "{} {}".format(tag, attrib) if attrib else tag  # タグ名と属性を結合する。
	return "<{}>".format(n)	 
def createReplaceFunc(namespaces):  # 引数はキー名前空間名、値は名前空間を波括弧がくくった文字列、の辞書。
	def replaceWithValue(txt):  # 名前空間の辞書のキーを値に置換する。
		for key, val in namespaces.items():
			txt = txt.replace("{}:".format(key), val)
		return txt
	def replaceWithKey(txt):  # 名前空間の辞書の値をキーに置換する。
		for key, val in namespaces.items():
			txt = txt.replace(val, "{}:".format(key))
		return txt
	return replaceWithValue, replaceWithKey
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