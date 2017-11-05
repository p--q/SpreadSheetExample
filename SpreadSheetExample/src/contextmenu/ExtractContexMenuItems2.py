#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.uno import RuntimeException
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet import CellFlags as cf # 定数
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
	def wrapper(*args, **kwargs):
		frame = None
		doc = XSCRIPTCONTEXT.getDocument()
		if doc:  # ドキュメントが取得できた時
			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
		else:
			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
			frame = currentframe.getCreator()
		if frame:   
			import time
			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
			t = 1  # プレグレスバーの初期値。
			while t<=maxrange:  # プログレスバーの最大値以下の間。
				indicator.setValue(t)  # プレグレスバーの位置を設定。
				time.sleep(1)  # 1秒待つ。
				t += 1  # プログレスバーの目盛りを増やす。
			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
def macro(documentevent=None):  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラーを取得。
	contextmenuinterceptor = ContextMenuInterceptor(doc)
	controller.registerContextMenuInterceptor(contextmenuinterceptor)
	if __name__ == "__main__":  # オートメーションで実行するときのみ。ScriptingURLにグローバル変数は渡せない。
		print("Press 'Return' to remove the context menu interceptor.")
		input()  # 入力待ちにしないとスクリプトが終了してしまう。逆にマクロでinput()はフリーズする。
		controller.releaseContextMenuInterceptor(contextmenuinterceptor)
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
	def __init__(self, doc):		
		ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
		self.args = getBaseURL(ctx, smgr, doc), createDispatchCommandLabelReader(ctx, smgr)
# 	@enableRemoteDebugging
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。
		baseurl, getDiapatchCommandLabel = self.args
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
		controller = contextmenuexecuteevent.Selection  # ドキュメントのコントローラの取得。
		global enumerateMenuEntries  # ScriptingURLで呼び出す関数。オートメーションやAPSOでは不可。
		enumerateMenuEntries = createEnumerator(controller, contextmenu, getDiapatchCommandLabel)  # クロージャーでScriptingURLで呼び出す関数に変数を渡す。
		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "MenuEntries", "CommandURL": baseurl.format(enumerateMenuEntries.__name__)})  # CommandURLで渡す関数にデコレーターは不可。
		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})  # 区切り線の挿入。
		return EXECUTE_MODIFIED # EXECUTE_MODIFIED, IGNORED, CANCELLED, CONTINUE_MODIFIED	
def createEnumerator(controller, contextmenu, getDiapatchCommandLabel):
	props = "Text", "CommandURL", "HelpURL", "Image", "SubContainer"  # ActionTriggerのプロパティ。
	separatortypes = {0:"LINE", 1:"SPACE", 2:"LINEBREAK"}  # 定数ActionTriggerSeparatorTypeを文字列に変換。		
	def enumerateMenuEntries():  # ScriptingURLで渡すので引数は受け取れない。
		sheet = controller.getActiveSheet()  # アクティブなシートを取得。
		def _enumarateEntries(container, k, c):  # 第2引数は出力先の開始行。第3引数は出力先の開始列。
			r = k - 1
			for menuentry in container:
				r += 1
				if menuentry.supportsService("com.sun.star.ui.ActionTrigger"):
					text, commandurl, helpurl, image, subcontainer = [menuentry.getPropertyValue(prop) for prop in props]  # getPropertyValues()は実装されていない。
					propvalues = [text, commandurl, helpurl]
					propvalues.append("icon" if image else str(image)) 
					propvalues.append("submenu" if subcontainer else str(subcontainer)) 
					if commandurl.startswith(".uno:"):
						label = getDiapatchCommandLabel(commandurl)
						if label:
							propvalues.append(label)
					sheet[r, c].setString(", ".join(propvalues))
					if subcontainer:
						r = _enumarateEntries(subcontainer, r, c+1)  # 再帰呼出し。
				elif menuentry.supportsService("com.sun.star.ui.ActionTriggerSeparator"):
					separatortype = menuentry.getPropertyValue("SeparatorType")
					sheet[r, c].setString(separatortypes[separatortype])	
			return r		
		sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。cf.HARDATTR+cf.STYLESでセル結合も解除。		
		datarows = (contextmenu.getName(),),\
				(", ".join(props),)
		sheet[:len(datarows), :len(datarows[0])].setDataArray(datarows)
		_enumarateEntries(contextmenu[2:], 2, 0)  # このマクロで追加した項目と線は出力しない。つまり項目インデックス2から出力。第2引数は出力先の開始行。第3引数は出力先の開始列。
		cellcursor = sheet.createCursor()  # シート全体のセルカーサーを取得。
		cellcursor.gotoEndOfUsedArea(True)  # 使用範囲の右下のセルまでにセルカーサーのセル範囲を変更する。
		cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。
	return enumerateMenuEntries
def createDispatchCommandLabelReader(ctx, smgr):	
	rootpaths = "/org.openoffice.Office.UI.CalcCommands/UserInterface/Commands/{}", \
				"/org.openoffice.Office.UI.CalcCommands/UserInterface/Popups/{}", \
				"/org.openoffice.Office.UI.GenericCommands/UserInterface/Commands/{}"  # ルートパスのCalc用のタプル。
	props = "PopupLabel", "Label"  # , "ContextLabel"  # 取得するプロパティのタプル。存在すれば使用したいラベル順に並べる。PopupLabelはコンテクストメニュー用、ContextLabelはツールバー用。
	configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。
	def getDiapatchCommandLabel(dispatchcommand):
		if dispatchcommand.startswith(".uno:"):
			for rootpath in rootpaths:
				rootpath = rootpath.format(dispatchcommand)
				try:
					root = configreader(rootpath)
					propvalues = root.getPropertyValues(props)  # 設定されていないプロパティはNoneが入る
					for label in propvalues:
						if label is not None:
							return label
				except RuntimeException:
					continue
	return getDiapatchCommandLabel
def createConfigReader(ctx, smgr):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	def configReader(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return configReader
def getBaseURL(ctx, smgr, doc):	 # 埋め込みマクロ、オートメーション、マクロセレクターに対応してScriptingURLのbaseurlを返す。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	if modulepath.startswith(ucp):  # 埋め込みマクロの時。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
		filepath = modulepath.replace(ucp, "")  #  ucpを除去。
		transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
		transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
		contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
		macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
		location = "document"  # マクロの場所。	
	else:
		filepath = unohelper.fileUrlToSystemPath(modulepath) if modulepath.startswith("file://") else modulepath # オートメーションの時__file__はシステムパスだが、マクロセレクターから実行するとfileurlが返ってくる。
		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
		fileurl = pathsubstservice.substituteVariables("$(user)/Scripts/python", True)  # $(user)を変換する。fileurlが返ってくる。
		macrofolder =  unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステムパスに変換する。マイマクロフォルダへのパス。	
		location = "user"  # マクロの場所。
	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
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