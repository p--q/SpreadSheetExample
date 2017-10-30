#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet import CellFlags  # 定数
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
def macro():  
	
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
# 	global tcu
# 	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
	
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラーを取得。
	contextmenuinterceptor = ContextMenuInterceptor()
	controller.registerContextMenuInterceptor(contextmenuinterceptor)
	if __name__ == "__main__":  # オートメーションで実行するときのみ。
		print("Press 'Return' to remove the context menu interceptor.")
		input()  # 入力待ちにしないとスクリプトが終了してしまう。逆にマクロでinput()はフリーズする。
		controller.releaseContextMenuInterceptor(contextmenuinterceptor)
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
	def __init__(self):
# 		filename = os.path.basename(__file__)  # このファイル名を取得。埋め込みマクロのフルパスは"vnd.sun.star.tdoc:/4/Scripts/python/filename.py"というように番号(LibreOfficeバージョン番号?)が入ってしまう。
		if __file__.startswith("vnd.sun.star.tdoc:"):  # このスクリプトをドキュメントに埋め込んでいる時__file__は"vnd.sun.star.tdoc:/4/Scripts/python/filename.py"というように番号(LibreOfficeバージョン番号?)が入ってしまう。
			fullpath = __file__.replace("vnd.sun.star.tdoc:", "")
			macropath = ""
			flg = False
			for p in fullpath.split("/"):
				if p=="Scripts":
					flg = True
				if flg:
					if p=="python":
						
				
				
				
			
			self.baseurl = "vnd.sun.star.script:{}${}?language=Python&location=document".format(filename, "{}")  # ScriptingURLのbaseurlを取得。
		else:
				
			
		# このスクリプトをマイマクロフォルダに入れている時
# 		vnd.sun.star.script:SpreadSheetExample|SpreadSheetExample|src|etc|calcmacro.py$macro?language=Python&location=user
		self.baseurl = "vnd.sun.star.script:{}${}?language=Python&location=user".format(filename, "{}")  # ScriptingURLのbaseurlを取得。
# 	@enableRemoteDebugging
	def notifyContextMenuExecute(self, contextmenuexecuteevent): 		
		global contextmenu
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer
		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "MenuEntries", "CommandURL": baseurl.format(outputMenuEntries.__name__)})
		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})
		return EXECUTE_MODIFIED # EXECUTE_MODIFIED, IGNORED, CANCELLED, CONTINUE_MODIFIED	
# @enableRemoteDebugging
def outputMenuEntries():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。
	sheets = doc.getSheets()  # ドキュメントのシートコレクションを取得。
	sheet = sheets[0]  # シートコレクションのインデックス0のシートを取得。	
	sheet.clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA+CellFlags.HARDATTR+CellFlags.STYLES)  # セルの内容を削除。
	propnames = "Text", "CommandURL", "HelpURL", "Image", "SubContainer"
	headers = ["MenuType"].extend(propnames)
	sheet[0, :len(headers)].setDataArray((headers,))
	actiontriggerseparatortypes = {0:"ActionTriggerSeparatorType.LINE", 1:"ActionTriggerSeparatorType.SPACE", 2:"ActionTriggerSeparatorType.LINEBREAK"}
	for i, menuentry in enumerate(contextmenu, start=1):
		if menuentry.supportsService("com.sun.star.ui.ActionTrigger"):
			props = menuentry.getPropertyValues(propnames)
			image = False if props[3] is None else True
			subcontainer = False if props[4] is None else True
			cols = props[:3] + (image, subcontainer)
			sheet[i, :len(cols)].setDataArray((cols,))
		elif menuentry.supportsService("com.sun.star.ui.ActionTriggerSeparator"):
			separatortype = menuentry.getPropertyValue("SeparatorType")
			sheet[i, 1:len(headers)].merge(True)
			cols = "ActionTriggerSeparator", actiontriggerseparatortypes[separatortype]
			sheet[i, :len(cols)].setDataArray((cols,))
	sheet[0, :len(headers)].getColumns().setPropertyValue("OptimalWidth", True)  # 列幅を最適化する。
	
		
		




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