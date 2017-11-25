#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.lang import Locale
def macro():  # 引数はcom.sun.star.document.DocumentEvent Struct。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	sheets = doc.getSheets()  # シートコレクションの取得。
	sheet = sheets[0]  # インデックス0のシートを取得。
	sheet.clearContents(511)  # シートのセルの内容をすべてを削除。
	controller = doc.getCurrentController()  # コントローラの取得。
	baseurl = getBaseURL(ctx, smgr, doc)
	contextmenuinterceptor = ContextMenuInterceptor(baseurl)
	controller.registerContextMenuInterceptor(contextmenuinterceptor)  # コントローラにContextMenuInterceptorを登録する。
	global getFormatKey  # ScriptingURLで呼び出されたマクロで使うためにグローバル変数にする。
	getFormatKey = formatkeyCreator(doc)  # ドキュメントのformatkeyを返す関数を取得。
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)
	global todayvalue
	todayvalue = functionaccess.callFunction("TODAY", ())  # スプレッドシート関数で今日の日付のシリアル値を取得。
	if __name__ == "__main__":  # オートメーションで実行するときのみ。
		print("Press 'Return' to remove the context menu interceptor.")
		input()  # 入力待ちにしないとスクリプトが終了してしまう。逆にマクロでinput()はフリーズする。
		controller.releaseContextMenuInterceptor(contextmenuinterceptor)
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, baseurl):
		self.baseurl = baseurl
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 引数はContextMenuExecuteEvent Struct。
		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # すでにあるコンテクストメニュー(アクショントリガーコンテナ)を取得。
		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer")  # サブメニューにするアクショントリガーコンテナをインスタンス化。
		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Set ~Today", "CommandURL": baseurl.format(getToday.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Set ~Yesterday", "CommandURL": baseurl.format(getYesterday.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 2, {"Text": "Set T~omorrow", "CommandURL": baseurl.format(getTomorrow.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "Customized Menu", "SubContainer": submenucontainer})  # サブメニューを一番上に挿入。
		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})  # アクショントリガーコンテナのインデックス1にセパレーターを挿入。
		return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def getToday():
	setDate(todayvalue)  # 今日の日付を取得。
def getYesterday():
	setDate(todayvalue-1)  # 昨日の日付を取得。
def getTomorrow():
	setDate(todayvalue+1)  # 明日の日付を取得。
def setDate(datetimevalue):  # セルに日付を入力する。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	selection = doc.getCurrentSelection()  # セル範囲を取得。
	firstcell = getFirtstCell(selection)  # セル範囲の左上のセルを取得。
	firstcell.setFormula(datetimevalue)  # 日付の入力は年-月-日 または 月/日/年 にしないといけないらしい。
	firstcell.setPropertyValue("NumberFormat", getFormatKey("YYYY-MM-DD"))  # セルの書式を設定。	
def formatkeyCreator(doc):
	numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
	locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。	
	def getFormatKey(formatstring):  # formatstringからFormatKeyを返す。
		formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。	
		if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
			formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。
		return formatkey
	return getFormatKey
def getFirtstCell(rng):  # セル範囲の左上のセルを返す。引数はセルまたはセル範囲またはセル範囲コレクション。
	if rng.supportsService("com.sun.star.sheet.SheetCellRanges"):  # セル範囲コレクションのとき
		rng = rng[0]  # 最初のセル範囲のみ取得。
	return rng[0, 0]  # セル範囲の最初のセルを返す。
def getBaseURL(ctx, smgr, doc):	 # 埋め込みマクロ、オートメーション、マクロセレクターに対応してScriptingURLのbaseurlを返す。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	if modulepath.startswith(ucp):  # 埋め込みマクロの時。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
		filepath = modulepath.replace(ucp, "")  #  ucpを除去。
		transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
		transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
		contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
		macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
		location = "document"	
	else:
		filepath = unohelper.fileUrlToSystemPath(modulepath) if modulepath.startswith("file://") else modulepath # オートメーションの時__file__はシステムパスだが、マクロセレクターから実行するとfileurlが返ってくる。
		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
		fileurl = pathsubstservice.substituteVariables("$(user)/Scripts/python", True)  # $(user)を変換する。fileurlが返ってくる。
		macrofolder =  unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステムパスに変換する。マイマクロフォルダへのパス。	
		location = "user"
	relpath = os.path.relpath(filepath, start=macrofolder)  # パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにパス区切りを置換。
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
g_exportedScripts = macro, #マクロセレクター(ScriptingURLで呼び出すための設定は不要)に限定表示させる関数をタプルで指定。
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