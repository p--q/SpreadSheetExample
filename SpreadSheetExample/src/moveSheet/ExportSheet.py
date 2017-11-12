#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os, sys
from com.sun.star.sheet import CellFlags as cf # 定数
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.ui.dialogs import ExtendedFilePickerElementIds  # 定数
from com.sun.star.ui.dialogs import ControlActions  # 定数
from com.sun.star.ui.dialogs import TemplateDescription  # 定数
from com.sun.star.ui.dialogs import ExecutableDialogResults  # 定数
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	sheet = getNewSheet(doc, "ExportExample")  # 新規シートの取得。
	cmds = "CSVとしてエクスポート", "PNGとしてエクスポート", "PDFとしてエクスポート", "ODSとしてエクスポート"
	datarows = [(i,) for i in cmds]
	cellrange = sheet[:len(datarows), :1]
	cellrange.setDataArray(datarows)
	cellrange.getColumns().setPropertyValue("OptimalWidth", True)  # 列幅を最適化する。		
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.setActiveSheet(sheet)  # シートをアクティブにする。	
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
	args = ctx, smgr, doc
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(args))  # マウスハンドラをコントローラに設定。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler): # マウスハンドラ
	def __init__(self, args):
		ctx, smgr, doc = args
		configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。
		props = "UIName", "UIComponent", "ExportExtension"  # 取得するプロパティ名のタプル。
		filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS,), ctx)
		if doc.hasLocation():  # ドキュメントが保存されているとき。
			fileurl = os.path.dirname(doc.getLocation())
		else:
			pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
			fileurl = pathsubstservice.getSubstituteVariableValue("$(home)")
		filepicker.setDisplayDirectory(fileurl)  # デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
		self.args = ctx, smgr, filepicker, configreader, props, doc, fileurl
# 	@enableRemoteDebugging  # ダブルクリニックで2回呼ばれる。2回目はGUIで操作できないときあり。
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		ctx, smgr, filepicker, configreader, props, doc, fileurl = self.args
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 				import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # これでブレークすべき。
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					
					cmd = target.getString()
					if cmd=="CSVとしてエクスポート":
						sheet = target.getSpreadsheet()  # ターゲットがあるシートを取得。
						name = sheet.getName()

						
						propertyvalues = PropertyValue(Name="Hidden",Value=True),
						newdoc = ctx.getByName('/singletons/com.sun.star.frame.theDesktop').loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  
						newsheets = newdoc.getSheets()
						newsheets.importSheet(doc, name, 0)
						del newsheets["Sheet1"]  # デフォルトシートを削除する。
					
						filepicker.setTitle(cmd)
						filtername = "Text - txt - csv (StarCalc)"
						root = configreader("/org.openoffice.TypeDetection.Filter/Filters/{}".format(filtername))  # コンフィギュレーションのルートを取得。	
						uiname, uicomponent, exportextension = root.getPropertyValues(props)  # フィルターのプロパティを取得。
						newfilename = "{}.{}".format(name, exportextension)
						filepicker.setDefaultName(newfilename)
						displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # Windowsの場合は拡張子を含めない。
						filepicker.appendFilter(displayfilter, exportextension)
						filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)
						if filepicker.execute()==ExecutableDialogResults.OK:
							filteroption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.GET_SELECTED_ITEM)	
							if filteroption:
								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
								propertyvalues = PropertyValue(Name="FilterName", Value=filtername),  # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
								filteroptiondialog.setPropertyValues(propertyvalues)  # XPropertyAccessインターフェイスのメソッド。
								if filteroptiondialog.execute()==ExecutableDialogResults.OK:  # フィルターのオプションダイアログを表示。
									propertyvalues = filteroptiondialog.getPropertyValues()  # 戻り値はPropertyValue Structのタプル。XPropertyAccessインターフェイスのメソッド。	
									
									
									newdoc.storeAsURL("{}/{}".format(fileurl, newfilename), propertyvalues)		
						
												
					
					
						newdoc.close(True)
						return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。
	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
		return True  # Trueでイベントを次のハンドラに渡す。
	def disposing(self, eventobject):
		pass	
def getNewSheet(doc, sheetname):  # docに名前sheetnameのシートを返す。sheetnameがすでにあれば連番名を使う。
	cellflags = cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES
	sheets = doc.getSheets()  # シートコレクションを取得。
	c = 1  # 連番名の最初の番号。
	newname = sheetname
	while newname in sheets: # 同名のシートがあるとき。sheets[sheetname]ではFalseのときKeyErrorになる。
		if not sheets[newname].queryContentCells(cellflags):  # シートが未使用のとき
			return sheets[sheetname]  # 未使用の同名シートを返す。
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
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
		import officehelper
		from functools import wraps
		import sys
		from com.sun.star.beans import PropertyValue  # Struct
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