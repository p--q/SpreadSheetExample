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
		if doc.hasLocation():  # ドキュメントが保存されているとき。
			fileurl = os.path.dirname(doc.getLocation())  # ドキュメントの親ディレクトリを取得。
		else:  # ドキュメントが保存れていない時はホームディレクトリを取得。
			pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
			fileurl = pathsubstservice.getSubstituteVariableValue("$(home)")
		self.args = ctx, smgr, configreader, props, doc, fileurl
# 	@enableRemoteDebugging  # ダブルクリニックで2回呼ばれる。2回目はGUIで操作できないときあり。
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		ctx, smgr, configreader, props, doc, fileurl = self.args
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 				import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # これでブレークすべき。
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
					if celladdress.Row<4 and celladdress.Column<1:
						txt = target.getString()  # セルの文字列を取得。
						sheet = target.getSpreadsheet()  # ターゲットがあるシートを取得。
						name = sheet.getName()  # シート名を取得。
						if txt.startswith("CSV"):
							filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS,), ctx)  # キャッシュするとおかしくなる。
							filepicker.setDisplayDirectory(fileurl)  # ファイル保存ダイアログで、デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
							filtername = "Text - txt - csv (StarCalc)"  # フィルターネーム。
							root = configreader("/org.openoffice.TypeDetection.Filter/Filters/{}".format(filtername))  # コンフィギュレーションのルートを取得。	
							uiname, uicomponent, exportextension = root.getPropertyValues(props)  # フィルターのプロパティを取得。	
							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
							filepicker.setTitle(txt)  # ファイル保存ダイアログのタイトルを設定。
							filepicker.setDefaultName(newfilename)  # ファイル保存ダイアログのデフォルトファイル名を設定。
							displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
							filepicker.appendFilter(displayfilter, exportextension)  # ファイル選択ダイアログに表示フィルターを設定。filepickerをキャッシュしている2回目で止まる。
							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
								propertyvalues = PropertyValue(Name="Hidden",Value=True),  # 新しいドキュメントのプロパティ。
								newdoc = ctx.getByName('/singletons/com.sun.star.frame.theDesktop').loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  # 新規ドキュメントの取得。
								newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
								newsheets.importSheet(doc, name, 0)  # 新規ドキュメントにシートをコピー。
								del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。							
								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
								propertyvalues = PropertyValue(Name="FilterName", Value=filtername),  # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
								filteroptiondialog.setPropertyValues(propertyvalues)  # XPropertyAccessインターフェイスのメソッド。
								filteroption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのフィルター編集チェックボックスの状態を取得。								
								propertyvalues = list(propertyvalues)  # CSVの場合にはFilterNameが入ってこないのでリストにしてFilterOptionsを追加する。
								if filteroption:  # ファイル選択ダイアログのフィルター編集チェックボックスがチェックされている時。			
									if filteroptiondialog.execute()==ExecutableDialogResults.OK:  # フィルターのオプションダイアログを表示。execute()で前回の設定値が入るわけではない模様。
										propertyvalues.extend(filteroptiondialog.getPropertyValues())  # FilterOptionsを取得。
									else:
										return True
								else:  # ファイル選択ダイアログのフィルター編集チェックボックスがチェックされていない時は決め打ちする。	
									propertyvalues.append(PropertyValue(Name="FilterOptions", Value="44,34,76,1,,0,false,true,true,false"))			
								newdoc.storeAsURL(filepicker.getFiles()[0], propertyvalues)  # ファイル選択ダイアログで取得したファイルを保存する。		
						elif txt.startswith("PNG"):
							filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (TemplateDescription.FILESAVE_AUTOEXTENSION_SELECTION,), ctx)  # キャッシュするとおかしくなる。
							filepicker.setDisplayDirectory(fileurl)  # ファイル保存ダイアログで、デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
							filtername = "calc_png_Export"
							root = configreader("/org.openoffice.TypeDetection.Filter/Filters/{}".format(filtername))  # コンフィギュレーションのルートを取得。	
							uiname, uicomponent, exportextension = root.getPropertyValues(props)  # フィルターのプロパティを取得。	
							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
							filepicker.setTitle(txt)  # ファイル保存ダイアログのタイトルを設定。
							filepicker.setDefaultName(newfilename)  # ファイル保存ダイアログのデフォルトファイル名を設定。
							displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
							filepicker.appendFilter(displayfilter, exportextension)  # ファイル選択ダイアログに表示フィルターを設定。
							filepicker.setValue(ExtendedFilePickerElementIds.CHECKBOX_SELECTION, ControlActions.SET_SELECT_ITEM, True)  # 選択範囲チェックボックスにチェックを付ける。
							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_SELECTION, False)  # 選択範囲チェックボックスを無効にする。
							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
								propertyvalues = PropertyValue(Name="Hidden",Value=True),  # 新しいドキュメントのプロパティ。
								newdoc = ctx.getByName('/singletons/com.sun.star.frame.theDesktop').loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  # 新規ドキュメントの取得。
								newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
								newsheets.importSheet(doc, name, 0)  # 新規ドキュメントにシートをコピー。
								cellcursor = newsheets[0].createCursor()
								cellcursor.gotoEndOfUsedArea(True)
								controller = newdoc.getCurrentController()  # コントローラの取得。
								controller.select(cellcursor)
								del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。							
								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
								propertyvalues = PropertyValue(Name="FilterName", Value=filtername),  # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
								filteroptiondialog.setPropertyValues(propertyvalues)  # XPropertyAccessインターフェイスのメソッド。
								filteroption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのフィルター編集チェックボックスの状態を取得。										
								if filteroptiondialog.execute()==ExecutableDialogResults.OK:  # フィルターのオプションダイアログを表示。
									propertyvalues = filteroptiondialog.getPropertyValues()
									newdoc.storeAsURL(filepicker.getFiles()[0], propertyvalues)									
						elif txt.startswith("PDF"):
							filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS,), ctx)  # キャッシュするとおかしくなる。
							filepicker.setDisplayDirectory(fileurl)  # ファイル保存ダイアログで、デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
							filtername = "calc_pdf_Export"
							root = configreader("/org.openoffice.TypeDetection.Filter/Filters/{}".format(filtername))  # コンフィギュレーションのルートを取得。	
							uiname, uicomponent, exportextension = root.getPropertyValues(props)  # フィルターのプロパティを取得。	
							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
							filepicker.setTitle(txt)  # ファイル保存ダイアログのタイトルを設定。
							filepicker.setDefaultName(newfilename)  # ファイル保存ダイアログのデフォルトファイル名を設定。
							displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
							filepicker.appendFilter(displayfilter, exportextension)  # ファイル選択ダイアログに表示フィルターを設定。
							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
							filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.SET_SELECTED_ITEM, True)  # ファイル保存ダイアログのフィルター編集チェックボックスにチェックをつける。	
							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, False)  # パスワードチェックボックスを無効にする。	
							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
								propertyvalues = PropertyValue(Name="Hidden",Value=True),  # 新しいドキュメントのプロパティ。
								newdoc = ctx.getByName('/singletons/com.sun.star.frame.theDesktop').loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  # 新規ドキュメントの取得。
								newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
								newsheets.importSheet(doc, name, 0)  # 新規ドキュメントにシートをコピー。
								del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。							
								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
								propertyvalues = PropertyValue(Name="FilterName", Value=filtername),  # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
								filteroptiondialog.setPropertyValues(propertyvalues)  # XPropertyAccessインターフェイスのメソッド。										
								if filteroptiondialog.execute()==ExecutableDialogResults.OK:  # フィルターのオプションダイアログを表示。
									propertyvalues = filteroptiondialog.getPropertyValues()
									newdoc.storeAsURL(filepicker.getFiles()[0], propertyvalues)	
						elif txt.startswith("ODS"):
							filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD,), ctx)  # キャッシュするとおかしくなる。
							filepicker.setDisplayDirectory(fileurl)  # ファイル保存ダイアログで、デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
							
							
							
							filtername = "calc_pdf_Export"
							root = configreader("/org.openoffice.TypeDetection.Filter/Filters/{}".format(filtername))  # コンフィギュレーションのルートを取得。	
							uiname, uicomponent, exportextension = root.getPropertyValues(props)  # フィルターのプロパティを取得。	
							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
							filepicker.setTitle(txt)  # ファイル保存ダイアログのタイトルを設定。
							filepicker.setDefaultName(newfilename)  # ファイル保存ダイアログのデフォルトファイル名を設定。
							displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
							filepicker.appendFilter(displayfilter, exportextension)  # ファイル選択ダイアログに表示フィルターを設定。
							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
							filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.SET_SELECTED_ITEM, True)  # ファイル保存ダイアログのフィルター編集チェックボックスにチェックをつける。	
							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, False)  # パスワードチェックボックスを無効にする。	
							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
								propertyvalues = PropertyValue(Name="Hidden",Value=True),  # 新しいドキュメントのプロパティ。
								newdoc = ctx.getByName('/singletons/com.sun.star.frame.theDesktop').loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  # 新規ドキュメントの取得。
								newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
								newsheets.importSheet(doc, name, 0)  # 新規ドキュメントにシートをコピー。
								del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。							
								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
								propertyvalues = PropertyValue(Name="FilterName", Value=filtername),  # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
								filteroptiondialog.setPropertyValues(propertyvalues)  # XPropertyAccessインターフェイスのメソッド。										
								if filteroptiondialog.execute()==ExecutableDialogResults.OK:  # フィルターのオプションダイアログを表示。
									propertyvalues = filteroptiondialog.getPropertyValues()
									newdoc.storeAsURL(filepicker.getFiles()[0], propertyvalues)										
									
									
																			
								
						newdoc.close(True)  # 新規ドキュメントを閉じないと.~lock.ExportExample.csv#といったファイルが残ってしまう。
						return False  # セル編集モードにしない。
		return True
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