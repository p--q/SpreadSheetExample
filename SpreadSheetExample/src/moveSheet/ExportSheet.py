#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。

from itertools import zip_longest
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet import CellFlags as cf # 定数


import os, sys
# 
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.ui.dialogs import ExtendedFilePickerElementIds  # 定数
from com.sun.star.ui.dialogs import ControlActions  # 定数
from com.sun.star.ui.dialogs import TemplateDescription  # 定数
from com.sun.star.ui.dialogs import ExecutableDialogResults  # 定数
# from com.sun.star.awt import XEnhancedMouseClickHandler
# from com.sun.star.awt import MouseButton  # 定数
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	sheet = getNewSheet(doc, "ExportExample")  # 新規シートの取得。
	datarows = ("データ1", "データ2","","データ4"),\
				(),\
				("データ5","","データ6")  # サンプルデータ。
	rowsToSheet(sheet, datarows)  # datarowsをシートに書き出し。	
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.setActiveSheet(sheet)  # シートをアクティブにする。		
	controller.registerContextMenuInterceptor(ContextMenuInterceptor(ctx, doc, sheet))  # コントローラにContextMenuInterceptorを登録する。
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
def rowsToSheet(cellrange, datarows):  # 引数のセル範囲を左上端にして一括書き込みして列幅を最適化する。datarowsはタプルのタプル。
	datarows = tuple(zip(*zip_longest(*datarows, fillvalue="")))  # 一番長い行の長さに合わせて空文字を代入。
	sheet = cellrange.getSpreadsheet()  # セル範囲のあるシートを取得。
	cellcursor = sheet.createCursorByRange(cellrange)  # セル範囲のセルカーサーを取得。
	cellcursor.collapseToSize(len(datarows[0]), len(datarows))  # (列、行)で指定。セルカーサーの範囲をdatarowsに合せる。
	cellcursor.setDataArray(datarows)  # セルカーサーにdatarowsを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。	
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, doc, sheet):
		self.baseurl = getBaseURL(ctx, doc)  # ScriptingURLのbaseurlを取得。
		global exportAsCSV, exportAsPNG, exportAsPDF, exportAsODS  # ScriptingURLで呼び出す関数。オートメーションやAPSOでは不可。
		exportAsCSV, exportAsPNG, exportAsPDF, exportAsODS = globalFunctionCreator(ctx, doc, sheet)  # クロージャーでScriptingURLで呼び出す関数に変数を渡す。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 引数はContextMenuExecuteEvent Struct。
		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # すでにあるコンテクストメニュー(アクショントリガーコンテナ)を取得。
		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer")  # サブメニューにするアクショントリガーコンテナをインスタンス化。
		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Export as CSV...", "CommandURL": baseurl.format(exportAsCSV.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Export as PNG...", "CommandURL": baseurl.format(exportAsPNG.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 2, {"Text": "Export as PDF...", "CommandURL": baseurl.format(exportAsPDF.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 3, {"Text": "Export as ODS...", "CommandURL": baseurl.format(exportAsODS.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "ExportAs", "SubContainer": submenucontainer})  # サブメニューを一番上に挿入。
		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})  # アクショントリガーコンテナのインデックス1にセパレーターを挿入。
		return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
def globalFunctionCreator(ctx, doc, sheet):
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	if doc.hasLocation():  # ドキュメントが保存されているとき。
		fileurl = os.path.dirname(doc.getLocation())  # ドキュメントの親ディレクトリを取得。
	else:  # ドキュメントが保存れていない時はホームディレクトリを取得。
		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
		fileurl = pathsubstservice.getSubstituteVariableValue("$(home)")
	configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。
	root = configreader("/org.openoffice.TypeDetection.Filter/Filters")  # コンフィギュレーションのルートを取得。		
	name = sheet.getName()  # シート名を取得。
	def exportAsCSV():
		title = "Export as CSV"
		filtername = "Text - txt - csv (StarCalc)"  # csvのフィルターネーム。
		exportextension = "csv"  # csvの拡張子。
		templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS  # パスワードとフィルター編集チェックボックス付きファイル選択ダイアログ。
		uiname, uicomponent = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
		newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
		kwargs = {"TemplateDescription": templatedescription, "setTitle": title, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}  # キーTemplateDescriptionは必須。
		filepicker = createFilePicker(ctx, smgr, kwargs)  # ファイル選択ダイアログを取得。
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
		if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
			newfileurl = filepicker.getFiles()[0]  # ファイル選択ダイアログからfileurlを取得。
			newdoc = toNewDoc(ctx, doc, name)  # docのシート名nameのシートを入れたドキュメントを取得。					
			filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
			filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
			filteroptiondialog.setPropertyValues((PropertyValue(Name="FilterName", Value=filtername),)) # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
			filteroption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのフィルター編集チェックボックスの状態を取得。								
			propertyvalues = [PropertyValue(Name="FilterName", Value=filtername)]
			if not filteroption:  # ファイル選択ダイアログのフィルター編集チェックボックスがチェックされていない時。設定決め打ち。	
				propertyvalues.append(PropertyValue(Name="FilterOptions", Value="44,34,76,1,,0,false,true,true,false"))  # UTF-8、フィールド区切りカンマ、テキスト区切りダブルクォーテーション、セルの内容を表示通り。
				newdoc.storeAsURL(newfileurl, propertyvalues)  # ファイル選択ダイアログで取得したパスに保存する。	
			elif filteroption and filteroptiondialog.execute()==ExecutableDialogResults.OK:  # ファイル選択ダイアログのフィルター編集チェックボックスがチェックされている時はフィルターオプションダイアログを表示してそれがOKの時。
				propertyvalues.extend(filteroptiondialog.getPropertyValues())
				newdoc.storeAsURL(newfileurl, propertyvalues)  # ファイル選択ダイアログで取得したパスに保存する。			
			newdoc.close(True)  # 新規ドキュメントを閉じないと.~lock.ExportExample.csv#といったファイルが残ってしまう。
	def exportAsPNG():
		title = "Export as PNG"
		filtername = "calc_png_Export"  # pngのフィルターネーム。
		exportextension = "png"  # pngの拡張子。
		templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_SELECTION # 選択範囲チェックボックス付きファイル選択ダイアログ。			
		uiname, uicomponent = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
		newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
		kwargs = {"TemplateDescription": templatedescription, "setTitle": title, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}  # キーTemplateDescriptionは必須。
		filepicker = createFilePicker(ctx, smgr, kwargs)  # ファイル選択ダイアログを取得。							
		filepicker.setValue(ExtendedFilePickerElementIds.CHECKBOX_SELECTION, ControlActions.SET_SELECT_ITEM, True)  # 選択範囲チェックボックスにチェックを付ける。
		if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
			newfileurl = filepicker.getFiles()[0]  # ファイル選択ダイアログからfileurlを取得。
			
			newdoc = toNewDoc(ctx, doc, name)  # docのシート名nameのシートを入れたドキュメントを取得。	
			
						
			cellcursor = newdoc.getSheets()[0].createCursor()
			cellcursor.gotoEndOfUsedArea(True)
			controller = newdoc.getCurrentController()  # コントローラの取得。
			controller.select(cellcursor)		
			filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
			filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
			filteroptiondialog.setPropertyValues((PropertyValue(Name="FilterName", Value=filtername),))   # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
			filteroption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのフィルター編集チェックボックスの状態を取得。										
			if filteroption:
				if filteroptiondialog.execute()==ExecutableDialogResults.CANCEL:  # フィルターのオプションダイアログを表示。デフォルト値がfilteroptiondialogのPropertyValuesに入る。
					return True  # キャンセルボタンがクリックされたとき。
		newdoc.storeToURL(filepicker.getFiles()[0], filteroptiondialog.getPropertyValues())	 # storeAsURLはダメ。		
	def exportAsPDF():
		pass
	def exportAsODS():
		pass
	return exportAsCSV, exportAsPNG, exportAsPDF, exportAsODS

	
	
	

# 						elif txt.startswith("PNG"):  # pngで切り出す。
# 							filtername = "calc_png_Export"  # pngのフィルターネーム。
# 							exportextension = "png"  # pngの拡張子。
# 							templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_SELECTION # 選択範囲チェックボックス付きファイル選択ダイアログ。			
# 							uiname, uicomponent = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
# 							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
# 							kwargs = {"TemplateDescription": templatedescription, "setTitle": txt, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}  # キーTemplateDescriptionは必須。
# 							filepicker = createFilePicker(ctx, smgr, kwargs)  # ファイル選択ダイアログを取得。							
# 							filepicker.setValue(ExtendedFilePickerElementIds.CHECKBOX_SELECTION, ControlActions.SET_SELECT_ITEM, True)  # 選択範囲チェックボックスにチェックを付ける。
# # 							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_SELECTION, False)  # 選択範囲チェックボックスを無効にする。
# 							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
# 								newdoc = toNewDoc(ctx, doc, name)  # docのシート名nameのシートを入れたドキュメントを取得。	
# 								
# 											
# 								cellcursor = newdoc.getSheets()[0].createCursor()
# 								cellcursor.gotoEndOfUsedArea(True)
# 								controller = newdoc.getCurrentController()  # コントローラの取得。
# 								controller.select(cellcursor)		
# 								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
# 								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。
# 								filteroptiondialog.setPropertyValues((PropertyValue(Name="FilterName", Value=filtername),))   # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
# 								filteroption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのフィルター編集チェックボックスの状態を取得。										
# 								if filteroption:
# 									if filteroptiondialog.execute()==ExecutableDialogResults.CANCEL:  # フィルターのオプションダイアログを表示。デフォルト値がfilteroptiondialogのPropertyValuesに入る。
# 										return True  # キャンセルボタンがクリックされたとき。
# 								newdoc.storeToURL(filepicker.getFiles()[0], filteroptiondialog.getPropertyValues())	 # storeAsURLはダメ。								
# 						elif txt.startswith("PDF"): 
# 							filtername = "calc_pdf_Export"  # フィルターネーム。
# 							exportextension = "pdf"
# 							templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS				
# 							uiname, uicomponent = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
# 							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
# 							kwargs = {"TemplateDescription": templatedescription, "setTitle": txt, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}
# 							filepicker = createFilePicker(ctx, smgr, kwargs)		
# 							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
# 							filepicker.setValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.SET_SELECT_ITEM, True)  # ファイル保存ダイアログのフィルター編集チェックボックスにチェックをつける。	
# 							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, False)  # パスワードチェックボックスを無効にする。	
# 							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
# 								newdoc = toNewDoc(ctx, doc, name)								
# 								filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
# 								filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。。
# 								filteroptiondialog.setPropertyValues((PropertyValue(Name="FilterName", Value=filtername),)) 										
# 								if filteroptiondialog.execute()==ExecutableDialogResults.CANCEL:  # フィルターのオプションダイアログを表示。
# 									return True
# 								newdoc.storeToURL(filepicker.getFiles()[0], filteroptiondialog.getPropertyValues())	 # storeAsURL()はだめ。
# 						elif txt.startswith("ODS"):
# 							filtername = "calc8"
# 							exportextension = "ods"
# 							templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD
# 							uiname, dummy = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
# 							newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
# 							kwargs = {"TemplateDescription": templatedescription, "setTitle": txt, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}
# 							filepicker = createFilePicker(ctx, smgr, kwargs)								
# # 							filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。	
# 							if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
# 								newdoc = toNewDoc(ctx, doc, name)				
# 								passwordoption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのパスワードチェックボックスの状態を取得。
# 								if passwordoption:
# 									pass  # パスワード入力ダイアログの実装が必要。
# 								newdoc.storeToURL(filepicker.getFiles()[0], ())																						
# 						if newdoc is not None:	
# 							newdoc.close(True)  # 新規ドキュメントを閉じないと.~lock.ExportExample.csv#といったファイルが残ってしまう。
# 							return False  # セル編集モードにしない。
# 		return True
# 	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
# 		return True  # Trueでイベントを次のハンドラに渡す。
# 	def disposing(self, eventobject):
# 		pass	
def getBaseURL(ctx, doc):	 # 埋め込みマクロ、オートメーション、マクロセレクターに対応してScriptingURLのbaseurlを返す。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
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
def toNewDoc(ctx, doc, name):  # 移動元doc、移動させるシート名name
	propertyvalues = PropertyValue(Name="Hidden",Value=True),  # 新しいドキュメントのプロパティ。
	newdoc = ctx.getByName('/singletons/com.sun.star.frame.theDesktop').loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  # 新規ドキュメントの取得。
	newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
	newsheets.importSheet(doc, name, 0)  # 新規ドキュメントにシートをコピー。
	del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。	
	return newdoc
def createFilePicker(ctx, smgr, kwargs):
	key = "TemplateDescription"
	if key in kwargs:
		filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (kwargs.pop(key),), ctx)  # キャッシュするとおかしくなる。
		if kwargs:
			key = "appendFilter"
			if key in kwargs:
				uiname, exportextension = kwargs.pop(key)
				displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
				getattr(filepicker, key)(displayfilter, exportextension)
			if kwargs:
				[getattr(filepicker, key)(val) for key, val in kwargs.items()]
		return filepicker
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