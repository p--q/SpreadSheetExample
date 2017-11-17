#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os, sys
from itertools import zip_longest
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet import CellFlags as cf # 定数
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.ui.dialogs import ExtendedFilePickerElementIds  # 定数
from com.sun.star.ui.dialogs import ControlActions  # 定数
from com.sun.star.ui.dialogs import TemplateDescription  # 定数
from com.sun.star.ui.dialogs import ExecutableDialogResults  # 定数
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
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, doc, sheet):
		self.baseurl = getBaseURL(ctx, doc)  # ScriptingURLのbaseurlを取得。
		global exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet   # ScriptingURLで呼び出す関数。オートメーションやAPSOでは不可。
		exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet = globalFunctionCreator(ctx, doc, sheet)  # クロージャーでScriptingURLで呼び出す関数に変数を渡す。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 引数はContextMenuExecuteEvent Struct。
		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # すでにあるコンテクストメニュー(アクショントリガーコンテナ)を取得。
		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer")  # サブメニューにするアクショントリガーコンテナをインスタンス化。
		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Export as CSV...", "CommandURL": baseurl.format(exportAsCSV.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Export as PDF...", "CommandURL": baseurl.format(exportAsPDF.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 2, {"Text": "Export as ODS...", "CommandURL": baseurl.format(exportAsODS.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
		addMenuentry(submenucontainer, "ActionTrigger", 3, {"Text": "Selection to New Sheet", "CommandURL": baseurl.format(SelectionToNewSheet.__name__)})
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
	def exportAsCSV():  # アクティブシートをCSV形式で保存する。
		title = "Export as CSV"  # ファイル選択ダイアログのタイトル。
		filtername = "Text - txt - csv (StarCalc)"  # csvのフィルターネーム。
		exportextension = "csv"  # csvの拡張子。
		templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS  # パスワードとフィルター編集チェックボックス付きファイル選択ダイアログ。
		uiname, uicomponent = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
		newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
		kwargs = {"TemplateDescription": templatedescription, "setTitle": title, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}  # キーTemplateDescriptionは必須。
		filepicker = createFilePicker(ctx, smgr, kwargs)  # ファイル選択ダイアログを取得。
		filepicker.setValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.SET_SELECT_ITEM, True)  # ファイル保存ダイアログのフィルター編集チェックボックスにチェックをつける。	
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_AUTOEXTENSION, False)  # 拡張子をつけるチェックボックスを無効にする。Windowsのみ関係する。
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
	def exportAsPDF():  # アクティブシートをPDF形式で保存する。
		title = "Export as PDF"  # ファイル選択ダイアログのタイトル。
		filtername = "calc_pdf_Export"  # フィルターネーム。
		exportextension = "pdf"  # 拡張子。
		templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS				
		uiname, uicomponent = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
		newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
		kwargs = {"TemplateDescription": templatedescription, "setTitle": title, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}
		filepicker = createFilePicker(ctx, smgr, kwargs)		
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。
		filepicker.setValue(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, ControlActions.SET_SELECT_ITEM, True)  # ファイル保存ダイアログのフィルター編集チェックボックスにチェックをつける。	
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_FILTEROPTIONS, False)  # パスワードチェックボックスを無効にする。	
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_AUTOEXTENSION, False)  # 拡張子をつけるチェックボックスを無効にする。Windowsのみ関係する。
		if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
			newdoc = toNewDoc(ctx, doc, name)								
			filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
			filteroptiondialog.setSourceDocument(newdoc)  # 変換元のドキュメントを設定。。
			filteroptiondialog.setPropertyValues((PropertyValue(Name="FilterName", Value=filtername),)) 										
			if filteroptiondialog.execute()==ExecutableDialogResults.CANCEL:  # フィルターのオプションダイアログを表示。
				return True
			newdoc.storeToURL(filepicker.getFiles()[0], filteroptiondialog.getPropertyValues())	 # storeAsURL()はだめ。
			newdoc.close(True)  # 新規ドキュメントを閉じないと.~lock.ExportExample.csv#といったファイルが残ってしまう。		
	def exportAsODS():  # アクティブシートをODS形式で保存する。
		title = "Export as ODS"  # ファイル選択ダイアログのタイトル。
		filtername = "calc8"  # フィルターネーム。
		exportextension = "ods"  # 拡張子。
		templatedescription = TemplateDescription.FILESAVE_AUTOEXTENSION_PASSWORD
		uiname, dummy = root[filtername].getPropertyValues(("UIName", "UIComponent"))  # フィルターのプロパティを取得。
		newfilename = "{}.{}".format(name, exportextension)  # 新規ファイル名を作成。
		kwargs = {"TemplateDescription": templatedescription, "setTitle": title, "setDisplayDirectory": fileurl, "setDefaultName": newfilename, "appendFilter": (uiname, exportextension)}
		filepicker = createFilePicker(ctx, smgr, kwargs)								
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, False)  # パスワードチェックボックスを無効にする。	
		filepicker.enableControl(ExtendedFilePickerElementIds.CHECKBOX_AUTOEXTENSION, False)  # 拡張子をつけるチェックボックスを無効にする。Windowsのみ関係する。
		if filepicker.execute()==ExecutableDialogResults.OK:  # ファイル保存ダイアログを表示する。
			newdoc = toNewDoc(ctx, doc, name)				
			passwordoption = filepicker.getValue(ExtendedFilePickerElementIds.CHECKBOX_PASSWORD, ControlActions.GET_SELECTED_ITEM)  # ファイル保存ダイアログのパスワードチェックボックスの状態を取得。
			if passwordoption:  # パスワードチェックボックスがチェックサれている時。
				pass  # パスワード入力ダイアログの実装が必要。
			newdoc.storeToURL(filepicker.getFiles()[0], ())		
			newdoc.close(True)  # 新規ドキュメントを閉じないと.~lock.ExportExample.csv#といったファイルが残ってしまう。		
	def SelectionToNewSheet():  # 選択範囲を新しいシートに切り出す。
		newsheet = getNewSheet(doc, "Selection")  # 新しいシートの取得。
		newindex = newsheet.getRangeAddress().Sheet  # 新しいシートのインデックスを取得。
		newcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 新しいシートでのセル範囲コレクション。あとでselectするために使う。
		def _copyRange(cellrange):  # セル範囲を新しいシートの同位置にコピー。
			celladdress = cellrange[0, 0].getCellAddress()  # セル範囲の左上のセルのアドレスを取得。
			celladdress.Sheet = newindex  # 新しいシートのアドレスにする。
			sheet.copyRange(celladdress, cellrange.getRangeAddress())  # セル範囲を新しいシートの同じ位置にコピー。	
			cellrangeaddress = cellrange.getRangeAddress()  # セル範囲のアドレスを取得。
			cellrangeaddress.Sheet = newindex  # 新しいシートでのセル範囲のアドレスにする。
			newcellranges.addRangeAddress(cellrangeaddress, False)  # セル範囲コレクションに追加する。セル範囲は結合しない。
		selection = doc.getCurrentSelection()  # 選択範囲を取得。
		if selection.supportsService("com.sun.star.sheet.SheetCellRanges"):  # セル範囲コレクションの時。
			for cellrange in selection:  # 各セル範囲について。
				_copyRange(cellrange)  # セル範囲を新しいシートの同位置にコピー。			
		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル範囲の時。
			_copyRange(selection)  # セル範囲を新しいシートの同位置にコピー。
		controller = doc.getCurrentController()  # コントローラの取得。
		controller.setActiveSheet(newsheet)  # シートをアクティブにする。
		controller.select(newcellranges)  # 元のシートと同位置のセルを選択状態にする。
		cellcursor = newsheet.createCursor()  # シート全体のセルカーサーを取得。
		cellcursor.gotoEndOfUsedArea(True)  # 使用範囲の右下のセルまでにセルカーサーのセル範囲を変更する。
		cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。	
	return exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet
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
def createFilePicker(ctx, smgr, kwargs):  # ファイル選択ダイアログを返す。kwargsはFilePickerのメソッドをキー、引数を値とする辞書。
	key = "TemplateDescription"
	if key in kwargs:
		filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (kwargs.pop(key),), ctx)  # キャッシュするとおかしくなる。
		if kwargs:
			key = "appendFilter"
			if key in kwargs:
				uiname, exportextension = kwargs.pop(key)
				displayfilter = uiname if sys.platform.startswith('win') else "{} (.{})".format(uiname, exportextension)  # 表示フィルターの作成。Windowsの場合は拡張子を含めない。
				getattr(filepicker, key)(displayfilter, "*.{}".format(exportextension))
			if kwargs:
				[getattr(filepicker, key)(val) for key, val in kwargs.items()]
		return filepicker
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
	index = len(sheets)
	sheets.insertNewByName(newname, index)   # 新しいシートを挿入。同名のシートがあるとRuntimeExceptionがでる。
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
def createConfigReader(ctx, smgr):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	def configReader(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return configReader
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
