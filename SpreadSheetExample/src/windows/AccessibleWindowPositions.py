#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from itertools import zip_longest
from com.sun.star.sheet import CellFlags as cf # 定数
DIC_ACCESSIBLEROLE = {'26': 'HEADING', '24': 'GROUP_BOX', '39': 'PAGE_TAB_LIST', '41': 'PARAGRAPH', '17': 'FILLER', '51': 'SCROLL_PANE', '48': 'ROW_HEADER', '29': 'INTERNAL_FRAME', '84': 'DOCUMENT_SPREADSHEET', '60': 'TEXT', '7': 'COMBO_BOX', '35': 'MENU_BAR', '36': 'MENU_ITEM', '70': 'CAPTION', '27': 'HYPER_LINK', '21': 'FRAME', '71': 'CHART', '11': 'DIRECTORY_PANE', '40': 'PANEL', '57': 'STATUS_BAR', '80': 'TREE_TABLE', '33': 'LIST_ITEM', '78': 'SECTION', '61': 'TEXT_FRAME', '32': 'LIST', '3': 'CANVAS', '0': 'UNKNOWN', '5': 'CHECK_MENU_ITEM', '20': 'FOOTNOTE', '64': 'TOOL_TIP', '13': 'DOCUMENT', '28': 'ICON', '69': 'BUTTON_MENU', '53': 'SEPARATOR', '81': 'COMMENT', '18': 'FONT_CHOOSER', '50': 'SCROLL_BAR', '56': 'SPLIT_PANE', '82': 'COMMENT_END', '45': 'PROGRESS_BAR', '38': 'PAGE_TAB', '47': 'RADIO_MENU_ITEM', '44': 'PUSH_BUTTON', '43': 'POPUP_MENU', '55': 'SPIN_BOX', '2': 'COLUMN_HEADER', '14': 'EMBEDDED_OBJECT', '72': 'EDIT_BAR', '49': 'ROOT_PANE', '77': 'RULER', '75': 'NOTE', '66': 'VIEW_PORT', '65': 'TREE', '79': 'TREE_ITEM', '19': 'FOOTER', '23': 'GRAPHIC', '67': 'WINDOW', '52': 'SHAPE', '34': 'MENU', '22': 'GLASS_PANE', '37': 'OPTION_PANE', '83': 'DOCUMENT_PRESENTATION', '8': 'DATE_EDITOR', '30': 'LABEL', '10': 'DESKTOP_PANE', '58': 'TABLE', '59': 'TABLE_CELL', '9': 'DESKTOP_ICON', '6': 'COLOR_CHOOSER', '85': 'DOCUMENT_TEXT', '46': 'RADIO_BUTTON', '76': 'PAGE', '68': 'BUTTON_DROPDOWN', '12': 'DIALOG', '1': 'ALERT', '25': 'HEADER', '42': 'PASSWORD_TEXT', '63': 'TOOL_BAR', '62': 'TOGGLE_BUTTON', '15': 'END_NOTE', '54': 'SLIDER', '74': 'IMAGE_MAP', '73': 'FORM', '31': 'LAYERED_PANE', '16': 'FILE_CHOOSER', '4': 'CHECK_BOX'}
def macro(documentevent=None):  # 引数は文書のイベント駆動用。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	outputs = [("ComponentWindow",)]
	componentwindow = controller.ComponentWindow  # コントローラーのアトリビュートからコンポーネントウィンドウを取得。
	getAccessibleChildren((), componentwindow, outputs)
	outputs.append(("",))
	outputs.append(("ContainerWindow",))
	frame = controller.getFrame()  # フレームを取得。
	containerwindow = frame.getContainerWindow()
	getAccessibleChildren((), containerwindow, outputs)
	sheet = getNewSheet(doc, "SubWindows")  # 連番名の新規シートの取得。OnTitleChanged→OnModifyChangedが呼ばれてしまう。
	rowsToSheet(sheet["A1"], outputs)
	controller.setActiveSheet(sheet)  # 新規シートをアクティブにする。
def getAccessibleChildren(head, accessiblecontext, outputs):
	outputs.append(head+("AccessibleRole", "X", "Y", "Width", "Height", "X onScreen", "Y onScreen"))
	accessiblecontext = accessiblecontext.getAccessibleContext()
	for i in range(accessiblecontext.getAccessibleChildCount()):	
		accessiblechild = accessiblecontext.getAccessibleChild(i)
		childaccessiblecontext = accessiblechild.getAccessibleContext()
		bounds = childaccessiblecontext.getBounds()
		locationonscreen = childaccessiblecontext.getLocationOnScreen()
		accessiblerole = childaccessiblecontext.getAccessibleRole()
		outputrow = "{}={}".format(DIC_ACCESSIBLEROLE[str(accessiblerole)], accessiblerole), bounds.X, bounds.Y, bounds.Width, bounds.Height, locationonscreen.X, locationonscreen.Y
		outputs.append(head+outputrow)	
		if accessiblerole==51:  # SCROLL_PANEの時。
			getAccessibleChildren(("",)*7, accessiblechild, outputs)
def getNewSheet(doc, sheetname):  # docに名前sheetnameのシートを返す。sheetnameがすでにあれば連番名を使う。
	cellflags = cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES
	sheets = doc.getSheets()  # シートコレクションを取得。
	c = 1  # 連番名の最初の番号。
	newname = sheetname
	while newname in sheets: # 同名のシートがあるとき。sheets[sheetname]ではFalseのときKeyErrorになる。
		if not sheets[newname].queryContentCells(cellflags):  # シートが未使用のとき
			return sheets[newname]  # 未使用の同名シートを返す。
		newname = "{}{}".format(sheetname, c)  # 連番名を作成。
		c += 1 
	index = len(sheets)  # 最終シートにする。
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
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
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
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
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
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
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
	macro()
