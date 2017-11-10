#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from itertools import zip_longest
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.sheet import CellFlags as cf # 定数
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。ただしマウスハンドラはフリーズするので直接pydevを書き込んだほうがよい。
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
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。
	root = configreader("/org.openoffice.TypeDetection.Filter/Filters")  # コンフィギュレーションのルートを取得。
	props = "UIName", "UIComponent", "DocumentService"  # 取得するプロパティ名のタプル。
	outputs = []  # 出力行のリスト。
	for childname in root.getElementNames():  # 子ノードの名前のタプルを取得。ノードオブジェクトの直接取得はできない模様。
		uiname, uicomponent, documentservice = root[childname].getPropertyValues(props)  # フィルターのプロパティを取得。
		if uicomponent and documentservice=="com.sun.star.sheet.SpreadsheetDocument":  # スプレッドシートを処理するダイアログがあるもののみ抽出。
			datarow = uiname, childname, uicomponent, documentservice
			outputs.append(datarow)
	datarows = []  # シートに書き込む行のリスト。
	datarow = props[0], "FilterName", *props[1:]  # ヘッダー行。右辺のタプルのアンパックはPython3.5以上でのみ可能。
	datarows.append(datarow)
	sheetname = "UIComponents"  # UIComponentsプロパティがあるノードのみそれでソートして出力する。
	outputs.sort(key=lambda r: r[2])  # 行の列インデックス2。つまりUIComponentでソートする。
	datarows.extend(outputs)  # データ行を追加。	
	sheet = getNewSheet(doc, sheetname)  # 新規シートの取得。
	rowsToSheet(sheet, datarows)  # datarowsをシートに書き出し。		
	annotations = sheet.getAnnotations()  # シートのセル注釈コレクションを取得。
	txt = "Export the sheet in the UIName format to the home directory."
	annotations.insertNew(sheet["A1"].getCellAddress(), txt)  # セル注釈を挿入。
	txt = "Expand the return value of the FilterOptionsDialog."
	annotations.insertNew(sheet["B1"].getCellAddress(), txt)  # セル注釈を挿入。
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.setActiveSheet(sheet)  # シートをアクティブにする。	
	r = len(datarows) + 1  # データの最終行の2つ下の行のインデックス。
	args = ctx, smgr, doc, configreader, r
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(args))  # マウスハンドラをコントローラに設定。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler): # マウスハンドラ
	def __init__(self, args):
		self.args = args
# 	@enableRemoteDebugging  # ダブルクリニックで2回呼ばれる。2回目はGUIで操作できないときあり。
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		ctx, smgr, doc, configreader, r = self.args
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
		# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					sheet = target.getSpreadsheet()  # ターゲットがあるシートを取得。
					celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
					if 0<celladdress.Row<r and sheet[0, celladdress.Column].getString()=="FilterName":  # 行ヘッダーがFilterNameの列のとき。
						filtername = target.getString()  # ターゲットセルの文字列を取得。
						uicomponent = configreader("/org.openoffice.TypeDetection.Filter/Filters/{}".format(filtername)).getPropertyValue("UIComponent")  # フィルター名からUIComponent名を取得。
						filteroptiondialog = smgr.createInstanceWithContext(uicomponent, ctx)  # UIコンポーネントをインスタンス化。
						filteroptiondialog.setSourceDocument(doc)  # 変換元のドキュメントを設定。
						propertyvalues = PropertyValue(Name="FilterName", Value=filtername),  # 複数のフィルターに対応しているUIComponentはFilterNameを設定しないといけない。
						filteroptiondialog.setPropertyValues(propertyvalues)  # XPropertyAccessインターフェイスのメソッド。
						if filteroptiondialog.execute()==1:  # フィルターのオプションダイアログを表示。
							propertyvalues = filteroptiondialog.getPropertyValues()  # 戻り値はPropertyValue Structのタプル。XPropertyAccessインターフェイスのメソッド。
							outputs = []  # 出力行のリスト。
							expandPropertyValueStructs(outputs, propertyvalues, 0)  # PropertyValue Structを展開する。
							header = filtername,
							datarows = [header]
							datarows.extend(outputs)
							sheet.getRows().insertByIndex(r, len(datarows)+1)  # 行インデックスrに挿入するデータの行数の行を挿入する。
							rowsToSheet(sheet[r, 0], datarows)
						return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。
	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
		return True  # Trueでイベントを次のハンドラに渡す。
	def disposing(self, eventobject):
		pass	
def expandPropertyValueStructs(outputs, structs, h):  # outputs: 出力行のリスト、structs: PropertyValuesのタプル、h: 出力階層。
	flg = True  # 先頭行のフラグ。
	for struct in structs:
		ind = [""]*h
		if flg and h>0:  # 出力階層1以上で先頭行のとき。
			ind[-1] = "Value"
			flg = False
		name = "Name", struct.Name
		ind.extend(name)  # 戻り値はない。
		outputs.append(ind)
		v = struct.Value
		if isinstance(v, tuple):  # Valueがタプルの時は入れ子になっているので再帰呼出し。
			expandPropertyValueStructs(outputs, v, h+1)
		else:
			if isinstance(v, bool) or not isinstance(v, int):  # vがint以外の時。bool型はintに含まれるので注意。
				v = str(v)  # setDataArrayではint以外はStringしか代入できないのでStringに変換する。
			ind = [""]*h
			value = "Value", v
			ind.extend(value)  # 戻り値はない。
			outputs.append(ind)
def rowsToSheet(cellrange, datarows):  # 引数のセル範囲を左上端にして一括書き込みして列幅を最適化する。datarowsはタプルのタプル。
	datarows = tuple(zip(*zip_longest(*datarows, fillvalue="")))  # 一番長い行の長さに合わせて空文字を代入。
	sheet = cellrange.getSpreadsheet()  # セル範囲のあるシートを取得。
	cellcursor = sheet.createCursorByRange(cellrange)  # セル範囲のセルカーサーを取得。
	cellcursor.collapseToSize(len(datarows[0]), len(datarows))  # (列、行)で指定。セルカーサーの範囲をdatarowsに合せる。
	cellcursor.setDataArray(datarows)  # セルカーサーにdatarowsを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。		
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