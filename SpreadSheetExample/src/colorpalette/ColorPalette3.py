#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import glob
from itertools import zip_longest
from xml.etree import ElementTree
from com.sun.star.sheet import CellFlags as cf # 定数
from com.sun.star.table.CellHoriJustify import CENTER  # enum
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	pathsettingssingleton = ctx.getByName('/singletons/com.sun.star.util.thePathSettings')  # thePathSettings
	fileurls = pathsettingssingleton.getPropertyValue("Palette").split(";")  # Paletteへのパスを取得。セミコロン区切りで複数返ってくるのでリストにする。
	lst_socs = ["standard.soc", "chart-palettes.soc"]  # 出力順を決まっているファイル名。
	for fileurl in reversed(fileurls):  # ユーザーフォルダにある方を先に取得するため逆順にする。
		sheet = getNewSheet(doc, "Palette")  # レイヤー毎にシートを作成する。
		c = 0  # 出力列インデックス。
		palettepath = os.path.normpath(unohelper.fileUrlToSystemPath(fileurl))  # システムパスに変換。
		os.chdir(palettepath)  # socファイルのあるフォルダに移動。
		set_socs = set(glob.glob("*.soc"))  # socファイルのリストを集合にして取得。
		socs = lst_socs.copy()
		socs.extend(set_socs.difference(lst_socs))  # ファイルの順番を変更。
		xpath = './/draw:color'
		namespaces1 = {"draw": "{http://openoffice.org/2000/drawing}"}  # 名前空間の辞書。replace()で置換するのに使う。
		replaceWithValue1, replaceWithKey1 = createReplaceFunc(namespaces1)	
		namespaces2 = {"draw": "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"}  # drawはもうひとつの名前空間が割り当てられている。
		replaceWithValue2, replaceWithKey2 = createReplaceFunc(namespaces2)	
		for socname in socs:  # socファイルを取得。
			if os.path.exists(socname):
				tree = ElementTree.parse(socname)  # xmlの木を取得。
				xpath1 = replaceWithValue1(xpath)  # 名前空間1の辞書のキーを値に変換。
				nodes = tree.findall(xpath1)  # xpahのノードを取得。	
				replaceWithKey = replaceWithKey1  # 名前空間1を戻す関数。
				if not nodes:  # ノードが取得出来なかった時。
					xpath2 = replaceWithValue2(xpath)  # 名前空間を2に変える。
					nodes = tree.findall(xpath2)  # xpahのノードを取得。	
					replaceWithKey = replaceWithKey2  # 名前空間2を戻す関数。
				if nodes:  # ノードが取得出来た時。
					outputs = getAttrib(nodes, replaceWithKey)
					rowsToSheet(sheet[2, c], outputs)  # シートに書き込む。
					sheet[1, c].setString(socname)  # socファイル名を出力。
					sheet[1, c:c+5].merge(True)
					sheet[1, c:c+5].setPropertyValue("HoriJustify", CENTER)
					rows = sheet[2:2+len(outputs), c+3].getDataArray()  # 色の10進数を取得。
					for i, row in enumerate(rows):
						if row[0]!="":  # 0の時もあるので空文字かどうかで判断する。
							sheet[2+i, c].setPropertyValue("CellBackColor", int(row[0]))  # floatで返ってくるのでintにしないといけない。律速。
					c += 5
		sheet["A1"].setString(palettepath)
		sheet["A1:H1"].merge(True)
def getAttrib(nodes, replaceWithKey):
	outputs = []
	c = 0  # 行カウンタ。
	for node in nodes:  # 取得した各ノードについて。
		name, color = "", ""
		for key, val in node.items():  # ノードの各属性について。
			attrib = replaceWithKey(key)  # 名前空間の辞書の値をキーに変換。
			if attrib=="draw:name":
				name = val
			elif attrib=="draw:color":
				color = val.upper().replace("#", "0x")  # Pythonの16進数にする。#を0xに変換する。
		if name:  # 色名が取得出来ている時。
			if c==12:  # 12行ずつ空行を挿入。
				outputs.append(("",))
				c = 0
			outputs.append(("", name, color, int(color, 16)))  # 出力行に追加。
			c += 1
	return outputs
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
def getNewSheet(doc, sheetname):  # docに名前sheetnameのシートを返す。sheetnameがすでにあれば連番名を使う。
	cellflags = cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES
	sheets = doc.getSheets()  # シートコレクションを取得。
	c = 1  # 連番名の最初の番号。
	newname = sheetname
	while newname in sheets: # 同名のシートがあるとき。sheets[newname]ではFalseのときKeyErrorになる。
		if not sheets[newname].queryContentCells(cellflags):  # シートが未使用のとき
			return sheets[newname]  # 未使用の同名シートを返す。
		newname = "{}{}".format(sheetname, c)  # 連番名を作成。
		c += 1 
	index = len(sheets)  # 最終シートにする。
#  index = 0  # 先頭シートにする。
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
	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。行幅は限定サれない。  	
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
	#  doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
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
	