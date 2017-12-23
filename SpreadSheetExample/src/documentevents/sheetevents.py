#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
# from itertools import zip_longest
# from com.sun.star.sheet import CellFlags as cf # 定数
# FLG = True  # ドキュメントイベントを有効にするフラグ。
# ORDER = 0  # 呼び出された順番。
# RESULTS = [("Order", "Event Name")]
# DIC = {'OnFocus': ("文書を有効化した時", "Activate Document"),\
# 		'OnUnfocus': ("文書を無効化した時", "Deactivate Document"),\
# 		'OnSelect': ("選択を変更した時", "Selection changed"),\
# 		'OnDoubleClick': ("ダブルクリックした時", "Double click"),\
# 		'OnRightClick': ("右クリックした時", "Right click"),\
# 		'OnChange': ("内容を変更した時", "Content changed"),\
# 		'OnCalculate': ("計算が完了した時", "Formulas calculated")}  # イベント名の辞書。	
def macro(arg=None):  # 引数は文書のイベント駆動用。  
# 	if FLG:
# 		global ORDER  # RESULTSはリストなのでglobalに指定しなくても書き込める。
# 	FLG = False
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()
# 	sheet[ORDER, 0].setString(str(arg))
	sheet[0, 0].setString(str(arg))
# 	FLG = True
# 		ORDER += 1
		
		
		
# 		if arg.typeName=="com.sun.star.document.DocumentEvent":  # 引数がDocumentEventのとき。
# 	# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグ用。
# 			eventname = arg.EventName  # イベント名を取得。
# 			RESULTS.append((ORDER, eventname, *DIC[eventname]))  # 呼び出され順、イベント名、イベントの日本語UI名、英語UI名、arg.Souruce(イベントを発火させたオブジェクト)はドキュメントモデル。。
# 		else:  # 引数がDocumentEvent以外のときはない?
# 			RESULTS.append((ORDER, str(arg)))
# 		ORDER += 1
# def output(documentevent=None):  # Calcのシートにイベントの呼び出し順を書き出す。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグ用。
# # 	macro(documentevent)  # このイベント自身の呼び出し順を取得。
# 	global FLG
# 	FLG = False  # フラグを倒してドキュメントイベントの結果をRESULTSに取得しないようにする。
# 	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
# 	sheet = getNewSheet(doc, "Events")  # 連番名の新規シートの取得。OnTitleChanged→OnModifyChangedが呼ばれてしまう。
# 	rowsToSheet(sheet[0, 0], RESULTS)  # 結果をシートに出力。シートを書き込むときのイベントを取得しないためにこの時点のRESULTSのコピーを渡す。
# 	controller = doc.getCurrentController()  # コントローラの取得。
# 	controller.setActiveSheet(sheet)  # 新規シートをアクティブにする。
# 	FLG = True
# def getNewSheet(doc, sheetname):  # docに名前sheetnameのシートを返す。sheetnameがすでにあれば連番名を使う。
# 	cellflags = cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES
# 	sheets = doc.getSheets()  # シートコレクションを取得。
# 	c = 1  # 連番名の最初の番号。
# 	newname = sheetname
# 	while newname in sheets: # 同名のシートがあるとき。sheets[sheetname]ではFalseのときKeyErrorになる。
# 		if not sheets[newname].queryContentCells(cellflags):  # シートが未使用のとき
# 			return sheets[sheetname]  # 未使用の同名シートを返す。
# 		newname = "{}{}".format(sheetname, c)  # 連番名を作成。
# 		c += 1	
# 	index = len(sheets)  # 最終シートにする。
# # 	index = 0  # 先頭シートにする。
# 	sheets.insertNewByName(newname, index)   # 新しいシートを挿入。同名のシートがあるとRuntimeExceptionがでる。
# 	if "Sheet1" in sheets:  # デフォルトシートがあるとき。
# 		if not sheets["Sheet1"].queryContentCells(cellflags):  # シートが未使用のとき
# 			del sheets["Sheet1"]  # シートを削除する。
# 	return sheets[newname]
# def rowsToSheet(cellrange, datarows):  # 引数のセル範囲を左上端にして一括書き込みして列幅を最適化する。datarowsはタプルのタプル。
# 	datarows = tuple(zip(*zip_longest(*datarows, fillvalue="")))  # 一番長い行の長さに合わせて空文字を代入。
# 	sheet = cellrange.getSpreadsheet()  # セル範囲のあるシートを取得。
# 	cellcursor = sheet.createCursorByRange(cellrange)  # セル範囲のセルカーサーを取得。
# 	cellcursor.collapseToSize(len(datarows[0]), len(datarows))  # (列、行)で指定。セルカーサーの範囲をdatarowsに合せる。
# 	cellcursor.setDataArray(datarows)  # セルカーサーにdatarowsを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
# 	cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。
# g_exportedScripts = macro, output #マクロセレクターに限定表示させる関数をタプルで指定。	
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。	
