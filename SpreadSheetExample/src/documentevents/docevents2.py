#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from itertools import zip_longest
from com.sun.star.sheet import CellFlags as cf # 定数
FLG = True  # ドキュメントイベントを有効にするフラグ。
ORDER = 0  # 呼び出された順番。
RESULTS = [("Order", "Event Name")]
DIC = {'OnStartApp': ("アプリケーションの開始時", "Start Application"),\
		'OnCloseApp': ("アプリケーション終了時", "Close Application"),\
		'OnCreate': ("文書作成時", "Document created"),\
		'OnNew': ("新規文書の開始時", "New Document"),\
		'OnLoadFinished': ("文書の読み込み終了時", "Document loading finished"),\
		'OnLoad': ("文書を開いた時", "Open Document"),\
		'OnPrepareUnload': ("文書が閉じられる直前", "Document is going to be closed"),\
		'OnUnload': ("文書を閉じた時", "Document closed"),\
		'OnSave': ("文書を保存する時", "Save Document"),\
		'OnSaveDone': ("文書を保存した時", "Document has been saved"),\
		'OnSaveFailed': ("文書の保存が失敗した時", "Saving of document failed"),\
		'OnSaveAs': ("別名で保存する時", "Save Document As"),\
		'OnSaveAsDone': ("文書を別名で保存した時", "Document has been saved as"),\
		'OnSaveAsFailed': ("'別名で保存'が失敗した時", "'Save as' has failed"),\
		'OnCopyTo': ("文書の保存もしくはエクスポート", "Storing or exporting copy of document"),\
		'OnCopyToDone': ("文書のコピーを作った時", "Document copy has been created"),\
		'OnCopyToFailed': ("文書のコピーが失敗した時", "Creating of document copy failed"),\
		'OnFocus': ("文書を有効化した時", "Activate Document"),\
		'OnUnfocus': ("文書を無効化した時", "Deactivate Document"),\
		'OnPrint': ("文書の印刷時", "Print Document"),\
		'OnViewCreated': ("ビューの作成時", "View created"),\
		'OnPrepareViewClosing': ("ビューが閉じられる直前", "View is going to be closed"),\
		'OnViewClosed': ("ビューを閉じた時", "View closed"),\
		'OnModifyChanged': ("'変更'ステータス変更時", "'Modified' status was changed"),\
		'OnTitleChanged': ("文書のタイトルを変更した時", "Document title changed"),\
		'OnVisAreaChanged': ("OnStorageChanged", "OnStorageChanged"),\
		'OnModeChanged': ("OnModeChanged", "OnStorageChanged"),\
		'OnStorageChanged': ("OnStorageChanged", "OnStorageChanged")}  # イベント名の辞書。	
def macro(arg=None):  # 引数は文書のイベント駆動用。  
	if FLG:
		global ORDER  # RESULTSはリストなのでglobalに指定しなくても書き込める。
		if arg.typeName=="com.sun.star.document.DocumentEvent":  # 引数がDocumentEventのとき。
	# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグ用。
			eventname = arg.EventName  # イベント名を取得。
			RESULTS.append((ORDER, eventname, *DIC[eventname]))  # 呼び出され順、イベント名、イベントの日本語UI名、英語UI名、arg.Souruce(イベントを発火させたオブジェクト)はドキュメントモデル。。
		else:  # 引数がDocumentEvent以外のときはない?
			RESULTS.append((ORDER, str(arg)))
		ORDER += 1
def output(documentevent=None):  # Calcのシートにイベントの呼び出し順を書き出す。
	macro(documentevent)  # このイベント自身の呼び出し順を取得。
	global FLG
	FLG = False  # フラグを倒してドキュメントイベントの結果をRESULTSに取得しないようにする。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	sheet = getNewSheet(doc, "Events")  # 連番名の新規シートの取得。OnTitleChanged→OnModifyChangedが呼ばれてしまう。
	rowsToSheet(sheet[0, 0], RESULTS)  # 結果をシートに出力。シートを書き込むときのイベントを取得しないためにこの時点のRESULTSのコピーを渡す。
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.setActiveSheet(sheet)  # 新規シートをアクティブにする。
	FLG = True
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
	index = len(sheets)  # 最終シートにする。
# 	index = 0  # 先頭シートにする。
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
g_exportedScripts = macro, output #マクロセレクターに限定表示させる関数をタプルで指定。	
