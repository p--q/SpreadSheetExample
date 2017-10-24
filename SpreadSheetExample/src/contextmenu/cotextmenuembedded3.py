#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import date, timedelta 
import os
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.lang import Locale
global XSCRIPTCONTEXT  # PyDevでUndefined variableといわれないためだけのコード。
def macro(documentevent):  # 引数はcom.sun.star.document.DocumentEvent Struct。
	doc = documentevent.Source  # ドキュメントの取得。
	sheets = doc.getSheets()  # シートコレクションの取得。
	sheet = sheets[0]  # インデックス0のシートを取得。
	sheet.clearContents(511)  # シートのセルの内容をすべてを削除。
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.registerContextMenuInterceptor(ContextMenuInterceptor())  # コントローラにContextMenuInterceptorを登録する。
	global getFormatKey  # ScriptingURLで呼び出されたマクロで使うためにグローバル変数にする。
	getFormatKey = formatkeyCreator(doc)  # ドキュメントのformatkeyを返す関数を取得。
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self):
		filename = os.path.basename(__file__)  # このファイル名を取得。フルパスは"vnd.sun.star.tdoc:/4/Scripts/python/filename.py"というように番号(LibreOfficeバージョン番号?)が入ってしまう。
		self.baseurl = "vnd.sun.star.script:{}${}?language=Python&location=document".format(filename, "{}")  # ScriptingURLのbaseurlを取得。
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
	setDate(date.today())  # 今日の日付を取得。
def getYesterday():
	setDate(date.today()-timedelta(days=1))  # 昨日の日付を取得。
def getTomorrow():
	setDate(date.today()+timedelta(days=1))  # 明日の日付を取得。
def setDate(day):  # セルに日付を入力する。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	selection = doc.getCurrentSelection()  # セル範囲を取得。
	firstcell = getFirtstCell(selection)  # セル範囲の左上のセルを取得。
	firstcell.setFormula(day.isoformat())  # 日付の入力は年-月-日 または 月/日/年 にしないといけないらしい。
	firstcell.setPropertyValue("NumberFormat", getFormatKey("YYYY-MM-DD"))  # セルの書式を設定。	
def formatkeyCreator(doc):
	numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
# 	locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。	
	locale = Locale(Language="de", Country="DE")  # フォーマット一覧をくくる言語と国を設定。	
# 	locale = Locale(Language="en", Country="US")  # フォーマット一覧をくくる言語と国を設定。
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
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
g_exportedScripts = macro, #マクロセレクター(ScriptingURLで呼び出すための設定は不要)に限定表示させる関数をタプルで指定。
