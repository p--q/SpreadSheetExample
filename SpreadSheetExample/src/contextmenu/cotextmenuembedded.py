#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
def macro(documentevent):
	doc = documentevent.Source  # ドキュメントの取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.registerContextMenuInterceptor(ContextMenuInterceptor())  # コントローラにContextMenuInterceptorを登録する。
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューに割り込ませる。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 引数はContextMenuExecuteEvent Struct。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # すでにあるコンテクストメニュー(アクショントリガーコンテナ)を取得。
		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer")  # サブメニューにするアクショントリガーコンテナをインスタンス化。
		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Content", "CommandURL": ".uno:HelpIndex", "HelpURL": "5401"})  # アクショントリガーコンテナのインデックス0にアクショントリガーを挿入。
		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Tips", "CommandURL": ".uno:HelpTip", "HelpURL": "5404"})  # アクショントリガーコンテナのインデックス1にアクショントリガーを挿入。
		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "Help", "CommandURL": ".uno:HelpMenu", "HelpURL": "5410", "SubContainer": submenucontainer})  # アクショントリガーコンテナのインデックス0にアクショントリガーを挿入。サブメニューも挿入。
		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})  # アクショントリガーコンテナのインデックス1にセパレーターを挿入。
		return EXECUTE_MODIFIED  
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
