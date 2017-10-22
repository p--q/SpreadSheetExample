#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui import ActionTriggerSeparatorType
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED
def macro(documentevent):
	doc = documentevent.Source  # ドキュメントの取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.registerContextMenuInterceptor(ContextMenuInterceptor())
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # com.sun.star.ui.ActionTriggerにsetPropertyValuesでは設定できない。エラーも出ない。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer
		addMenuentry = menuentryCreator(contextmenu)
		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer") 
		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Content", "CommandURL": ".uno:HelpIndex", "HelpURL": "5401"})
		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Tips", "CommandURL": ".uno:HelpTip", "HelpURL": "5404"})
		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "Help", "CommandURL": "slot:5410", "HelpURL": "5410", "SubContainer": submenucontainer})
		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})
		return EXECUTE_MODIFIED
def menuentryCreator(contextmenu):
	def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。
		menuentry = contextmenu.createInstance("com.sun.star.ui.{}".format(menutype))
		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。
	return addMenuentry
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
