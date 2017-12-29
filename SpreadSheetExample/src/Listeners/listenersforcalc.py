#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import inspect
from datetime import datetime
from com.sun.star.frame import XTerminateListener
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.frame import XTitleChangeListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.frame import XBorderResizeListener
from com.sun.star.document import XDocumentEventListener
from com.sun.star.awt import XKeyHandler
from com.sun.star.document import XEventListener
from com.sun.star.util import XModifyListener
from com.sun.star.view import XPrintJobListener
from com.sun.star.view.PrintableState import JOB_STARTED, JOB_COMPLETED, JOB_SPOOLED, JOB_ABORTED, JOB_FAILED, JOB_SPOOLING_FAILED  # enum 
from com.sun.star.document import XStorageChangeListener
from com.sun.star.util import XChangesListener
from com.sun.star.chart import XChartDataChangeEventListener
from com.sun.star.chart.ChartDataChangeType import ALL, DATA_RANGE, COLUMN_INSERTED, ROW_INSERTED, COLUMN_DELETED, ROW_DELETED  # enum
def macro(documentevent=None):  # 引数は文書のイベント駆動用。OnStartAppでもDocumentEventが入るがSourceはNoneになる。# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	path = doc.getURL() if __file__.startswith("vnd.sun.star.tdoc:") else __file__  # このスクリプトのパス。fileurlで返ってくる。
	thisscriptpath = unohelper.fileUrlToSystemPath(path)
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	desktop = XSCRIPTCONTEXT.getDesktop()
	desktop_terminatelistener = TerminateListener(dirpath, "desktop_terminatelistener")  # TerminateListener
	desktop.addTerminateListener(desktop_terminatelistener)  # 使い終わったらremoveしないといけない。
	controller = doc.getCurrentController()
	frame = controller.getFrame()
	frame.addFrameActionListener(FrameActionListener(dirpath, "frame_frameactionlistener")) 
	frame.addCloseListener(CloseListener(dirpath, "frame_closelistener"))
	frame.addTitleChangeListener(TitleChangeListener(dirpath, "frame_titlechangelistener"))
	controller.addActivationEventListener(ActivationEventListener(dirpath, "controller_activationeventlistener"))
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(dirpath, "controller_enhancedmouseclickhandler"))
	controller.addSelectionChangeListener(SelectionChangeListener(dirpath, "controller_selectionchangelistener"))
	controller.addBorderResizeListener(BorderResizeListener(dirpath, "controller_borderresizelistener"))
	controller.addTitleChangeListener(TitleChangeListener(dirpath, "controller_titlechangelistener"))	
	controller.addKeyHandler(KeyHandler(dirpath, "controller_keyhandler"))		
	doc.addDocumentEventListener(DocumentEventListener(dirpath, "doc_documenteventlistener"))	
	doc.addEventListener(EventListener(dirpath, "doc_eventlistener"))		
	doc.addModifyListener(ModifyListener(dirpath, "doc_modifylistener"))		
	doc.addPrintJobListener(PrintJobListener(dirpath, "doc_printjoblistener"))	
	doc.addStorageChangeListener(StorageChangeListener(dirpath, "doc_storagechangelistener"))	
	doc.addTitleChangeListener(TitleChangeListener(dirpath, "doc_titlechangelistener"))
	doc.addChangesListener(ChangesListener(dirpath, "doc_changelistener"))	
	sheet = controller.getActiveSheet()
	sheet.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "sheet_chartdatachangeeventlistener"))	
	sheet.addModifyListener(ModifyListener(dirpath, "sheet_modifylistener"))	
	cell = sheet["A1"]
	cell.addModifyListener(ModifyListener(dirpath, "cell_modifylistener"))
	cells = sheet["A2:C4"]
	cells.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "cells_chartdatachangeeventlistener"))	
	cells.addModifyListener(ModifyListener(dirpath, "cells_modifylistener"))	
class ChartDataChangeEventListener(unohelper.Base, XChartDataChangeEventListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
		enums = ALL, DATA_RANGE, COLUMN_INSERTED, ROW_INSERTED, COLUMN_DELETED, ROW_DELETED  # enum
		chartdatachangetypenames = "ALL", "DATA_RANGE", "COLUMN_INSERTED", "ROW_INSERTED", "COLUMN_DELETED", "ROW_DELETED"
		self.args = dirpath, name, zip(enums, chartdatachangetypenames)
	def chartDataChanged(self, chartdatachangeevent):
		dirpath, name, chartdatachangetypes = self.args
		chartdatachangetype = chartdatachangeevent.Type
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		for enum, chartdatachangetypename in chartdatachangetypes:
			if chartdatachangetype==enum:
				methodname = "_".join((name, inspect.currentframe().f_code.co_name))
				createLog(dirpath, methodname, "ChartDataChangeType: {}".format(chartdatachangetypename))
				return
	def disposing(self, eventobject):
		pass		
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def changesOccurred(self, changesevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Changes: {}".format(changesevent.Changes))	
	def disposing(self, eventobject):
		pass		
class StorageChangeListener(unohelper.Base, XStorageChangeListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def notifyStorageChange(self, document, storage):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Storage: {}".format(storage))	
	def disposing(self, eventobject):
		pass			
class PrintJobListener(unohelper.Base, XPrintJobListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
		enums = JOB_STARTED, JOB_COMPLETED, JOB_SPOOLED, JOB_ABORTED, JOB_FAILED, JOB_SPOOLING_FAILED  # enum
		printablestatenames = "JOB_STARTED", "JOB_COMPLETED", "JOB_SPOOLED", "JOB_ABORTED", "JOB_FAILED", "JOB_SPOOLING_FAILED"
		self.args = dirpath, name, zip(enums, printablestatenames)
	def printJobEvent(self, printjobevent):
		dirpath, name, printablestates = self.args
		printablestate = printjobevent.State
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		for enum, printablestatename in printablestates:
			if printablestate==enum:
				methodname = "_".join((name, inspect.currentframe().f_code.co_name))
				createLog(dirpath, methodname, "FrameAction: {}, Source: {}".format(printablestatename, printjobevent.Source))
				return
	def disposing(self, eventobject):
		pass		
class ModifyListener(unohelper.Base, XModifyListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def modified(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Source: {}".format(eventobject.Source))	
	def disposing(self, eventobject):
		pass		
class EventListener(unohelper.Base, XEventListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def notifyEvent(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "EventName: {}, Source: {}".format(eventobject.EventName, eventobject.Source))	
	def disposing(self, eventobject):
		pass		
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def documentEventOccured(self, documentevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "EventName: {}, Source: {}".format(documentevent.EventName, documentevent.Source))	
	def disposing(self, eventobject):
		pass	
class KeyHandler(unohelper.Base, XKeyHandler):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def keyPressed(self, keyevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "KeyCode: {}, KeyChar: {}, KeyFunc: {}, Modifiers: {}".format(keyevent.KeyCode, keyevent.KeyChar, keyevent.KeyFunc, keyevent.Modifiers))		
		return False
	def keyReleased(self, keyevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "KeyCode: {}, KeyChar: {}, KeyFunc: {}, Modifiers: {}".format(keyevent.KeyCode, keyevent.KeyChar, keyevent.KeyFunc, keyevent.Modifiers))		
		return False		
	def disposing(self, eventobject):
		pass			
class BorderResizeListener(unohelper.Base, XBorderResizeListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def borderWidthsChanged(self, obj, borderwidths):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Top: {}, Left: {}, Right: {}, Bottom: {}, Object: {}".format(borderwidths.Top, borderwidths.Left, borderwidths.Right, borderwidths.Bottom, obj))	
	def disposing(self, eventobject):
		pass		
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def selectionChanged(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Source: {}".format(eventobject.Source))	
	def disposing(self, eventobject):
		pass	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def mousePressed(self, enhancedmouseevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Buttons: {}, ClickCount: {}, PopupTrigger {}, Modifiers: {}, Target: {}".format(enhancedmouseevent.Buttons, enhancedmouseevent.ClickCount, enhancedmouseevent.PopupTrigger, enhancedmouseevent.Modifiers, enhancedmouseevent.Target))	
		return True
	def mouseReleased(self, enhancedmouseevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Buttons: {}, ClickCount: {}, PopupTrigger {}, Modifiers: {}, Target: {}".format(enhancedmouseevent.Buttons, enhancedmouseevent.ClickCount, enhancedmouseevent.PopupTrigger, enhancedmouseevent.Modifiers, enhancedmouseevent.Target))	
		return True
	def disposing(self, eventobject):
		pass
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def activeSpreadsheetChanged(self, activationevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Source: {}".format(activationevent.Source))	
	def disposing(self, eventobject):
		pass
class TitleChangeListener(unohelper.Base, XTitleChangeListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def titleChanged(self, titlechangedevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Title: {}, Source: {}".format(titlechangedevent.Title, titlechangedevent.Source))	
	def disposing(self, eventobject):
		pass
class CloseListener(unohelper.Base, XCloseListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def queryClosing(self, eventobject, getsownership):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "getsownership: {}, Source: {}".format(getsownership, eventobject.Source))	
	def notifyClosing(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "Source: {}".format(eventobject.Source))
	def disposing(self, eventobject):
		pass	
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def __init__(self, dirpath, name):
		enums = COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
		frameactionnames = "COMPONENT_ATTACHED", "COMPONENT_DETACHING", "COMPONENT_REATTACHED", "FRAME_ACTIVATED", "FRAME_DEACTIVATING", "CONTEXT_CHANGED", "FRAME_UI_ACTIVATED", "FRAME_UI_DEACTIVATING"
		self.args = dirpath, name, zip(enums, frameactionnames)
	def frameAction(self, frameactionevent):
		dirpath, name, frameactions = self.args
		frameaction = frameactionevent.Action
		for enum, frameactionname in frameactions:
			if frameaction==enum:
				methodname = "_".join((name, inspect.currentframe().f_code.co_name))
				createLog(dirpath, methodname, "FrameAction: {}, Source: {}".format(frameactionname, frameactionevent.Source))
				return
	def disposing(self, eventobject):
		pass
class TerminateListener(unohelper.Base, XTerminateListener):  # TerminateListener
	def __init__(self, dirpath, name):  # 出力先ディレクトリのパス、リスナーのインスタンス名。
		self.args = dirpath, name
	def queryTermination(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
		createLog(dirpath, methodname, "Source: {}".format(eventobject.Source))  # Sourceを出力。
	def notifyTermination(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
		createLog(dirpath, methodname, "Source: {}".format(eventobject.Source))  # Sourceを出力。
		desktop = eventobject.Source
		desktop.removeTerminateListener(self)  # TerminateListenerを除去。除去しなmethodname = inspect.currentframe().f_code.co_nameいとLibreOfficeのプロセスが残って起動できなくなる。
	def disposing(self, eventobject):
		pass
C = 10
TIMESTAMP = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
def createLog(dirpath, methodname, txt):  # 年月日T時分秒リスナーのインスタンス名_methodname.logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	global C
	filename = "".join((TIMESTAMP, "_", str(C), methodname, ".log"))
	C += 1
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。	
