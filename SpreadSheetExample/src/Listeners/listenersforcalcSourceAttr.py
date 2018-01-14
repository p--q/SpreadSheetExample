#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import inspect
import platform
from datetime import datetime
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import XKeyHandler
from com.sun.star.awt import XTopWindowListener
from com.sun.star.awt import Key  # 定数
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view import XPrintJobListener
from com.sun.star.view.PrintableState import JOB_STARTED, JOB_COMPLETED, JOB_SPOOLED, JOB_ABORTED, JOB_FAILED, JOB_SPOOLING_FAILED  # enum 
from com.sun.star.util import XCloseListener
from com.sun.star.util import XModifyListener
from com.sun.star.util import XChangesListener
from com.sun.star.frame import XTerminateListener
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame import XTitleChangeListener
from com.sun.star.frame import XBorderResizeListener
from com.sun.star.frame.FrameAction import COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
from com.sun.star.document import XDocumentEventListener
from com.sun.star.document import XEventListener
from com.sun.star.document import XStorageChangeListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.chart import XChartDataChangeEventListener
from com.sun.star.chart.ChartDataChangeType import ALL, DATA_RANGE, COLUMN_INSERTED, ROW_INSERTED, COLUMN_DELETED, ROW_DELETED  # enum
def macro(documentevent=None):  # 引数は文書のイベント駆動用。OnStartAppでもDocumentEventが入るがSourceはNoneになる。# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	desktop = XSCRIPTCONTEXT.getDesktop()  # デスクトップの取得。
	path = doc.getURL() if __file__.startswith("vnd.sun.star.tdoc:") else __file__  # このスクリプトのパス。fileurlで返ってくる。埋め込みマクロの時は埋め込んだドキュメントのURLで代用する。
	thisscriptpath = unohelper.fileUrlToSystemPath(path)  # fileurlをsystempathに変換。
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	listeners = {}
	listeners["desktop_terminatelistener"] = TerminateListener(dirpath, "desktop_terminatelistener")
	listeners["desktop_frameactionlistener"] = FrameActionListener(dirpath, "desktop_frameactionlistener")
	desktop.addTerminateListener(listeners["desktop_terminatelistener"])  # TerminateListener
	desktop.addFrameActionListener(listeners["desktop_frameactionlistener"])  # FrameActionListener
	controller = doc.getCurrentController()  # コントローラーの取得。
	frame = controller.getFrame()  # フレームの取得。
	listeners["frame_frameactionlistener"] = FrameActionListener(dirpath, "frame_frameactionlistener")
	listeners["frame_closelistener"] = CloseListener(dirpath, "frame_closelistener")
	listeners["frame_titlechangelistener"] = TitleChangeListener(dirpath, "frame_titlechangelistener", frame)
	frame.addFrameActionListener(listeners["frame_frameactionlistener"])  # FrameActionListener 
	frame.addCloseListener(listeners["frame_closelistener"])  # CloseListener
	frame.addTitleChangeListener(listeners["frame_titlechangelistener"])  # TitleChangeListener
	containerwindow = frame.getContainerWindow()  # フレームのコンテナウィンドウの取得。
	listeners["containerwindow_topwindowlistener"] = TopWindowListener(dirpath, "containerwindow_topwindowlistener")
	containerwindow.addTopWindowListener(listeners["containerwindow_topwindowlistener"])
	controller.addActivationEventListener(ActivationEventListener(dirpath, "controller_activationeventlistener", controller))  # ActivationEventListener
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(dirpath, "controller_enhancedmouseclickhandler", controller))  # EnhancedMouseClickHandler
	controller.addSelectionChangeListener(SelectionChangeListener(dirpath, "controller_selectionchangelistener", controller))  # SelectionChangeListener
	controller.addBorderResizeListener(BorderResizeListener(dirpath, "controller_borderresizelistener", controller))  # BorderResizeListener
	controller.addTitleChangeListener(TitleChangeListener(dirpath, "controller_titlechangelistener", controller))  # TitleChangeListener	
	controller.addKeyHandler(KeyHandler(dirpath, "controller_keyhandler", controller))  # KeyHandler		
	doc.addDocumentEventListener(DocumentEventListener(dirpath, "doc_documenteventlistener", doc, desktop, frame, containerwindow, listeners))  # DocumentEventListener	
	doc.addEventListener(EventListener(dirpath, "doc_eventlistener", doc))  # EventListener	
	doc.addModifyListener(ModifyListener(dirpath, "doc_modifylistener", doc))  # ModifyListener		
	doc.addPrintJobListener(PrintJobListener(dirpath, "doc_printjoblistener", doc))  # PrintJobListener	
	doc.addStorageChangeListener(StorageChangeListener(dirpath, "doc_storagechangelistener", doc))  # StorageChangeListener	
	doc.addTitleChangeListener(TitleChangeListener(dirpath, "doc_titlechangelistener", doc))  # TitleChangeListener
	doc.addChangesListener(ChangesListener(dirpath, "doc_changelistener", doc))  # ChangesListener	
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	sheet.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "sheet_chartdatachangeeventlistener", sheet))  # ChartDataChangeEventListener	
	sheet.addModifyListener(ModifyListener(dirpath, "sheet_modifylistener", sheet))  # ModifyListener	
	cell = sheet["A1"]  # セルの取得。
	cell.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "cell_chartdatachangeeventlistener", cell))  # ChartDataChangeEventListener		
	cell.addModifyListener(ModifyListener(dirpath, "cell_modifylistener", cell))  # ModifyListener
	cells = sheet["A2:C4"]  # セル範囲の取得。
	cells.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "cells_chartdatachangeeventlistener", cells))  # ChartDataChangeEventListener	
	cells.addModifyListener(ModifyListener(dirpath, "cells_modifylistener", cells))  # ModifyListener	
class TopWindowListener(unohelper.Base, XTopWindowListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def windowOpened(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))		
	def windowClosing(self, eventobject):  # 呼ばれない。いつ呼ばれる?
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))			
	def windowClosed(self, eventobject):  # フレームが閉じた後デスクトップが閉じる前に呼ばれる。
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
	def windowMinimized(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))		
	def windowNormalized(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
	def windowActivated(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))		
	def windowDeactivated(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
class ChartDataChangeEventListener(unohelper.Base, XChartDataChangeEventListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
		enums = ALL, DATA_RANGE, COLUMN_INSERTED, ROW_INSERTED, COLUMN_DELETED, ROW_DELETED  # enum
		chartdatachangetypenames = "ALL", "DATA_RANGE", "COLUMN_INSERTED", "ROW_INSERTED", "COLUMN_DELETED", "ROW_DELETED"
		self.args = dirpath, name, zip(enums, chartdatachangetypenames)
	def chartDataChanged(self, chartdatachangeevent):
		dirpath, name, chartdatachangetypes = self.args
		chartdatachangetype = chartdatachangeevent.Type
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		for enum, chartdatachangetypename in chartdatachangetypes:
			if chartdatachangetype==enum:
				filename = "_".join((name, inspect.currentframe().f_code.co_name, chartdatachangetypename))  # ChartDataChangeType名を追加。
				createLog(dirpath, filename, "ChartDataChangeType: {}\nSource: {}".format(chartdatachangetypename, chartdatachangeevent.Source))
				return
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeChartDataChangeEventListener(self)		
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def changesOccurred(self, changesevent):
		dirpath, name = self.args
		base = changesevent.Base
		if base.supportsService("com.sun.star.sheet.SpreadsheetDocument"):  # ドキュメントの時
			basetxt = "Base URL: {}".format(__file__)  # ドキュメントのURLを取得。
		else:
			basetxt = "Base: {}".format(base)	
		txts = [basetxt]  # ログファイルに出力する行のリスト。	
		changes = changesevent.Changes
		for change in changes:
			txts.append("Accessor: {}".format(change.Accessor))
			for element in change.Element:
				if hasattr(element, "Name") and hasattr(element, "Value"):
					propertyname, propertyvalue = element.Name, element.Value
					if "Color" in propertyname:  # 色の時は16進数で出力する。
						propertyvalue = hex(propertyvalue)
					txts.append("{}: {}".format(propertyname, propertyvalue))
			replacedelement = getStringAddressFromCellRange(change.ReplacedElement)  # 変更対象オブジェクトから文字列アドレスを取得する。
			replacedelement = replacedelement or change.ReplaceElement  # 文字列アドレスを取得できないオブジェクトの時はオブジェクトをそのまま文字列にする。
			txts.append("ReplacedElement: {}".format(replacedelement))	
		txts.append("Source: {}".format(changesevent.Source))		
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "\n".join(txts))
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeChangesListener(self)	
class StorageChangeListener(unohelper.Base, XStorageChangeListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def notifyStorageChange(self, document, storage):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Storage: {}".format(storage))	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeStorageChangeListener(self)		
class PrintJobListener(unohelper.Base, XPrintJobListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
		enums = JOB_STARTED, JOB_COMPLETED, JOB_SPOOLED, JOB_ABORTED, JOB_FAILED, JOB_SPOOLING_FAILED  # enum
		printablestatenames = "JOB_STARTED", "JOB_COMPLETED", "JOB_SPOOLED", "JOB_ABORTED", "JOB_FAILED", "JOB_SPOOLING_FAILED"
		self.args = dirpath, name, zip(enums, printablestatenames)
	def printJobEvent(self, printjobevent):
		dirpath, name, printablestates = self.args
		printablestate = printjobevent.State
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		for enum, printablestatename in printablestates:
			if printablestate==enum:
				filename = "_".join((name, inspect.currentframe().f_code.co_name, printablestatename))  # State名も追加。
				createLog(dirpath, filename, "PrintableState: {}, Source: {}".format(printablestatename, printjobevent.Source))
				return
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removePrintJobListener(self)	
class ModifyListener(unohelper.Base, XModifyListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def modified(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeModifyListener(self)  # 最後に実行しないとクラッシュする。	
class EventListener(unohelper.Base, XEventListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def notifyEvent(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name, eventobject.EventName))  # イベント名も追加。
		createLog(dirpath, filename, "EventName: {}, Source: {}".format(eventobject.EventName, eventobject.Source))	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeEventListener(self)	
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, dirpath, name, subj, desktop, frame, containerwindow, listeners):
		self.subj = subj
		self.args = dirpath, name, desktop, frame, containerwindow, listeners
	def documentEventOccured(self, documentevent):
		dirpath, name, desktop, frame, containerwindow, listeners = self.args
		eventname = documentevent.EventName
		filename = "_".join((name, inspect.currentframe().f_code.co_name, eventname))  # イベント名も追加。
		if eventname=="OnUnload":  # ドキュメントを閉じてもdisposeされないデスクトップ、フレーム、コンテナウィンドウにつけたリスナーを除去する。
			desktop.removeTerminateListener(listeners["desktop_terminatelistener"])  # TerminateListener
			desktop.removeFrameActionListener(listeners["desktop_frameactionlistener"])  # FrameActionListener
			frame.removeFrameActionListener(listeners["frame_frameactionlistener"])  # FrameActionListener 
			frame.removeCloseListener(listeners["frame_closelistener"])  # CloseListener
			frame.removeTitleChangeListener(listeners["frame_titlechangelistener"])  # TitleChangeListener
			containerwindow.removeTopWindowListener(listeners["containerwindow_topwindowlistener"])  # TopWindowListener
			filename = "_".join((name, inspect.currentframe().f_code.co_name, eventname, "RemoveListeners"))  # イベント名も追加。
		createLog(dirpath, filename, "EventName: {}, Source: {}".format(eventname, documentevent.Source))	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeDocumentEventListener(self)
class KeyHandler(unohelper.Base, XKeyHandler):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.keycodes = {
			Key.DOWN: "DOWN",
			Key.UP: "UP",
			Key.LEFT: "LEFT", 
			Key.RIGHT: "RIGHT", 
			Key.HOME: "HOME", 
			Key.END: "END",
			Key.RETURN: "RETURN",
			Key.ESCAPE: "ESCAPE",
			Key.TAB: "TAB",		
			Key.BACKSPACE: "BACKSPACE", 
			Key.SPACE: "SPACE", 
			Key.DELETE: "DELETE"
		}  # キーは定数。特殊文字を文字列に置換する。
		self.args = dirpath, name	
	def keyPressed(self, keyevent):
		keychar = self._keycharToText(keyevent)
		self._createLogFile(keyevent, keychar, inspect.currentframe().f_code.co_name)
		return False
	def keyReleased(self, keyevent):
		keychar = "" if platform.system()=="Windows" else self._keycharToText(keyevent)  # Windowsの時日本語入力ではKeyCharを使うとすべて文字化けするので使わない。
		self._createLogFile(keyevent, keychar, inspect.currentframe().f_code.co_name)
		return False		
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeKeyHandler(self)	
	def _keycharToText(self, keyevent):
		keycode = keyevent.KeyCode
		keychar = ""  # KeyCharが特殊文字の場合はその後のテキストが表示されないときがあるので書き込まない。
		if keycode in self.keycodes:  # self.keycodesにキーがある特殊文字は文字列に置換する。
			keychar = self.keycodes[keycode]
		elif 255<keycode<267 or 511<keycode<538:  # 数値かアルファベットの時
			keychar = keyevent.KeyChar.value
		return keychar
	def _createLogFile(self, keyevent, keychar, methodname):
		dirpath, name = self.args
		if keychar:
			filename = "_".join((name, methodname, keychar))
			txt = "KeyCode: {}, KeyChar: {}, KeyFunc: {}, Modifiers: {}".format(keyevent.KeyCode, keychar, keyevent.KeyFunc, keyevent.Modifiers)
		else:
			filename = "_".join((name, methodname))
			txt = "KeyCode: {}, KeyFunc: {}, Modifiers: {}".format(keyevent.KeyCode, keyevent.KeyFunc, keyevent.Modifiers)			
		createLog(dirpath, filename, txt)
class BorderResizeListener(unohelper.Base, XBorderResizeListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def borderWidthsChanged(self, obj, borderwidths):
		dirpath, name = self.args
		if obj.supportsService("com.sun.star.sheet.SpreadsheetView"):  # objがコントローラーの時。
			cellrangeaddressconversion = obj.getModel().createInstance("com.sun.star.table.CellRangeAddressConversion")  # ドキュメントからCellRangeAddressConversionを取得。
			cellrangeaddressconversion.Address = obj.getVisibleRange()  # 表示されているセル範囲のCellRangeAddressを取得。
			txt = "Visible Range: {}".format(cellrangeaddressconversion.PersistentRepresentation)  # 表示されているセル範囲の文字列アドレスの取得。
		else:
			txt = "Top: {}, Left: {}, Right: {}, Bottom: {}, Object: {}".format(borderwidths.Top, borderwidths.Left, borderwidths.Right, borderwidths.Bottom, obj)
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, txt)	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeBorderResizeListener(self)		
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def selectionChanged(self, eventobject):
		dirpath, name = self.args
		txt = ""
		source = eventobject.Source
		if source.supportsService("com.sun.star.sheet.SpreadsheetView"):  # sourceがコントローラーのとき
			selection = source.getSelection()  # 選択しているオブジェクトを取得。
			stringaddress = getStringAddressFromCellRange(selection)
			if stringaddress:
				filename = "_".join((name, inspect.currentframe().f_code.co_name, stringaddress.replace(":", "")))
				txt = "Selection: {}\nSource: {}".format(stringaddress, source)
		if not txt:
			txt = "Source: {}".format(source)
		createLog(dirpath, filename, txt)	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeSelectionChangeListener(self)
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def mousePressed(self, enhancedmouseevent):
		self._createLog(enhancedmouseevent, inspect.currentframe().f_code.co_name)
		return True
	def mouseReleased(self, enhancedmouseevent):
		self._createLog(enhancedmouseevent, inspect.currentframe().f_code.co_name)
		return True
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeEnhancedMouseClickHandler(self)
	def _createLog(self, enhancedmouseevent, methodname):
		dirpath, name = self.args
		target = enhancedmouseevent.Target
		target = getStringAddressFromCellRange(target) or target  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。
		clickcount = enhancedmouseevent.ClickCount
		filename = "_".join((name, methodname, "ClickCount", str(clickcount)))
		createLog(dirpath, filename, "Buttons: {}, ClickCount: {}, PopupTrigger {}, Modifiers: {}, Target: {}\nSource: {}".format(enhancedmouseevent.Buttons, clickcount, enhancedmouseevent.PopupTrigger, enhancedmouseevent.Modifiers, target, enhancedmouseevent.Source))	
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def activeSpreadsheetChanged(self, activationevent):
		dirpath, name = self.args
		activesheet = activationevent.ActiveSheet
		activesheetname = activesheet.getName()
		txt = ""
		source = activationevent.Source
		if source.supportsService("com.sun.star.sheet.SpreadsheetView"):  # sourceがコントローラーのとき
			selection = source.getSelection()  # 選択しているオブジェクトを取得。
			stringaddress = getStringAddressFromCellRange(selection)
			if stringaddress:
				txt = "Selection: {}\nSource: {}".format(stringaddress, source)
		if not txt:
			txt = "Source: {}".format(source)
		txt = "ActiveSheet: {}, {}".format(activesheetname, txt)  # アクティブシート名を取得。
		filename = "_".join((name, inspect.currentframe().f_code.co_name, activesheetname))
		createLog(dirpath, filename, txt)	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeActivationEventListener(self)
class TitleChangeListener(unohelper.Base, XTitleChangeListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
		self.args = dirpath, name	
	def titleChanged(self, titlechangedevent):
		dirpath, name = self.args
		title = titlechangedevent.Title
		filename = "_".join((name, inspect.currentframe().f_code.co_name, title))
		createLog(dirpath, filename, "Title: {}, Source: {}".format(title, titlechangedevent.Source))	
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
		self.subj.removeTitleChangeListener(self)
class CloseListener(unohelper.Base, XCloseListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def queryClosing(self, eventobject, getsownership):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "getsownership: {}, Source: {}".format(getsownership, eventobject.Source))	
	def notifyClosing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
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
				filename = "_".join((name, inspect.currentframe().f_code.co_name, frameactionname))  # Action名も追加。
				createLog(dirpath, filename, "FrameAction: {}, Source: {}".format(frameactionname, frameactionevent.Source))
				return
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
class TerminateListener(unohelper.Base, XTerminateListener):  # TerminateListener
	def __init__(self, dirpath, name):  # 出力先ディレクトリのパス、リスナーのインスタンス名。
		self.args = dirpath, name
	def queryTermination(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))  # Sourceを出力。
	def notifyTermination(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))  # Sourceを出力。
	def disposing(self, eventobject):
		dirpath, name = self.args
		filename = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, filename, "Source: {}".format(eventobject.Source))	
def getStringAddressFromCellRange(source):  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。文字列アドレスが取得できないオブジェクトの時はオブジェクトの文字列を返す。	
	stringaddress = ""
	propertysetinfo = source.getPropertySetInfo()  # PropertySetInfo
	if propertysetinfo.hasPropertyByName("AbsoluteName"):  # AbsoluteNameプロパティがある時。
		absolutename = source.getPropertyValue("AbsoluteName") # セル範囲コレクションは$Sheet1.$A$4:$A$6,$Sheet1.$B$4という形式で返る。
		names = absolutename.replace("$", "").split(",")  # $を削除してセル範囲の文字列アドレスのリストにする。
		stringaddress = ", ".join(names)  # コンマでつなげる。
	return stringaddress
C = 100  # カウンターの初期値。
TIMESTAMP = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
def createLog(dirpath, filename, txt):  # 年月日T時分秒リスナーのインスタンス名_メソッド名(_オプション).logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	global C
	filename = "".join((TIMESTAMP, "_", str(C), filename, ".log"))
	C += 1
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。	
