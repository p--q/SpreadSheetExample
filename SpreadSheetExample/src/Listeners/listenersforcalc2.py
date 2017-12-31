#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import inspect
from datetime import datetime
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import XKeyHandler
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
	path = doc.getURL() if __file__.startswith("vnd.sun.star.tdoc:") else __file__  # このスクリプトのパス。fileurlで返ってくる。埋め込みマクロの時は埋め込んだドキュメントのURLで代用する。
	thisscriptpath = unohelper.fileUrlToSystemPath(path)  # fileurlをsystempathに変換。
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	desktop = XSCRIPTCONTEXT.getDesktop()  # デスクトップの取得。
	desktop.addTerminateListener(TerminateListener(dirpath, "desktop_terminatelistener"))  # TerminateListenerは使い終わったらremoveしないと問題が起こる。
	desktop.addFrameActionListener(FrameActionListener(dirpath, "desktop_frameactionlistener"))  # FrameActionListener
	controller = doc.getCurrentController()  # コントローラーの取得。
	frame = controller.getFrame()  # フレームの取得。
	frame.addFrameActionListener(FrameActionListener(dirpath, "frame_frameactionlistener"))  # FrameActionListener 
	frame.addCloseListener(CloseListener(dirpath, "frame_closelistener"))  # CloseListener
	frame.addTitleChangeListener(TitleChangeListener(dirpath, "frame_titlechangelistener"))  # TitleChangeListener
	controller.addActivationEventListener(ActivationEventListener(dirpath, "controller_activationeventlistener"))  # ActivationEventListener
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(dirpath, "controller_enhancedmouseclickhandler"))  # EnhancedMouseClickHandler
	controller.addSelectionChangeListener(SelectionChangeListener(dirpath, "controller_selectionchangelistener"))  # SelectionChangeListener
	controller.addBorderResizeListener(BorderResizeListener(dirpath, "controller_borderresizelistener"))  # BorderResizeListener
	controller.addTitleChangeListener(TitleChangeListener(dirpath, "controller_titlechangelistener"))  # TitleChangeListener	
	controller.addKeyHandler(KeyHandler(dirpath, "controller_keyhandler"))  # KeyHandler		
	doc.addDocumentEventListener(DocumentEventListener(dirpath, "doc_documenteventlistener"))  # DocumentEventListener	
	doc.addEventListener(EventListener(dirpath, "doc_eventlistener"))  # EventListener	
	doc.addModifyListener(ModifyListener(dirpath, "doc_modifylistener"))  # ModifyListener		
	doc.addPrintJobListener(PrintJobListener(dirpath, "doc_printjoblistener"))  # PrintJobListener	
	doc.addStorageChangeListener(StorageChangeListener(dirpath, "doc_storagechangelistener"))  # StorageChangeListener	
	doc.addTitleChangeListener(TitleChangeListener(dirpath, "doc_titlechangelistener"))  # TitleChangeListener
	doc.addChangesListener(ChangesListener(dirpath, "doc_changelistener"))  # ChangesListener	
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	sheet.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "sheet_chartdatachangeeventlistener"))  # ChartDataChangeEventListener	
	sheet.addModifyListener(ModifyListener(dirpath, "sheet_modifylistener"))  # ModifyListener	
	cell = sheet["A1"]  # セルの取得。
	cell.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "cell_chartdatachangeeventlistener"))  # ChartDataChangeEventListener		
	cell.addModifyListener(ModifyListener(dirpath, "cell_modifylistener"))  # ModifyListener
	cells = sheet["A2:C4"]  # セル範囲の取得。
	cells.addChartDataChangeEventListener(ChartDataChangeEventListener(dirpath, "cells_chartdatachangeeventlistener"))  # ChartDataChangeEventListener	
	cells.addModifyListener(ModifyListener(dirpath, "cells_modifylistener"))  # ModifyListener	
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
				methodname = "_".join((name, inspect.currentframe().f_code.co_name, chartdatachangetypename))  # ChartDataChangeType名を追加。
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
				methodname = "_".join((name, inspect.currentframe().f_code.co_name, printablestatename))  # State名も追加。
				createLog(dirpath, methodname, "PrintableState: {}, Source: {}".format(printablestatename, printjobevent.Source))
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
		methodname = "_".join((name, inspect.currentframe().f_code.co_name, eventobject.EventName))  # イベント名も追加。
		createLog(dirpath, methodname, "EventName: {}, Source: {}".format(eventobject.EventName, eventobject.Source))	
	def disposing(self, eventobject):
		pass		
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def documentEventOccured(self, documentevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name, documentevent.EventName))  # イベント名も追加。
		createLog(dirpath, methodname, "EventName: {}, Source: {}".format(documentevent.EventName, documentevent.Source))	
	def disposing(self, eventobject):
		pass	
class KeyHandler(unohelper.Base, XKeyHandler):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def keyPressed(self, keyevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "KeyCode: {}, KeyFunc: {}, Modifiers: {}".format(keyevent.KeyCode, keyevent.KeyFunc, keyevent.Modifiers))	 # KeyCharが矢印などの場合はその後のテキストが表示されないので書き込まない。
		return False
	def keyReleased(self, keyevent):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, "KeyCode: {}, KeyFunc: {}, Modifiers: {}".format(keyevent.KeyCode, keyevent.KeyFunc, keyevent.Modifiers))	 # KeyCharが矢印などの場合はその後のテキストが表示されないので書き込まない。	
		return False		
	def disposing(self, eventobject):
		pass			
class BorderResizeListener(unohelper.Base, XBorderResizeListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def borderWidthsChanged(self, obj, borderwidths):
		dirpath, name = self.args
		if obj.supportsService("com.sun.star.sheet.SpreadsheetView"):  # objがコントローラーの時。
			cellrangeaddressconversion = obj.getModel().createInstance("com.sun.star.table.CellRangeAddressConversion")  # ドキュメントからCellRangeAddressConversionを取得。
			cellrangeaddressconversion.Address = obj.getVisibleRange()  # 表示されているセル範囲のCellRangeAddressを取得。
			txt = "Visible Range: {}".format(cellrangeaddressconversion.PersistentRepresentation)  # 表示されているセル範囲の文字列アドレスの取得。
		else:
			txt = "Top: {}, Left: {}, Right: {}, Bottom: {}, Object: {}".format(borderwidths.Top, borderwidths.Left, borderwidths.Right, borderwidths.Bottom, obj)
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, txt)	
	def disposing(self, eventobject):
		pass		
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def selectionChanged(self, eventobject):
		dirpath, name = self.args
		txt = getStringAddressFromController(eventobject.Source)  # sourceがコントローラーの時は選択範囲の文字列アドレスを返す。
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, txt)	
	def disposing(self, eventobject):
		pass	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def mousePressed(self, enhancedmouseevent):
		dummy, name = self.args
		self._createLog(enhancedmouseevent, "_".join((name, inspect.currentframe().f_code.co_name)))
		return True
	def mouseReleased(self, enhancedmouseevent):
		dummy, name = self.args
		self._createLog(enhancedmouseevent, "_".join((name, inspect.currentframe().f_code.co_name)))
		return True
	def disposing(self, eventobject):
		pass
	def _createLog(self, enhancedmouseevent, methodname):
		dirpath, dummy = self.args
		target = enhancedmouseevent.Target
		target = getStringAddressFromCellRange(target) or target  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。
		createLog(dirpath, methodname, "Buttons: {}, ClickCount: {}, PopupTrigger {}, Modifiers: {}, Target: {}".format(enhancedmouseevent.Buttons, enhancedmouseevent.ClickCount, enhancedmouseevent.PopupTrigger, enhancedmouseevent.Modifiers, target))	
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, dirpath, name):
		self.args = dirpath, name	
	def activeSpreadsheetChanged(self, activationevent):
		dirpath, name = self.args
		activesheet = activationevent.ActiveSheet
		txt = "ActiveSheet: {} ".format(activesheet.getName())  # アクティブシート名を取得。
		txt += getStringAddressFromController(activationevent.Source)  # sourceがコントローラーの時は選択範囲の文字列アドレスを返す。
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))
		createLog(dirpath, methodname, txt)	
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
				methodname = "_".join((name, inspect.currentframe().f_code.co_name, frameactionname))  # Action名も追加。
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
def getStringAddressFromController(source):	 # sourceがコントローラーの時は選択範囲の文字列アドレスを返す。文字列アドレスが取得できないオブジェクトの時はオブジェクトの文字列を返す。
	if source.supportsService("com.sun.star.sheet.SpreadsheetView"):  # sourceがコントローラーのとき
		selection = source.getSelection()  # 選択しているオブジェクトを取得。
		selection = getStringAddressFromCellRange(selection) or selection  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。
		return "Selection: {}".format(selection)
	else:  # Sourceがコントローラーでない時
		return "Source: {}".format(source)
def getStringAddressFromCellRange(source):  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。文字列アドレスが取得できないオブジェクトの時はオブジェクトの文字列を返す。	
	stringaddress = ""
	propertysetinfo = source.getPropertySetInfo()  # PropertySetInfo
	if propertysetinfo.hasPropertyByName("AbsoluteName"):  # AbsoluteNameプロパティがある時。
		absolutename = source.getPropertyValue("AbsoluteName") # セル範囲コレクションは$Sheet1.$A$4:$A$6,$Sheet1.$B$4という形式で返る。
		names = absolutename.replace("$", "").split(",")  # $を削除してセル範囲の文字列アドレスのリストにする。
		stringaddress = ", ".join(names)  # コンマでつなげる。
	return stringaddress
C = 10  # カウンターの初期値。
TIMESTAMP = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
def createLog(dirpath, methodname, txt):  # 年月日T時分秒リスナーのインスタンス名_methodname.logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	global C
	filename = "".join((TIMESTAMP, "_", str(C), methodname, ".log"))
	C += 1
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。	
