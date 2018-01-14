#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import inspect
from datetime import datetime
from com.sun.star.util import XCloseListener
from com.sun.star.frame import XTitleChangeListener
from com.sun.star.frame import XTerminateListener
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
def macro(documentevent=None):  # 引数は文書のイベント駆動用。OnStartAppでもDocumentEventが入る。LibreOfficeに保存する。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	thisscriptpath = unohelper.fileUrlToSystemPath(__file__)  # __file__はfileurlで返ってくるのでシステムパスに変換。
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
	desktop_terminatelistener = TerminateListener(dirpath, "desktop_terminatelistener", desktop)  # TerminateListener
	desktop.addTerminateListener(desktop_terminatelistener)  # 使い終わったらremoveしないといけない。
	desktop_frameactionlistener = FrameActionListener(dirpath, "desktop_frameactionlistener", desktop)  # FrameActionListener
	desktop.addFrameActionListener(desktop_frameactionlistener)  # いつ呼ばれる?
	controller = doc.getCurrentController()  # コントローラーの取得。
	frame = controller.getFrame()  # フレームの取得。
	frame.addFrameActionListener(FrameActionListener(dirpath, "frame_frameactionlistener", frame))  # FrameActionListener
	frame.addCloseListener(CloseListener(dirpath, "frame_closelistener", frame))  # CloseListener
	frame.addTitleChangeListener(TitleChangeListener(dirpath, "frame_titlechangelistener", frame))  # TitleChangeListener
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
	def __init__(self, dirpath, name, subj):
		self.subj = subj
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
		self.subj.removeCloseListener(self)
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def __init__(self, dirpath, name, subj):
		self.subj = subj
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
		self.subj.removeFrameActionListener(self)
class TerminateListener(unohelper.Base, XTerminateListener):  # TerminateListener
	def __init__(self, dirpath, name, subj):  # 出力先ディレクトリのパス、リスナーのインスタンス名。
		self.subj = subj
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
		self.subj.removeFrameTerminateListener(self)
C = 100  # カウンターの初期値。
TIMESTAMP = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
def createLog(dirpath, filename, txt):  # 年月日T時分秒リスナーのインスタンス名_メソッド名(_オプション).logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	global C
	filename = "".join((TIMESTAMP, "_", str(C), filename, ".log"))
	C += 1
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。
