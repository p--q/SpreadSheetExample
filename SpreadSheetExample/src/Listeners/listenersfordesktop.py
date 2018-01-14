#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import inspect
from datetime import datetime
from com.sun.star.frame import XTerminateListener
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
def macro(documentevent=None):  # 引数は文書のイベント駆動用。OnStartAppでもDocumentEventが入る。LibreOfficeに保存する。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	thisscriptpath = unohelper.fileUrlToSystemPath(__file__)  # __file__はfileurlで返ってくるのでシステムパスに変換。
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
	desktop_terminatelistener = TerminateListener(dirpath, "desktop_terminatelistener")  # TerminateListener
	desktop.addTerminateListener(desktop_terminatelistener)  # 使い終わったらremoveしないといけない。
	desktop_frameactionlistener = FrameActionListener(dirpath, "desktop_frameactionlistener")  # FrameActionListener
	desktop.addFrameActionListener(desktop_frameactionlistener)  # いつ呼ばれる?
class FrameActionListener(unohelper.Base, XFrameActionListener):  # FrameActionListener
	def __init__(self, dirpath, name):  # 出力先ディレクトリのパス、リスナーのインスタンス名。
		enums = COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
		frameactionnames = "COMPONENT_ATTACHED", "COMPONENT_DETACHING", "COMPONENT_REATTACHED", "FRAME_ACTIVATED", "FRAME_DEACTIVATING", "CONTEXT_CHANGED", "FRAME_UI_ACTIVATED", "FRAME_UI_DEACTIVATING"
		self.args = dirpath, name, zip(enums, frameactionnames)
	def frameAction(self, frameactionevent):
		dirpath, name, frameactions = self.args
		frameaction = frameactionevent.Action
		for enum, frameactionname in frameactions:
			if frameaction==enum:
				methodname = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
				createLog(dirpath, methodname, "frameactionevent.FrameAction(enum): {}".format(frameactionname))  # frameactionのenum名を出力。
				return
	def disposing(self, eventobject):
		pass
class TerminateListener(unohelper.Base, XTerminateListener):  # TerminateListener
	def __init__(self, dirpath, name):  # 出力先ディレクトリのパス、リスナーのインスタンス名。
		self.args = dirpath, name
	def queryTermination(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
		createLog(dirpath, methodname, "eventobject.Source: {}".format(eventobject.Source))  # Sourceを出力。
	def notifyTermination(self, eventobject):
		dirpath, name = self.args
		methodname = "_".join((name, inspect.currentframe().f_code.co_name))  # このメソッド名を取得。メソッド内で実行する必要がある。
		createLog(dirpath, methodname, "eventobject.Source: {}".format(eventobject.Source))  # Sourceを出力。
		desktop = eventobject.Source
		desktop.removeTerminateListener(self)  # TerminateListenerを除去。除去しなmethodname = inspect.currentframe().f_code.co_nameいとLibreOfficeのプロセスが残って起動できなくなる。
	def disposing(self, eventobject):
		pass
def createLog(dirpath, methodname, txt):  # 年月日T時分秒リスナーのインスタンス名_methodname.logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	timestamp = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
	filename = "".join((timestamp, methodname, ".log"))
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)	
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。	
