#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import sys
import time
from com.sun.star.view import DocumentZoomType
from com.sun.star.datatransfer.clipboard import XClipboardListener
from com.sun.star.datatransfer.clipboard import XClipboardOwner
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doc.getText().setString("""In the first step, paste the current content of the clipboard in the document!
The text \"Hello world!\" shall be insert at the current cursor position below.

In the second step, please select some words and put it into the clipboard! ...
	
Current clipboard content = """);
	controller = doc.getCurrentController()
	controller.getViewSettings().setPropertyValue("ZoomType", DocumentZoomType.OPTIMAL)
	systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)
	clipboardlistener = ClipboardListener()
	systemclipboard.addClipboardListener(clipboardlistener)
	readClipBoard(systemclipboard)
	print("Becoming a clipboard owner...\n")	
	clipboardowner = ClipboardOwner()
	systemclipboard.setContents(TextTransferable("Hello World!"), clipboardowner)
	first = 0
	while clipboardowner.isowner:
		if first!=2:
			if first==1:
				print("""Change clipboard ownership by putting something into the clipboard!

Still clipboard owner...""")
			else:
				print("Still clipboard owner...")
			first += 1
		else:
			print(".", end="")  # 改行が来るまで出力されない。
		time.sleep(1)  # 1秒待つ。
	readClipBoard(systemclipboard)
	systemclipboard.removeClipboardListener(clipboardlistener)
	if hasattr(doc, "close"):
		doc.close(True)
	else:
		doc.dispose() 
	sys.exit(0)
def readClipBoard(systemclipboard):
	transferable = systemclipboard.getContents()
	dataflavors = transferable.getTransferDataFlavors()
	print("""Reading the clipboard...
Available clipboard formats:""")
	flavor = None
	for dataflavor in dataflavors:
		print("""MimeType: {}
HumanPresentableName: {}""".format(dataflavor.MimeType, dataflavor.HumanPresentableName))
		if dataflavor.MimeType=="text/plain;charset=utf-16":
			flavor = dataflavor
	if flavor is None:
		print("""
Requested format is not available on the clipboard!""")
	else:
		print("""
Unicode text on the clipboard ...
Your selected text \"{}\" is now in the clipboard.
""".format(transferable.getTransferData(flavor)))
class TextTransferable(unohelper.Base, XTransferable):
	def __init__(self, txt):
		self.txt = txt
		self.unicode_content_type = "text/plain;charset=utf-16"
	def getTransferData(self, flavor):
		if flavor.MimeType.lower()!=self.unicode_content_type:
			raise UnsupportedFlavorException()
		return self.txt
	def getTransferDataFlavors(self):
		return DataFlavor(MimeType=self.unicode_content_type, HumanPresentableName="Unicode Text"),  # DataTypeの設定方法は不明。
	def isDataFlavorSupported(self, flavor):
		return flavor.MimeType.lower()==self.unicode_content_type
class ClipboardOwner(unohelper.Base, XClipboardOwner):
	def __init__(self):
		self.isowner = True
	def lostOwnership(self, clipboard, transferable):
		print("""
Lost clipboard ownership...
""")
		self.isowner = False
class ClipboardListener(unohelper.Base, XClipboardListener):
	def changedContents(self, clipboardEvent):
		print("\nClipboard content has changed!\n")
	def disposing(self, eventobject):
		pass
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
	import sys
	from com.sun.star.beans import PropertyValue
	from com.sun.star.script.provider import XScriptContext  
	def connectOffice(func):  # funcの前後でOffice接続の処理
		@wraps(func)
		def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
			try:
				ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
			except:
				print("Could not establish a connection with a running office.", file=sys.stderr)
				sys.exit()
			print("Connected to a running office ...")
			smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
			print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
			return func(ctx, smgr)  # 引数の関数の実行。
		def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
			cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
			node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
			ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
			return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
		return wrapper
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
		class ScriptContext(unohelper.Base, XScriptContext):
			def __init__(self, ctx):
				self.ctx = ctx
			def getComponentContext(self):
				return self.ctx
			def getDesktop(self):
				return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
			def getDocument(self):
				return self.getDesktop().getCurrentComponent()
		return ScriptContext(ctx)  
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
# 	doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
	if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
	flg = True
	while flg:
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		if doc is not None:
			flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
	macro()