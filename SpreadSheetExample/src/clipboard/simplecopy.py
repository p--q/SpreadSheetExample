#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException  # 例外
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard
	systemclipboard.setContents(TextTransferable(selection[0, 0].getString()), None)  # クリップボードにコピーする。
class TextTransferable(unohelper.Base, XTransferable):
	def __init__(self, txt):  # クリップボードに渡す文字列を受け取る。
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
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
