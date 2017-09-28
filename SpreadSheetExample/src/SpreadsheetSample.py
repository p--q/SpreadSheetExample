#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.table import BorderLine
from com.sun.star.table import TableBorder
from com.sun.star.awt import FontWeight
def macro():
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  

	doCellSamples(doc)
	doCellRangeSamples()
	doCellRangesSamples()
	doCellCursorSamples()
	doFormattingSamples()
	doDocumentSamples()
	doDatabaseSamples()
	doDataPilotSamples()
	doNamedRangesSamples()
	doFunctionAccessSamples()
	doApplicationSettingsSamples()

def	doCellSamples(doc):
	sheets = doc.getSheets()
	sheet = sheets[0]	
	prepareRange(sheet, "A1:C7", "Cells and Cell Ranges")
	
	

# ** Draws a colored border around the range and writes the headline in the first cell.
def prepareRange(sheet, rng, headline):
	# draw border
	cellrange = sheet[rng]
	borderline = BorderLine(Color=0x99CCFF, InnerLineWidth=0, LineDistance=0, OuterLineWidth=100)
	tableborder = TableBorder(TopLine=borderline, BottomLine=borderline, LeftLine=borderline, RightLine=borderline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
	cellrange.setPropertyValue("TableBorder", tableborder)
	# draw headline
	cellrange.setPropertyValue("CellBackColor", 0x99CCFF)
	# write headline
	cell = cellrange[0, 0]
	cell.setFormula(headline)
	cell.setPropertyValues(("CharColor", "CharWeight"), (0x003399, FontWeight.BOLD))
	
	
	
	
	
def	doCellRangeSamples():
	pass
def	doCellRangesSamples():
	pass
def	doCellCursorSamples():
	pass
def	doFormattingSamples():
	pass
def	doDocumentSamples():
	pass
def	doDatabaseSamples():
	pass
def	doDataPilotSamples():
	pass
def	doNamedRangesSamples():
	pass
def	doFunctionAccessSamples():
	pass
def	doApplicationSettingsSamples():
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
	doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
# 	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
	if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
	flg = True
	while flg:
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		if doc is not None:
			flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
	macro()