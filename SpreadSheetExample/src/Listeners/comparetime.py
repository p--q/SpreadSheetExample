#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import time, math
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.sheet import CellFlags as cf # 定数
def macro(documentevent=None):  # 引数は文書のイベント駆動用。OnStartAppでもDocumentEventが入るがSourceはNoneになる。# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラーの取得。
	sheet = controller.getActiveSheet()
	sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。
	sheet["A1:D1"].setDataArray((("EachCell", "addActionLock", "lockControllers", "setDataArray"),))
	sheet["A1:D1"].getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(doc, sheet))  # EnhancedMouseClickHandler
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, doc, sheet):
		self.args = doc, sheet
	def mousePressed(self, enhancedmouseevent):
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					txt = target.getString()
					doc, sheet = self.args
					sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。
					sheet["A1:D1"].setDataArray((("EachCell", "addActionLock", "lockControllers", "setDataArray"),))
					x = -3.14
					y = 3.14
					n = 30000
					d = (y-x)/n	
					if txt=="EachCell":
						start = time.perf_counter()
						for i in range(1, n+1):
							k = x + d*i
							sheet[i, 0].setValue(k)
							sheet[i, 1].setValue(math.sin(k))
							sheet[1, 2].setValue(i)
						end = time.perf_counter()
					elif txt=="addActionLock":
						start = time.perf_counter()
						doc.addActionLock()
						for i in range(1, n+1):
							k = x + d*i
							sheet[i, 0].setValue(k)
							sheet[i, 1].setValue(math.sin(k))
							sheet[1, 2].setValue(i)
						doc.removeActionLock()	
						end = time.perf_counter()
					elif txt=="lockControllers":
						start = time.perf_counter()
						doc.lockControllers()
						for i in range(1, n+1):
							k = x + d*i
							sheet[i, 0].setValue(k)
							sheet[i, 1].setValue(math.sin(k))
							sheet[1, 2].setValue(i)
						doc.unlockControllers()
						end = time.perf_counter()
					elif txt=="setDataArray":
						start = time.perf_counter()
						rows = [(x+d*i, math.sin(x+d*i)) for i in range(1, n+1)]
						sheet[1:len(rows)+1, 0:len(rows[0])].setDataArray(rows)
						sheet[1, 2].setValue(len(rows))
						end = time.perf_counter()
					sheet[2, 2].setString(txt)
					sheet[3, 2].setString("Finished: {}s".format(end-start))
					return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
	def mouseReleased(self, enhancedmouseevent):
		return True
	def disposing(self, eventobject):
		pass
g_exportedScripts = macro,  #マクロセレクターに限定表示させる関数をタプルで指定。	
