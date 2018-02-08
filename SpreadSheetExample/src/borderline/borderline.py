#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.table import BorderLine2  # Struct
from com.sun.star.table import BorderLineStyle  # 定数
from com.sun.star.table import TableBorder2  # Struct
def macro(documentevent=None):  # 引数は文書のイベント駆動用。   
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	colors = {"clearblue": 0x9999FF, "magenta": 0xFF00FF}  # 色の設定。
	# 枠線の作成。
	noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)  # 枠線を消すための空線。
	firstline = BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=colors["clearblue"])  # 青色の枠線。
	secondline =  BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=colors["magenta"])	# 桃色の枠線。
	tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)  # 上下左右の枠線。
	topbottomtableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=False, IsRightLineValid=False)  # 上下の枠線。
	leftrighttableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=False, IsBottomLineValid=False, IsLeftLineValid=True, IsRightLineValid=True)  # 左右の枠線。
	borders = noneline, tableborder2, topbottomtableborder, leftrighttableborder  # 作成した枠線をまとめたタプル。
	controller.addSelectionChangeListener(SelectionChangeListener(colors, borders))
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(controller, borders))  # EnhancedMouseClickHandler。このリスナーのメソッドの引数からコントローラーを取得する方法がない。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, controller, borders):
		self.controller = controller
		self.args = borders
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。固定行列の最初のクリックは同じ相対位置の固定していないセルが返ってくる(表示されている自由行の先頭行に背景色がる時のみ）。
		borders = self.args
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
					sheet = target.getSpreadsheet()
					drowBorders(self.controller, sheet, target, borders)
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
	def mouseReleased(self, enhancedmouseevent):
		pass
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # eventobject.SourceはNone。
		self.controller.removeEnhancedMouseClickHandler(self)	
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, colors, borders):
		self.args = colors, borders
	def selectionChanged(self, eventobject):  # マウスから呼び出した時の反応が遅い。このメソッドでエラーがでるとショートカットキーでの操作が必要。
		colors, borders = self.args	
		controller = eventobject.Source
		sheet = controller.getActiveSheet()
		selection = controller.getSelection()
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
			currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
			if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==colors["clearblue"],\
					currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==colors["magenta"])):  # 枠線の色を確認。
				return  # すでに枠線が書いてあったら何もしない。
		if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
			drowBorders(controller, sheet, selection, borders)	
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionChangeListener(self)	
def drowBorders(controller, sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders	
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	rangeaddress = cellrange.getRangeAddress()  # セル範囲アドレスを取得。
	sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
	sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
	cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
