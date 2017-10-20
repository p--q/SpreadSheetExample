#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
	def wrapper(*args, **kwargs):
		frame = None
		doc = XSCRIPTCONTEXT.getDocument()
		if doc:  # ドキュメントが取得できた時
			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
		else:
			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
			frame = currentframe.getCreator()
		if frame:   
			import time
			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
			t = 1  # プレグレスバーの初期値。
			while t<=maxrange:  # プログレスバーの最大値以下の間。
				indicator.setValue(t)  # プレグレスバーの位置を設定。
				time.sleep(1)  # 1秒待つ。
				t += 1  # プログレスバーの目盛りを増やす。
			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
# @enableRemoteDebugging
def macro(documentevent):
	doc = documentevent.Source  # ドキュメントの取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler())  # マウスハンドラをコントローラに設定。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
# 	@enableRemoteDebugging  # ダブルクリックで有効にするとLibreOfficeがクラッシュする。
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		target = enhancedmouseevent.Target  # ターゲットを取得。
# 		doc = XSCRIPTCONTEXT.getDocument()
# 		sheets = doc.getSheets()  # シートコレクション。
# 		sheet = sheets[0]  # 最初のシート。
# 		sheet[0, 0].setString(str(target))
		
		
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
				c = enhancedmouseevent.ClickCount  # クリック数を取得。
				target.setString("ClickCount: {}".format(c))  # セルにクリック数を出力する。
				if c==2:  # ダブルクリックのとき。クリック数が2になったときに実行されるので実質wクリック以上。
					row = target.getCellAddress().Row  # 行インデックスを取得。
					if row>9:  # 行インデックスが9(行10)以上のとき。
						target.setPropertyValue("CellBackColor", 0x8080FF)  # 背景を青紫色にする。
						return False  # セル編集モードにしない。
					else:  # returnしていないので3クリック以上できる。
						target.setPropertyValue("CellBackColor", 0xFFFF80)  # 背景を黄色にする。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。
# 	@enableRemoteDebugging
	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
		return True  # Trueでイベントを次のハンドラに渡す。
	def disposing(self, eventobject):
		pass
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
	
	