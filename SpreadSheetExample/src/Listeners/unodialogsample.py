#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import re
from collections import deque, namedtuple
from com.sun.star.util import Time, Date, XCloseListener
from com.sun.star.lang import Locale
from com.sun.star.awt.ScrollBarOrientation import VERTICAL
from com.sun.star.awt import XMouseListener, XTextListener, XFocusListener, XKeyListener, XSpinListener, XItemListener, XAdjustmentListener
from com.sun.star.awt.SystemPointer import REFHAND
from com.sun.star.awt.MessageBoxType import INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.FocusChangeReason import TAB
from com.sun.star.awt.AdjustmentType import ADJUST_LINE, ADJUST_PAGE, ADJUST_ABS 
from com.sun.star.awt.Key import BACKSPACE, SPACE, DELETE, LEFT, RIGHT, HOME, END
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
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。  
	dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": 102, "PositionY": 41, "Width": 380, "Height": 380, "Title": "LibreOffice", "Name": "MyTestDialog", "Step": 0, "Moveable": True})  # "TabIndex": 0
	textlistener = TextListener()
	spinlistener = SpinListener()
	itemlistener = ItemListener(dialog) 
	addControl("FixedText", {"Name": "Headerlabel", "PositionX": 106, "PositionY": 6, "Width": 300, "Height": 8, "Label": "This code-sample demonstrates how to create various controls in a dialog"})
	addControl("FixedText", {"PositionX": 106, "PositionY": 18, "Width": 100, "Height": 8, "Label": "My Label", "NoLabel": True}, {"addMouseListener": MouseListener(ctx, smgr)})  # , "Step": 0
	addControl("CurrencyField", {"PositionX": 106, "PositionY": 30, "Width": 60, "Height": 12, "PrependCurrencySymbol": True, "CurrencySymbol": "$", "Value": 2.93}, {"addTextListener": textlistener})
	addControl("FixedLine", {"PositionX": 106, "PositionY": 58, "Width": 100, "Height": 8, "Orientation": 0, "Label": "My FixedLine"})   
	addControl("Edit", {"PositionX": 106, "PositionY": 72, "Width": 60, "Height": 12, "Text": "MyText", "EchoChar": ord("*"), "HelpText": "EchoChar will be canceled when moving the focus with the tab key."}, {"addFocusListener": FocusListener(), "addKeyListener": KeyListener(dialog)})  
	addControl("FixedLine", {"PositionX": 106, "PositionY": 58, "Width": 100, "Height": 8, "Orientation": 0, "Label": "My FixedLine"}) 
	t, tmin, tmax = toTime(10, 0, 0), toTime(1, 0, 0), toTime(17, 5, 0)
	addControl("TimeField", {"PositionX": 106, "PositionY": 96, "Width": 50, "Height": 12, "Spin": True, "TimeFormat": 5, "Time": t.Time, "TimeMin": tmin.Time, "TimeMax": tmax.Time, "HelpText": "Min: {} Max: {}".format(tmin, tmax)})  # com.sun.star.util.Timeで時刻を指定。  
	d, dmin, dmax = toDate(2017, 7, 4), toDate(2017, 6, 16), toDate(2017, 8, 15)
	addControl("DateField", {"PositionX": 166, "PositionY": 96, "Width": 55, "Height": 12, "Dropdown": True, "DateFormat": 9, "DateMin": dmin.Date, "DateMax": dmax.Date, "Date": d.Date, "Spin": True, "HelpText": "Min: {} Max: {}".format(dmin, dmax)}, {"addSpinListener": spinlistener})	 # com.sun.star.util.Dateで日付を指定。
	addControl("GroupBox", {"PositionX": 102, "PositionY": 124, "Width": 100, "Height": 70, "Label": "My GroupBox"})   
	addControl("PatternField", {"PositionX": 106, "PositionY": 136, "Width": 50, "Height": 12, "LiteralMask": "__.05.2007", "EditMask": "NNLLLLLLLL", "StrictFormat": True, "HelpText": "_ means a digit can be entered"})   
	addControl("NumericField", {"PositionX": 106, "PositionY": 152, "Width": 50, "Height": 12, "Spin": True, "StrictFormat": True, "ValueMin": 0.0, "ValueMax": 1000.0, "Value": 500.0, "ValueStep": 100.0, "ShowThousandsSeparator": True, "DecimalAccuracy": 1})  
	addControl("CheckBox", {"PositionX": 106, "PositionY": 168, "Width": 150, "Height": 8, "Label": "~Enable Close dialog Button", "TriState": True, "State": 1}, {"addItemListener": itemlistener})  
	addControl("RadioButton", {"PositionX": 130, "PositionY": 200, "Width": 150, "Height": 8, "Label": "~First Option", "State": 1, "TabIndex": 50})	 
	addControl("RadioButton", {"PositionX": 130, "PositionY": 214, "Width": 150, "Height": 8, "Label": "~Second Option", "TabIndex": 51})	  
	addControl("ListBox", {"PositionX": 106, "PositionY": 230, "Width": 50, "Height": 30, "Dropdown": False, "Step": 0, "MultiSelection": True, "StringItemList": ("First Item", "Second Item", "ThreeItem"), "SelectedItems": (0, 2)})	 
	addControl("ComboBox", {"PositionX": 160, "PositionY": 230, "Width": 60, "Height": 12, "Dropdown": True, "MaxTextLen": 10, "ReadOnly": False, "Autocomplete": True, "StringItemList": ("First Entry", "Second Entry", "Third Entry", "Fourth Entry")}, {"addItemListener": itemlistener})  # 選択した文字列が取得できない。
	numberformatssupplier = smgr.createInstanceWithContext("com.sun.star.util.NumberFormatsSupplier", ctx)  # フォーマットサプライヤーをインスタンス化。
	numberformats = numberformatssupplier.getNumberFormats()  # フォーマットサプライヤーからフォーマット一覧を取得。
	formatstring = "NNNNMMMM DD, YYYY"  # フォーマット。デフォルトのフォーマット一覧はCalc→書式→セル→数値でみれる。
	locale = Locale(Language="en", Country="US")  # フォーマット一覧をくくる言語と国を設定。
	formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べる。第3引数のブーリアンは意味はないはず。
	if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
		formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと?
	addControl("FormattedField", {"PositionX": 106, "PositionY": 270, "Width": 100, "Height": 12, "EffectiveValue": 12348, "StrictFormat": True, "Spin": True, "FormatsSupplier": numberformatssupplier, "FormatKey": formatkey}, {"addSpinListener": spinlistener})  
	addControl("ScrollBar", {"PositionX": 230, "PositionY": 230, "Width": 8, "Height": 52, "Orientation": VERTICAL, "ScrollValueMin": 0, "ScrollValueMax": 100, "ScrollValue": 5, "LineIncrement": 2, "BlockIncrement": 10}, {"addAdjustmentListener": AdjustmentListener(ctx, smgr, docframe)})  
	workurl = ctx.getByName('/singletons/com.sun.star.util.thePathSettings').getPropertyValue("Work")  # Ubuntuではホームフォルダ、Windows10ではドキュメントフォルダのURIが返る。
	systemworkpath = unohelper.fileUrlToSystemPath(workurl)  # URIをシステム固有のパスに変換する。
	addControl("FileControl", {"PositionX": 106, "PositionY": 290, "Width": 200, "Height": 14, "Text": systemworkpath})  
	addControl("Button", {"PositionX": 106, "PositionY": 320, "Width": 50, "Height": 14, "Label": "~Close dialog", "PushButtonType": 1})  # PushButtonTypeの値はEnumではエラーになる。
	addControl("FixedHyperlink", {"PositionX": 106, "PositionY": 350, "Width": 100, "Height": 14, "Label": "p--q.blogspot.jp", "URL": "https://p--q.blogspot.jp/", "TextColor": 0x3D578C})
	dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
	h = dialog.getModel().getPropertyValue("Height")  # ダイアログの高さをma単位で取得。
	items = ("Introduction", True),\
			("Documents", True) # この順に0からIDがふられる。この順に表示される。
	addControl("Roadmap", {"PositionX": 0, "PositionY": 0, "Width": 85, "Height": h-26, "Complete": False, "Text": "Steps", "Items": items})  # Roadmapコントロールはダイアログウィンドウを描画してからでないと項目が表示されない。
	dialogwindow = dialog.getPeer()  # ダイアログウィンドウ(=ピア）を取得。
	textlistener.setPeer(dialogwindow)  # ダイアログのピアをリスナーに渡す。
	# ノンモダルダイアログにするとき。
# 	showModelessly(ctx, smgr, docframe, dialog)  
	# モダルダイアログにする。フレームに追加するとエラーになる。
	dialog.execute()  
	dialog.dispose()	
class MouseListener(unohelper.Base, XMouseListener):  # Editコントロールではうまく動かない。	
	def __init__(self, ctx, smgr):
		self.pointer = smgr.createInstanceWithContext("com.sun.star.awt.Pointer", ctx)  # ポインタのインスタンスを取得。
	def mousePressed(self, mouseevent):
		pass			
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		control, dummy_controlmodel, name = eventSource(mouseevent)
		if name == "FixedText1":
			self.pointer.setType(REFHAND)  # マウスポインタの種類を設定。
			control.getPeer().setPointer(self.pointer)  # マウスポインタを変更。コントロールからマウスがでるとポインタは元に戻る。
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		pass
class TextListener(unohelper.Base, XTextListener):
	def __init__(self):
		self.vals = {}  # 前値を保存する辞書。
	def setPeer(self, dialogwindow):  # ダイアログのピアを取得。
		self.dialogwindow = dialogwindow
		self.toolkit = dialogwindow.getToolkit()		
	def textChanged(self, textevent):  # 複数回呼ばれるので前値との比較が必要。
		dummy_control, controlmodel, name = eventSource(textevent)	
		val = controlmodel.Value if hasattr(controlmodel, "Value") else controlmodel.Text  # Textが数値の場合は有効桁数が変化するのでValueがあればValueを取得する。
		if name in self.vals:  # 前値の辞書にキーがあるとき
			if val == self.vals[name]:  # 前値と変化がなければなにもしない
				return
		self.vals[name] = val  # 辞書の値を更新。
		if name.startswith("CurrencyField"):	# CurrencyFieldコントロールすべてに対して。
			txt = controlmodel.getPropertyValue("Value")	
			msgbox = self.toolkit.createMessageBox(self.dialogwindow, INFOBOX, BUTTONS_OK, "TextListener", "{} has changed to {}".format(name, txt))  # コントロールのpeerを親にしてもよい。
			msgbox.execute()  # メッセージボックスを表示。
			msgbox.dispose()  # メッセージボックスを破棄。
	def disposing(self, eventobject):
		pass	
class FocusListener(unohelper.Base, XFocusListener):
	def focusGained(self, focusevent):
		dummy_control, controlmodel, name = eventSource(focusevent)
		if name == "Edit1":
			focuschangereason = focusevent.FocusFlags & TAB  # 論理積を取得。
			if focuschangereason==TAB:  # タブで移動してきたとき
				self.echochar = controlmodel.getPropertyValue("EchoChar")  # 伏せ文字を取得。
				controlmodel.setPropertyValue("EchoChar", 0)  # 伏せ文字を解除。
	def focusLost(self, focusevent):  # マウスでフォーカスを移動させたときはこれは呼ばれない。
		dummy_control, controlmodel, name = eventSource(focusevent)		
		if name == "Edit1":
			controlmodel.setPropertyValue("EchoChar", self.echochar)  # 伏せ文字を再設定。
	def disposing(self, eventobject):
		pass  
class KeyListener(unohelper.Base, XKeyListener):
	def __init__(self, dialog):
		dialogmodel = dialog.getModel()
		controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format("FixedText"))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		props = {"Name": "forKeyListener",  "PositionX": 170, "PositionY": 72, "Width": 200, "Height": 12, "Step": 0, "NoLabel": True}
		controlmodel.setPropertyValues(tuple(props.keys()), tuple(props.values()))
		dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		self.control = dialog.getControl(props["Name"])
		self.keycodes = {
			BACKSPACE: "BACKSPACE", 
			SPACE: "SPACE", 
			DELETE: "DELETE",
			LEFT: "LEFT", 
			RIGHT: "RIGHT", 
			HOME: "HOME", 
			END: "END"
			}
		self.reg = re.compile(r"[!\"#$%&'()=~|`{+*}<>?\-\^\\@[;:\],./\\\w]+")  # キーボードの文字を網羅。_は\wに含まれる。
# 	@enableRemoteDebugging
	def keyPressed(self, keyevent):
		dummy_control, dummy_controlmodel, name = eventSource(keyevent)
		if name == "Edit1":
			keycode = keyevent.KeyCode
			if keycode in self.keycodes.keys():
				key = self.keycodes[keycode]		
			else:
				key = keyevent.KeyChar.value
			if self.reg.match(key):  # キーボードにある文字のときのみ表示する。
				self.control.setText("Last Input valid Key: {}".format(key))		
			else:
				self.control.setText("")	
	def keyReleased(self, keyevnet):
		pass
	def disposing(self, eventobject):
		pass  
class SpinListener(unohelper.Base, XSpinListener):
	def up(self, spinevent):
		control, controlmodel, name = eventSource(spinevent)
		controlpeer = control.getPeer()  # コントロールのピアを取得。
		toolkit = controlpeer.getToolkit()  # ピアからツールキットを取得。
		if name == "FormattedField1":
			val = controlmodel.EffectiveValue
			msgbox = toolkit.createMessageBox(controlpeer, INFOBOX, BUTTONS_OK, "SpinListener", "Controlvalue:  {}" .format(val))  # コントロールのpeerを親にしている。
			msgbox.execute()  # メッセージボックスを表示。
			msgbox.dispose()  # メッセージボックスを破棄。
	def down(self, spinevent):
		pass
	def first(self, spinevent):
		pass
	def last(self, spinevent):
		pass
	def disposing(self, eventobject):
		pass  
class ItemListener(unohelper.Base, XItemListener): 
	def __init__(self, dialog):
		self.dialog = dialog
# 	@enableRemoteDebugging
	def itemStateChanged(self, itemevent):
		control, dummy_controlmodel, name = eventSource(itemevent)
		if name == "CheckBox1":
			button = self.dialog.getControl("Button1")
			buttonmodel = button.getModel()
			state = control.getState()
			btnenable = True
			if state==0 or state==2:
				btnenable = False
			buttonmodel.setPropertyValue("Enabled", btnenable)
		elif name == "ComboBox1":  # コンボボックスは選択した文字列が取得できない。
			control.setText(itemevent.Selected)		
	def disposing(self, eventobject):
		pass	  
class AdjustmentListener(unohelper.Base, XAdjustmentListener):	# ブレークするとマウスのクリックが無効になる。
	def __init__(self, ctx, smgr, parentframe):
		self.ctx = ctx
		self.smgr = smgr
		self.adjustmentdialog = None
		self.parentframe = parentframe  # モードレスダイアログの親フレームを取得。
		self.dic = {  # スクロールバーの操作の種類をキーとする。
			ADJUST_ABS.value: "The event has been triggered by dragging the thumb...",
			ADJUST_LINE.value: "The event has been triggered by a single line move..",
			ADJUST_PAGE.value: "The event has been triggered by a block move..."
			}
		self.txts = deque(maxlen=4)  # 要素4個順繰りになる配列を作成。
	def adjustmentValueChanged(self, adjustmentevent):  # 子ダイアログを表示させると2回呼ばれてしまう。
		control, dummy_controlmodel, name = eventSource(adjustmentevent)
		if name == "ScrollBar1":
			adjustmenttype = adjustmentevent.Type.value  # スクロールバーの操作の種類を取得。
			if self.adjustmentdialog is None:  # まだダイアログオブジェクトがないときはダイアログオブジェクトを作成
				controlpeer = control.getPeer()	 # コントロールのピアオプジェクトを取得。			
				toolkit = controlpeer.getToolkit()  # ツールキットを取得。
				self.adjustmentdialog, addControl = dialogCreator(self.ctx, self.smgr, {"PositionX": 150, "PositionY": 150, "Width": 200, "Height": 70, "Title": "AdjustmentListener", "Name": "adjustmentlistenerdialog", "Step": 0, "TabIndex": 0, "Moveable": True})
				self.adjustmentdialog.createPeer(toolkit, controlpeer)  # 新しいダイアログのピアを作成。
				addControl("FixedText", {"PositionX": 10, "PositionY": 8, "Width": 190, "Height": 8, "Step": 0, "NoLabel": True})
				addControl("FixedText", {"PositionX": 10, "PositionY": 16, "Width": 190, "Height": 32, "Step": 0, "NoLabel": True})
				addControl("Button", {"PositionX": 75, "PositionY": 50, "Width": 50, "Height": 14, "Label": "~Close dialog", "PushButtonType": 1})  # PushButtonTypeの値はEnumではエラーになる。
				frame = showModelessly(self.ctx, self.smgr, self.parentframe, self.adjustmentdialog)  # モードレスダイアログとして表示。
				frame.addCloseListener(CloseListener(self))  # モードレスダイアログが閉じられた時はダイアログオブジェクトをクリアする。ダイアログオブジェクトが残っていてもsetVisble(True)ではなぜか表示されない。
				text1 = self.adjustmentdialog.getControl("FixedText1")  # 1行目のコントロールを取得。
				text1.setText(self.dic[adjustmenttype])  # 1行目を代入。
				self.txts.clear()  # 2行目以降にいれるリストをクリア。
			else:  # すでにダイアログオブジェクトがあるとき
				if self.adjustmentdialog.isVisible():  # すでにダイアログが表示されている時
					text1 = self.adjustmentdialog.getControl("FixedText1")  # 1行目のコントロールを取得。
					if text1.getText() != self.dic[adjustmenttype]:  # 1行目について前回と異なるとき
						text1.setText(self.dic[adjustmenttype])  # 1行目を更新。
						self.txts.clear()  # 2行目以降にいれるリストをクリア。
			text2 = self.adjustmentdialog.getControl("FixedText2")  # 2行目以降のコントロールを取得。
			self.txts.append("The value of the scrollbar is: {}".format(adjustmentevent.Value))  # 2行目以降の内容にするリストを取得。
			text2.setText("\n".join(self.txts))  # 2行目を更新。
	def disposing(self, eventobject):
		pass
class CloseListener(unohelper.Base, XCloseListener):  # モードレスダイアログを閉じたときの処理をする。フレームにつける。
	def __init__(self, adjustmentlistener):
		self.adjustmentlistener = adjustmentlistener
	def queryClosing(self, eventobject, getownership):
		pass
	def notifyClosing(self, eventobject):
		if eventobject.Source.getName() == self.adjustmentlistener.adjustmentdialog.getModel().getPropertyValue("Name"):  # フレーム名を確認。
			self.adjustmentlistener.adjustmentdialog.dispose()  # dispose()してもNoneになるわけではない。
			self.adjustmentlistener.adjustmentdialog = None
	def disposing(self, eventobject):
		pass  
def toDate(year, month, day):  # 日付のnamedtupleを返す
	struct = Date(Year=year, Month=month, Day=day)	 # com.sun.star.util.Date
	class StructDate(namedtuple("StructDate", "Date y m d")):
		__slots__ = ()  # インスタンス辞書の作成抑制。
		def __str__(self):  # 文字列として呼ばれた場合に返す値を設定。
			return "{:0>4}-{}-{}".format(self.y, self.m, self.d)
	return StructDate(struct, year, month, day)  # namedtupleを返す
def toTime(hour=0, minute=0, second=0, microsecond=None, tzinfo=None):  # 時刻のnamedtupleを返す。
	microsecond, flg = (0, False) if microsecond is None else (microsecond, True)  # flgはマイクロ秒の表示のためのフラグ。
	tzinfo = False if tzinfo is None else tzinfo
	struct = Time(Hours=hour, Minutes=minute, Seconds=second, NanoSeconds=microsecond*1000, IsUTC=tzinfo) # com.sun.star.util.Time
	class StructTime(namedtuple("StructTime", "Time h m s ms")):
		__slots__ = ()  # インスタンス辞書の作成抑制。
		def __str__(self):  # 文字列として呼ばれた場合に返す値を設定。tzinfoは出力で使っていません。
			return "{:>2}:{:0>2}:{:0>2}.{:0>6}".format(self.h, self.m, self.s, self.ms)	if flg else "{:>2}:{:0>2}:{:0>2}".format(self.h, self.m, self.s)
	return StructTime(struct, hour, minute, second, microsecond)  # namedtupleを返す
def eventSource(event):  # イベントからコントロール、コントロールモデル、コントロール名を取得。
	control = event.Source  # イベントを駆動したコントロールを取得。
	controlmodel = control.getModel()  # コントロールモデルを取得。
	name = controlmodel.getPropertyValue("Name")  # コントロール名を取得。	
	return control, controlmodel, name	
def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションではリスナー動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
	frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
	frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。	
	frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。
	parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
	dialog.setVisible(True)  # ダイアログを見えるようにする。   
	return frame  # フレームにリスナーをつけるときのためにフレームを返す。
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
	dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
	if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
		dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
	dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
	dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
	dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
	dialog.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		control = None
		items, currentitemid = None, None
		if controltype == "Roadmap":  # Roadmapコントロールのとき、Itemsはダイアログモデルに追加してから設定する。そのときはCurrentItemIDもあとで設定する。
			if "Items" in props:  # Itemsはダイアログモデルに追加されてから設定する。
				items = props.pop("Items")
				if "CurrentItemID" in props:  # CurrentItemIDはItemsを追加されてから設定する。
					currentitemid = props.pop("CurrentItemID")
		if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
			control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
			dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
		else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		if items is not None:  # コントロールに追加されたRoadmapモデルにしかRoadmapアイテムは追加できない。
			for i, j in enumerate(items):  # 各Roadmapアイテムについて
				item = controlmodel.createInstance()
				item.setPropertyValues(("Label", "Enabled"), j)
				controlmodel.insertByIndex(i, item)  # IDは0から整数が自動追加される	   
			if currentitemid is not None:  #Roadmapアイテムを追加するとそれがCurrentItemIDになるので、Roadmapアイテムを追加してからCurrentIDを設定する。
				controlmodel.setPropertyValue("CurrentItemID", currentitemid)
		if control is None:  # コントロールがまだインスタンス化されていないとき
			control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if not "Name" in props:
			props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
		controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[controlmodel.setPropertyValue(key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
			else:
				controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
		return controlmodel
	def _generateSequentialName(controltype):  # コントロールの連番名の作成。
		i = 1
		flg = True
		while flg:
			name = "{}{}".format(controltype, i)
			flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return dialog, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
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
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	if not hasattr(doc, "getCurrentController"):  # ドキュメント以外のとき。スタート画面も除外。
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, ())  # Writerのドキュメントを開く。
		while doc is None:  # ドキュメントのロード待ち。
			doc = XSCRIPTCONTEXT.getDocument()
	macro()	   
