#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.awt import XMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.frame.FrameAction import FRAME_UI_DEACTIVATING  # enum
from com.sun.star.frame import XFrameActionListener
from com.sun.star.util import MeasureUnit
from com.sun.star.document import XDocumentEventListener
from com.sun.star.style.VerticalAlignment import MIDDLE
from com.sun.star.awt import XActionListener
from com.sun.star.awt import Point  # Struct
from com.sun.star.util import XCloseListener
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import Key  # 定数
from com.sun.star.awt import Selection  # Struct
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	mouseclickhandler = MouseClickHandler(controller, ctx, smgr, doc)
	controller.addMouseClickHandler(mouseclickhandler)  # EnhancedMouseClickHandler
	doc.addDocumentEventListener(DocumentEventListener(mouseclickhandler))  # DocumentEventListener 
class MouseClickHandler(unohelper.Base, XMouseClickHandler):
	def __init__(self, subj, ctx, smgr, doc):
		self.subj = subj  # disposing()用。コントローラは取得し直さないと最新の画面の状態が反映されない。
		self.args = ctx, smgr, doc
	def mousePressed(self, mouseevent):
		ctx, smgr, doc = self.args
		target = doc.getCurrentSelection()  # ターゲットのセルを取得。
		if mouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if mouseevent.ClickCount==2:  # ダブルクリックの時
					controller = doc.getCurrentController()  # 現在のコントローラを取得。分割しているのとしていないシートで発火しないことがある問題はself.subjでも解決しない。
					frame = controller.getFrame()  # フレームを取得。
					containerwindow = frame.getContainerWindow()  # コンテナウィドウの取得。
					framepointonscreen = containerwindow.getAccessibleContext().getAccessibleParent().getAccessibleContext().getLocationOnScreen()  # フレームの左上角の点（画面の左上角が原点)。
					componentwindow = frame.getComponentWindow()  # コンポーネントウィンドウを取得。
					sourcepointonscreen = mouseevent.Source.getAccessibleContext().getLocationOnScreen()  # クリックした枠の左上の点（画面の左上角が原点)。
					x = sourcepointonscreen.X + mouseevent.X - framepointonscreen.X
					y = sourcepointonscreen.Y + mouseevent.Y - framepointonscreen.Y
					dialogpoint = componentwindow.convertPointToLogic(Point(X=x, Y=y), MeasureUnit.APPFONT)  # ピクセル単位をma単位に変換。
					actionlistener = ActionListener(target)
					keylistener = KeyListener(target)
					m = 6  # コントロール間の間隔
					name = {"PositionX": m, "Width": 50, "Height": 12, "NoLabel": True, "Align": 0, "VerticalAlign": MIDDLE}  
					address = {"PositionX": m, "Width": 50, "Height": name["Height"], "VerticalAlign": MIDDLE}  
					controldialog =  {"PositionX": dialogpoint.X, "PositionY": dialogpoint.Y, "Width": address["PositionX"]+address["Width"]+m, "Title": "Popup Window", "Name": "PopupWindow", "Step": 0, "Moveable": True}  # コントロールダイアログのプロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
					dialog, addControl = dialogCreator(ctx, smgr, controldialog)
					name["PositionY"] = m
					name["Label"] = "Target Address"
					addControl("FixedText", name)
					address["PositionY"] = name["PositionY"] + name["Height"] + m
					stringaddress = getStringAddressFromCellRange(target)
					address["Text"] = stringaddress
					textlength = len(stringaddress)
					edit1selection = Selection(Min=textlength, Max=textlength)  # カーソルの位置を最後にする。指定しないと先頭になる。
					edit1 = addControl("Edit", address, {"addKeyListener": keylistener})
					button = {"PositionY": m, "Height": name["Height"]+2, "Width": 40, "Label": "~To Target", "PushButtonType": 0}  # ボタン。PushButtonTypeの値はEnumではエラーになる。
					button["PositionX"] = address["PositionX"] + address["Width"] - button["Width"]
					button["PositionY"] = address["PositionY"] + address["Height"] + m
					addControl("Button", button, {"setActionCommand": "totarget" ,"addActionListener": actionlistener}) 
					dialog.getModel().setPropertyValue("Height", button["PositionY"]+button["Height"]+m)  # コントロールダイアログの高さを設定。
					toolkit = componentwindow.getToolkit()  # ピアからツールキットを取得。
					dialog.createPeer(toolkit, componentwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
					edit1.setSelection(edit1selection)  # テクストボックスコントロールのカーソルの位置を変更。ピア作成後でないと反映されない。
					dialogframe = showModelessly(ctx, smgr, frame, dialog)  # ノンモダルダイアログとして表示。ダイアログのフレームを取得。
					actionlistener.frame = dialogframe
					keylistener.frame = dialogframe
					frameactionlistener = FrameActionListener()
					dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListener。ウィンドウを閉じる時はそのコンテナウィンドウをsetVisible(False)にするのでdialogを渡しておく。
					dialogframe.addCloseListener(CloseListener(dialog, frameactionlistener, actionlistener, keylistener))  # CloseListener
					return True  # セル編集モードにしない。
		return False  # セル編集モードにする。
	def mouseReleased(self, mouseevent):
		return False  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):
		self.subj.removeMouseClickHandler(self)
class KeyListener(unohelper.Base, XKeyListener):
	def __init__(self, target):
		self.args = target
		self.frame = None
	def keyPressed(self, keyevent):
		if keyevent.KeyCode==Key.RETURN:  # リターンキーが押された時。
			target = self.args
			source = keyevent.Source  # テキストボックスコントロールが返る。
			context = source.getContext()  # コントロールダイアログが返ってくる。
			target.setString(context.getControl("Edit1").getText())  # テキストボックスコントロールの内容を選択セルに代入する。
			self.frame.close(True)
	def keyReleased(self, keyevnet):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeKeyListener(self)
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, dialog, frameactionlistener, actionlistener, keylistener):
		self.args = dialog, frameactionlistener, actionlistener, keylistener
	def queryClosing(self, eventobject, getsownership):  # コントロールダイアログを閉じると直前。フレーム削除する。
		dialog, frameactionlistener, actionlistener, keylistener = self.args
		dialog.getControl("Button1").removeActionListener(actionlistener)
		dialog.getControl("Edit1").removeKeyListener(keylistener)
		eventobject.Source.removeFrameActionListener(frameactionlistener)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		eventobject.Source.removeCloseListener(self)
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, target):
		self.args = target
		self.frame = None
	def actionPerformed(self, actionevent):
		target = self.args
		cmd = actionevent.ActionCommand
		source = actionevent.Source  # ボタンコントロールが返る。
		context = source.getContext()  # コントロールダイアログが返ってくる。
		if cmd == "totarget":
			target.setString(context.getControl("Edit1").getText())  # テキストボックスコントロールの内容を選択セルに代入する。
			self.frame.close(True)
	def disposing(self, eventobject):
		eventobject.Source.removeActionListener(self)
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def frameAction(self, frameactionevent):
		if frameactionevent.Action==FRAME_UI_DEACTIVATING:  # フレームがアクティブでなくなった時。TopWindowListenerのwindowDeactivated()だとウィンドウタイトルバーをクリックしただけで発火してしまう。
			frameactionevent.Frame.removeFrameActionListener(self)  # フレームにつけたリスナーを除去。
			frameactionevent.Frame.close(True)
	def disposing(self, eventobject):
		eventobject.Source.removeFrameActionListener(self)
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, mouseclickhandler):
		self.args = mouseclickhandler
	def documentEventOccured(self, documentevent):
		mouseclickhandler = self.args
		if documentevent.EventName=="OnUnload":  # Calcドキュメントを閉じる時。リスナーを削除する。
			source = documentevent.Source
			source.removeMouseClickHandler(mouseclickhandler)
			source.removeDocumentEventListener(self)
	def disposing(self, eventobject):
		eventobject.Source.removeDocumentEventListener(self)
def getStringAddressFromCellRange(source):  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。文字列アドレスが取得できないオブジェクトの時はオブジェクトの文字列を返す。 
	stringaddress = ""
	propertysetinfo = source.getPropertySetInfo()  # PropertySetInfo
	if propertysetinfo.hasPropertyByName("AbsoluteName"):  # AbsoluteNameプロパティがある時。
		absolutename = source.getPropertyValue("AbsoluteName") # セル範囲コレクションは$Sheet1.$A$4:$A$6,$Sheet1.$B$4という形式で返る。
		names = absolutename.replace("$", "").split(",")  # $を削除してセル範囲の文字列アドレスのリストにする。
		stringaddress = ", ".join(names)  # コンマでつなげる。
	return stringaddress
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
			if controltype=="Grid":
				control = smgr.createInstanceWithContext("com.sun.star.awt.grid.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			else:	
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
		if controltype=="Grid":
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.grid.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		else:	
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
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
