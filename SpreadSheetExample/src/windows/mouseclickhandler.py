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
	def mousePressed(self, MouseEvent):
		ctx, smgr, doc = self.args
		target = doc.getCurrentSelection()  # ターゲットのセルを取得。
		if MouseEvent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if MouseEvent.ClickCount==2:  # ダブルクリックの時
					controller = doc.getCurrentController()  # 現在のコントローラを取得。
					frame = controller.getFrame()  # フレームを取得。
					componentwindow = frame.getComponentWindow()  # コンポーネントウィンドウを取得。
					source = MouseEvent.Source  # クリックした枠のコンポーネントウィンドウが返る。
					point = componentwindow.convertPointToLogic(Point(X=MouseEvent.X, Y=MouseEvent.Y), MeasureUnit.APPFONT)  # EnhancedMouseClickHandlerの座標をmaに変換。
					# コントロールダイアログの左上の座標を設定。
					dialogX = point.X
					dialogY = point.Y
					m = 6  # コントロール間の間隔
					nameX = {"PositionX": m, "Width": 105, "Height": 12, "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}  # 名前Xの共通プロパティ。
					numX = {"PositionX": nameX["PositionX"]+nameX["Width"], "Width": 20, "Height": nameX["Height"], "VerticalAlign": MIDDLE}  # X値入力欄の共通プロパティ。
					unitX = {"PositionX": numX["PositionX"]+numX["Width"], "Width": 10, "Height": nameX["Height"], "Label": "px", "NoLabel": True, "VerticalAlign": MIDDLE}  # 単位の共通プロパティ。
					nameY, numY, unitY = nameX.copy(), numX.copy(), unitX.copy()  # コントロールのプロパティの辞書をコピーする。
					nameY["PositionX"] = unitX["PositionX"] + unitX["Width"]  # 左隣のコントロールのPositionXと幅からPositionXを算出。
					nameY["Width"] = 10
					numY["PositionX"] = nameY["PositionX"] + nameY["Width"]
					unitY["PositionX"] = numY["PositionX"] + numY["Width"]
					controls = nameX, numX, unitX, nameY, numY, unitY  # 1行に表示するコントロールのタプル。
					controldialog =  {"PositionX": dialogX, "PositionY": dialogY, "Width": unitY["PositionX"]+unitY["Width"]+m, "Title": "Position", "Name": "Position", "Step": 0, "Moveable": True}  # コントロールダイアログのプロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
					dialog, addControl = dialogCreator(ctx, smgr, controldialog)
					# 1行目
					for c in controls:
						c["PositionY"] = m	
					nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
					nameX["Label"] = "MouseEvent.X: "
					numX["Text"] = str(MouseEvent.X)  # プロパティに代入するときは文字列に変更必要。
					nameY["Label"] = ".Y: " 
					numY["Text"] = str(MouseEvent.Y)
					addControl("FixedText", nameX)
					addControl("Edit", numX)  
					addControl("FixedText", unitX)	
					addControl("FixedText", nameY)
					addControl("Edit", numY)  
					addControl("FixedText", unitY)		
					# 2行目
					y = nameX["PositionY"] + nameX["Height"] + m  
					for c in controls:
						c["PositionY"] = y
					nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
					nameX["Label"] = "Target X: "
					point = componentwindow.convertPointToPixel(target.getPropertyValue("Position"), MeasureUnit.MM_100TH)  # クリックしたセルの左上角の座標。1/100mmをpxに変換。
					numX["Text"] = str(point.X)
					nameY["Label"] = "Y: " 
					numY["Text"] = str(point.Y)				
					addControl("FixedText", nameX)
					addControl("Edit", numX)  
					addControl("FixedText", unitX)	
					addControl("FixedText", nameY)
					addControl("Edit", numY)  
					addControl("FixedText", unitY)	
					# 3行目
					y = nameX["PositionY"] + nameX["Height"] + m  
					for c in controls:
						c["PositionY"] = y
					nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
					nameX["Label"] = "Source.getPosSize().X: "
					possize = source.getPosSize()  # コンポーネントウィンドウのPosSize。
					point = componentwindow.convertPointToPixel(target.getPropertyValue("Position"), MeasureUnit.MM_100TH)  # クリックしたセルの左上角の座標。1/100mmをpxに変換。
					numX["Text"] = str(possize.X)
					nameY["Label"] = ".Y: " 
					numY["Text"] = str(possize.Y)				
					addControl("FixedText", nameX)
					addControl("Edit", numX)  
					addControl("FixedText", unitX)	
					addControl("FixedText", nameY)
					addControl("Edit", numY)  
					addControl("FixedText", unitY)		
					# 4行目
					y = nameX["PositionY"] + nameX["Height"] + m  
					for c in controls:
						c["PositionY"] = y
					nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
					nameX["Label"] = "AccessibleContext.getLocation().X: "
					accessiblecontext = source.getAccessibleContext()  # コンポーネントウィンドウのAccessibleContextを取得。
					point = accessiblecontext.getLocation()  # 位置を取得。
					numX["Text"] = str(point.X)
					nameY["Label"] = ".Y: " 
					numY["Text"] = str(point.Y)				
					addControl("FixedText", nameX)
					addControl("Edit", numX)  
					addControl("FixedText", unitX)	
					addControl("FixedText", nameY)
					addControl("Edit", numY)  
					addControl("FixedText", unitY)										
					# 5行目
					button = {"PositionY": nameX["PositionY"]+nameX["Height"]+m, "Height": nameX["Height"]+2, "Width": 30, "Label": "~Close", "PushButtonType": 2}  # ボタン。PushButtonTypeの値はEnumではエラーになる。
					button["PositionX"] = unitY["PositionX"] + unitY["Width"] - button["Width"]
					addControl("Button", button)
					dialog.getModel().setPropertyValue("Height", button["PositionY"]+button["Height"]+m)
					toolkit = componentwindow.getToolkit()  # ピアからツールキットを取得。コンテナウィンドウでもコンポーネントウィンドウでも結果は同じ。
					dialog.createPeer(toolkit, componentwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
					showModelessly(ctx, smgr, frame, dialog)  # ノンモダルダイアログとして表示。ダイアログのフレームを取得。
					return True  # セル編集モードにしない。
		return False  # セル編集モードにする。
	def mouseReleased(self, mouseevent):
		return False  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):
		self.subj.removeMouseClickHandler(self)
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
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, mouseclickhandler):
		self.args = mouseclickhandler
	def documentEventOccured(self, documentevent):
		mouseclickhandler = self.args
		if documentevent.EventName=="OnUnload":  
			source = documentevent.Source
			source.removeMouseClickHandler(mouseclickhandler)
			source.removeDocumentEventListener(self)
	def disposing(self, eventobject):
		eventobject.Source.removeDocumentEventListener(self)
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
