#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from calendar import Calendar
from com.sun.star.awt import PosSize  # 定数
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.frame.FrameAction import FRAME_UI_DEACTIVATING  # enum
from com.sun.star.frame import XFrameActionListener
from com.sun.star.awt import Point  # Struct
from com.sun.star.util import MeasureUnit
from com.sun.star.document import XDocumentEventListener
from com.sun.star.style.VerticalAlignment import MIDDLE
from com.sun.star.awt import XActionListener
from com.sun.star.awt import Point  # Struct
from com.sun.star.util import MeasureUnit
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	enhancedmouseclickhandler = EnhancedMouseClickHandler(controller, ctx, smgr, doc)
	controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)  # EnhancedMouseClickHandler
	doc.addDocumentEventListener(DocumentEventListener(enhancedmouseclickhandler))  # DocumentEventListener	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, subj, ctx, smgr, doc):
		self.subj = subj  # disposing()用。コントローラは取得し直さないと最新の画面の状態が反映されない。
		self.args = ctx, smgr, doc
	def mousePressed(self, enhancedmouseevent):
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		ctx, smgr, doc = self.args
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
					controller = doc.getCurrentController()  # 現在のコントローラを取得。
					frame = controller.getFrame()  # フレームを取得。
					componentwindow = frame.getComponentWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
					m = 6  # コントロール間の間隔
					nameX = {"PositionX": m, "Width": 10, "Height": 12, "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}  # 名前Xの共通プロパティ。
					numX = {"PositionX": nameX["PositionX"]+nameX["Width"], "Width": 40, "Height": nameX["Height"], "VerticalAlign": MIDDLE}  # X値入力欄の共通プロパティ。
					unitX = {"PositionX": numX["PositionX"]+numX["Width"], "Width": 32, "Height": nameX["Height"], "Label": "px", "NoLabel": True, "VerticalAlign": MIDDLE}  # 単位の共通プロパティ。
					nameY, numY, unitY = nameX.copy(), numX.copy(), unitX.copy()  # コントロールのプロパティの辞書をコピーする。
					nameY["PositionX"] = unitX["PositionX"] + unitX["Width"]  # 左隣のコントロールのPositionXと幅からPositionXを算出。
					nameY["Label"] = ".Y: " 
					numY["PositionX"] = nameY["PositionX"] + nameY["Width"]
					unitY["PositionX"] = numY["PositionX"] + numY["Width"]
					controls = nameX, numX, unitX, nameY, numY, unitY  # 1行に表示するコントロールのタプル。
					controldialog =  {"PositionX": 100, "PositionY": 40, "Width": unitY["PositionX"]+unitY["Width"]+m, "Title": "Position", "Name": "Position", "Step": 0, "Moveable": True}  # コントロールダイアログのプロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
					dialog, addControl = dialogCreator(ctx, smgr, controldialog)
					# 1行目
					for c in controls:
						c["PositionY"] = m	
					nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
					nameX["Label"] = "EnhancedMouseEvent.X: "
					numX["Text"] = enhancedmouseevent.X
					numY["Text"] = enhancedmouseevent.Y	
					addControl("FixedText", nameX)
					addControl("Edit", numX)  
					addControl("FixedText", unitX)	
					addControl("FixedText", nameY)
					addControl("Edit", numY)  
					addControl("FixedText", unitY)		
					# 2行目
					y = unitY["PositionY"] + unitY["Height"] + m  
					for c in controls:
						c["PositionY"] = y
					nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
					nameX["Label"] = "Target X: "
					point = componentwindow.convertPointToPixel(target.getPropertyValue("Position"), MeasureUnit.MM_100TH)
					numX["Text"] = point.X
					numY["Text"] = point.Y				
					addControl("FixedText", nameX)
					addControl("Edit", numX)  
					addControl("FixedText", unitX)	
					addControl("FixedText", nameY)
					addControl("Edit", numY)  
					addControl("FixedText", unitY)	
					# 3行目
					button = {"PositionY": nameX["PositionY"]+nameX["Height"]+m, "Height": nameX["Height"]+2, "Width": 30, "Label": "~Close", "PushButtonType": 2}  # ボタン。PushButtonTypeの値はEnumではエラーになる。
					button["PositionX"] = unitY["PositionX"] + unitY["Width"] - button["Width"]
					addControl("Button", button)
					dialog.getModel().setPropertyValue("Height", button["PositionY"]+button["Height"]+m)
			
					
					
					
# 					controller = doc.getCurrentController()
# 					frame = controller.getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
# 					containerwindow = frame.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
# 					componentwindow = frame.getComponentWindow()  # コンポーネントウィンドウを取得。
# 					border = controller.getBorder()  # 行ヘッダの高さ(border.Top)、列ヘッダの幅(border.Left)を取得できる。
					# enhancedmouseeventから取得できる座標は行と列のヘッダとの境界からの相対位置。
# 					x = enhancedmouseevent.X + border.Left
# 					y = enhancedmouseevent.Y + border.Top
					# コンテナウィンドウの位置を利用。
# 					containerwindowpossize= containerwindow.getPosSize()  # コンテナウィンドウの左上角の画面に対する相対位置が返る。
# 					x += containerwindowpossize.X
# 					y += containerwindowpossize.Y						
					# コンテナウィンドウのAccessibleContextの位置を利用。
#  					containerwindowlocationon = containerwindow.getAccessibleContext().getLocation()
# 					x += containerwindowlocationon.X
# 					y += containerwindowlocationon.Y
					# コンポーネントウィンドウの位置を利用。 
# 					componentwindowpossize= componentwindow.getPosSize()  # ツールバーも含んだ位置を返す。
# 					x += componentwindowpossize.X
# 					y += componentwindowpossize.Y				
					# コンポーネントウィンドウのAccessibleContextの位置を利用。
# 					componentwindowlocationon = componentwindow.getAccessibleContext().getLocation()  
# 					x += componentwindowlocationon.X
# 					y += componentwindowlocationon.Y  # enhancedmouseevent.Yは数式バーを含まないコンポーネントウィンドウの相対座標なのでY軸は上にずれる。					
		
# 					point = componentwindow.convertPointToLogic(Point(X=x, Y=y), MeasureUnit.APPFONT)  # ピクセル単位をma単位に変換。

					
# 					pixelpoint = containerwindow.convertPointToPixel(target.getPropertyValue("Position"), MeasureUnit.MM_100TH)  # ピクセル単位をma単位に変換。
# 					point = point = containerwindow.convertPointToLogic(pixelpoint, MeasureUnit.APPFONT)  # ピクセル単位をma単位に変換。
# 					
# 					dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": point.X, "PositionY": point.Y, "Width": 100, "Height": 100, "Title": "ノンモダルダイアログ", "Name": "NoneModalDialog", "Step": 0, "Moveable": True})  # PositionXとPositionYはそれぞれ親ウィンドウの左端と上端からの相対位置。
					toolkit = componentwindow.getToolkit()  # ピアからツールキットを取得。コンテナウィンドウでもコンポーネントウィンドウでも結果は同じ。
					dialog.createPeer(toolkit, componentwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
					showModelessly(ctx, smgr, frame, dialog)  # ノンモダルダイアログとして表示。ダイアログのフレームを取得。
					return False  # セル編集モードにしない。
		return True  # セル編集モードにする。
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):
		self.subj.removeEnhancedMouseClickHandler(self)
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
	def __init__(self, enhancedmouseclickhandler):
		self.args = enhancedmouseclickhandler
	def documentEventOccured(self, documentevent):
		enhancedmouseclickhandler = self.args
		if documentevent.EventName=="OnUnload":  
			source = documentevent.Source
			source.removeEnhancedMouseClickHandler(enhancedmouseclickhandler)
			source.removeDocumentEventListener(self)
	def disposing(self, eventobject):
		eventobject.Source.removeDocumentEventListener(self)
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
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
		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
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
		XSCRIPTCONTEXT = createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
	#  doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
		if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
			XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
		flg = True
		while flg:
			doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
			if doc is not None:
				flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
		return XSCRIPTCONTEXT
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。 
	macro()  # マクロの実行。