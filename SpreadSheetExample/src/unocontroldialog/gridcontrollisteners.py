#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import datetime
from com.sun.star.style.VerticalAlignment import MIDDLE
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XMouseListener
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import ScrollBarOrientation  # 定数
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.awt import XTopWindowListener
from com.sun.star.frame.FrameAction import COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.frame import XFrameActionListener
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	frame = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = frame.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  
	m = 6  # コントロール間の間隔
	grid = {"PositionX": m, "PositionY": m, "Width": 145, "Height": 100, "ShowColumnHeader": True, "ShowRowHeader": True}  # グリッドコントロールの基本プロパティ。
	label = {"PositionX": m, "Width": 45, "Height": 12, "Label": "Date and time: ", "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}  # ラベルフィールドコントロールの基本プロパティ。
	x = label["PositionX"]+label["Width"]  # ラベルフィールドコントロールの右端。
	textbox = {"PositionX": x, "Width": grid["PositionX"]+grid["Width"]-x, "Height": label["Height"], "VerticalAlign": MIDDLE}  # テクストボックスコントロールの基本プロパティ。
	button = {"PositionX": m, "Width": 30, "Height": label["Height"]+2, "PushButtonType": 0}  # ボタンの基本プロパティ。PushButtonTypeの値はEnumではエラーになる。
	controldialog =  {"PositionX": 100, "PositionY": 40, "Width": grid["PositionX"]+grid["Width"]+m, "Title": "Grid Example", "Name": "controldialog", "Step": 0, "Moveable": True}  # コントロールダイアログの基本プロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
	dialog, addControl = dialogCreator(ctx, smgr, controldialog)  # コントロールダイアログの作成。
	grid1 = addControl("Grid", grid, {"addMouseListener": MouseListener(doc)})  # グリッドコントロールの取得。
	gridmodel = grid1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	column0 = gridcolumn.createColumn()  # 列の作成。
	column0.Title = "Date"  # 列ヘッダー。
	column0.ColumnWidth = 60  # 列幅。
	gridcolumn.addColumn(column0)  # 列を追加。
	column1 = gridcolumn.createColumn()  # 列の作成。
	column1.Title = "Time"  # 列ヘッダー。
	column1.ColumnWidth = grid["Width"] - column0.ColumnWidth  #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
	gridcolumn.addColumn(column1)  # 列を追加。	
	griddata = grid1.getModel().getPropertyValue("GridDataModel")  # GridDataModel
	now = datetime.now()  # 現在の日時を取得。
	griddata.addRow(0, (now.date().isoformat(), now.time().isoformat()))  # グリッドに行を挿入。
	y = grid["PositionY"] + grid["Height"] + m  # 下隣のコントロールのY座標を取得。
	label["PositionY"] = textbox["PositionY"] = y
	textbox["Text"] = now.isoformat().replace("T", " ")
	addControl("FixedText", label)
	addControl("Edit", textbox)  
	y = label["PositionY"] + label["Height"] + m  # 下隣のコントロールのY座標を取得。 
	button1, button2 = button.copy(), button.copy()
	button1["PositionY"] = button2["PositionY"]  = y
	button1["Label"] = "~Now"
	button2["Label"] = "~Close"
	button2["PushButtonType"] = 2  # CANCEL		
	button2["PositionX"] = grid["Width"] - button2["Width"]
	button1["PositionX"] = button2["PositionX"] - m - button1["Width"]
	addControl("Button", button1, {"setActionCommand": "now" ,"addActionListener": ActionListener()})  
	addControl("Button", button2)  
	dialog.getModel().setPropertyValue("Height", button1["PositionY"]+button1["Height"]+m)  # コントロールダイアログの高さを設定。
	dialog.createPeer(toolkit, containerwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
	# ノンモダルダイアログにするとき。オートメーションでは動かない。
	dialogframe = showModelessly(ctx, smgr, frame, dialog)  
	dialog.addTopWindowListener(TopWindowListener())
	dialogframe.addFrameActionListener(FrameActionListener())  # FrameActionListener 
	dialogframe.addCloseListener(CloseListener())  # CloseListener
	
	# モダルダイアログにする。フレームに追加するとエラーになる。
# 	dialog.execute()  
# 	dialog.dispose()	

import os, inspect
from datetime import datetime
C = 100  # カウンターの初期値。
TIMESTAMP = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
def createLog(source, filename, txt):  # 年月日T時分秒リスナーのインスタンス名_メソッド名(_オプション).logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	path = doc.getURL() if __file__.startswith("vnd.sun.star.tdoc:") else __file__  # このスクリプトのパス。fileurlで返ってくる。埋め込みマクロの時は埋め込んだドキュメントのURLで代用する。
	thisscriptpath = unohelper.fileUrlToSystemPath(path)  # fileurlをsystempathに変換。
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	name = source.getImplementationName().split(".")[-1]
	global C
	filename = "".join((TIMESTAMP, "_", str(C), "{}_{}".format(name, filename), ".log"))
	C += 1
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)
		
		
class CloseListener(unohelper.Base, XCloseListener):
	def queryClosing(self, eventobject, getsownership):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "getsownership: {}\nSource: {}".format(getsownership, eventobject.Source))	
	def notifyClosing(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def disposing(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def __init__(self):
		enums = COMPONENT_ATTACHED, COMPONENT_DETACHING, COMPONENT_REATTACHED, FRAME_ACTIVATED, FRAME_DEACTIVATING, CONTEXT_CHANGED, FRAME_UI_ACTIVATED, FRAME_UI_DEACTIVATING  # enum
		frameactionnames = "COMPONENT_ATTACHED", "COMPONENT_DETACHING", "COMPONENT_REATTACHED", "FRAME_ACTIVATED", "FRAME_DEACTIVATING", "CONTEXT_CHANGED", "FRAME_UI_ACTIVATED", "FRAME_UI_DEACTIVATING"
		self.args = zip(enums, frameactionnames)
	def frameAction(self, frameactionevent):
		frameactions = self.args
		frameaction = frameactionevent.Action
		for enum, frameactionname in frameactions:
			if frameaction==enum:
				createLog(frameactionevent.Source, "{}_{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name, frameactionname), "FrameAction: {}\nSource: {}".format(frameactionname, frameactionevent.Source))	
				return
	def disposing(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
class TopWindowListener(unohelper.Base, XTopWindowListener):
	def windowOpened(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def windowClosing(self, eventobject):  # 呼ばれない。いつ呼ばれる?
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def windowClosed(self, eventobject):  # フレームが閉じた後デスクトップが閉じる前に呼ばれる。
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def windowMinimized(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def windowNormalized(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def windowActivated(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def windowDeactivated(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
	def disposing(self, eventobject):
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, doc):
		self.args = doc
	def mousePressed(self, mouseevent):
		if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。
			doc = self.args
			selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
				source = mouseevent.Source  # グリッドコントロールを取得。
				griddata = source.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
				rowdata = griddata.getRowData(source.getCurrentRow())  # グリッドコントロールで選択している行のすべての列をタプルで取得。
				selection.setString(" ".join(rowdata))  # 選択セルに書き込む。
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):  # モードレスダイアログにした時は呼ばれない。
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
		eventobject.Source.removeMouseListener(self)
class ActionListener(unohelper.Base, XActionListener):
	def actionPerformed(self, actionevent):
		cmd = actionevent.ActionCommand
		source = actionevent.Source  # ボタンコントロールが返る。
		context = source.getContext()  # コントロールダイアログが返ってくる。
		if cmd == "now":
			now = datetime.now()  # 現在の日時を取得。
			context.getControl("Edit1").setText(now.isoformat().replace("T", " "))  # テキストボックスコントロールに入力。
			grid1 = context.getControl("Grid1")  # グリッドコントロールを取得。
			griddata = grid1.getModel().getPropertyValue("GridDataModel")  # グリッドコントロールモデルからDefaultGridDataModelを取得。
			griddata.addRow(griddata.RowCount, (now.date().isoformat(), now.time().isoformat()))  # 新たな行を追加。
			accessiblecontext = grid1.getAccessibleContext()  # グリッドコントロールのAccessibleContextを取得。
			for i in range(accessiblecontext.getAccessibleChildCount()):  # 子要素をのインデックスを走査する。
				child = accessiblecontext.getAccessibleChild(i)  # 子要素を取得。
				if child.getAccessibleContext().getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
					if child.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
						child.setValue(child.getMaximum())  # 最大値にスクロールさせる。
						return
	def disposing(self, eventobject):  # モードレスダイアログの時は呼ばれない。
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
		eventobject.Source.removeActionListener(self)
def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションでは動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
	frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
	frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。	
	frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。ｽﾍﾟｰｽは不可。
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
	