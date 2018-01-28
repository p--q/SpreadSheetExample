#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.style.VerticalAlignment import MIDDLE
from com.sun.star.awt import XActionListener
from com.sun.star.awt import Point  # Struct
from com.sun.star.util import MeasureUnit
def macro(documentevent=None):  # 引数は文書のイベント駆動用。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。  
	m = 6  # コントロール間の間隔
	nameX = {"PositionX": m, "Width": 10, "Height": 12, "Label": "X: ", "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}  # 名前Xの共通プロパティ。
	numX = {"PositionX": nameX["PositionX"]+nameX["Width"], "Width": 40, "Height": nameX["Height"], "VerticalAlign": MIDDLE}  # X値入力欄の共通プロパティ。
	unitX = {"PositionX": numX["PositionX"]+numX["Width"], "Width": 32, "Height": nameX["Height"], "NoLabel": True, "VerticalAlign": MIDDLE}  # 単位の共通プロパティ。
	nameY, numY, unitY = nameX.copy(), numX.copy(), unitX.copy()  # コントロールのプロパティの辞書をコピーする。
	nameY["PositionX"] = unitX["PositionX"] + unitX["Width"]  # 左隣のコントロールのPositionXと幅からPositionXを算出。
	nameY["Label"] = "Y: " 
	numY["PositionX"] = nameY["PositionX"] + nameY["Width"]
	unitY["PositionX"] = numY["PositionX"] + numY["Width"]
	controls = nameX, numX, unitX, nameY, numY, unitY  # 1行に表示するコントロールのタプル。
	button = {"Height": nameX["Height"]+2, "PushButtonType": 0}  # ボタンの共通プロパティ。PushButtonTypeの値はEnumではエラーになる。
	controldialog =  {"PositionX": 100, "PositionY": 40, "Width": unitY["PositionX"]+unitY["Width"]+m, "Title": "Unit Converter", "Name": "ConvertUnits", "Step": 0, "Moveable": True}  # コントロールダイアログのプロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
	dialog, addControl = dialogCreator(ctx, smgr, controldialog)
	# 1行目
	for c in controls:
		c["PositionY"] = m	
	nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
	unitX["Label"] = unitY["Label"] = "px"
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
	unitX["Label"] = unitY["Label"] = "ma"
	addControl("FixedText", nameX)
	addControl("Edit", numX)  
	addControl("FixedText", unitX)	
	addControl("FixedText", nameY)
	addControl("Edit", numY)  
	addControl("FixedText", unitY)	
	# 3行目
	y = unitY["PositionY"] + unitY["Height"] + m  
	for c in controls:
		c["PositionY"] = y
	nameX, numX, unitX, nameY, numY, unitY = [c.copy() for c in controls]  # addControlに渡した辞書は変更されるのでコピーを渡す。
	unitX["Label"] = unitY["Label"] = "1/100mm"
	addControl("FixedText", nameX)
	addControl("Edit", numX)  
	addControl("FixedText", unitX)	
	addControl("FixedText", nameY)
	addControl("Edit", numY)  
	addControl("FixedText", unitY)	
	# 4行目
	y = unitY["PositionY"] + unitY["Height"] + m  
	message = {"Name": "Message", "PositionX": m, "PositionY": y, "Width": controldialog["Width"]-m*2, "Height": 12, "Label": "Pass any one of units and push Convert", "NoLabel": True}
	addControl("FixedText", message)
	# 5行目
	button1, button2 = button.copy(), button.copy()
	button1["PositionY"] = button2["PositionY"] = message["PositionY"] + message["Height"] + m  
	button1["Width"] = 40
	button1["Label"] = "Con~vert"
	button2["Width"] = 30
	button2["Label"] = "~Clear"	
	button2["PositionX"] = unitY["PositionX"] + unitY["Width"] - button2["Width"]
	button1["PositionX"] = button2["PositionX"] - m - button1["Width"]
	actionlistener = ActionListener()
	addControl("Button", button1, {"setActionCommand": "convert" ,"addActionListener": actionlistener})
	addControl("Button", button2, {"setActionCommand": "clear" ,"addActionListener": actionlistener})
	dialog.getModel().setPropertyValue("Height", button1["PositionY"]+button1["Height"]+m)
	dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
	# ノンモダルダイアログにするとき。オートメーションでは動かない。
# 	showModelessly(ctx, smgr, docframe, dialog)   # リスナー削除は未処理。
	# モダルダイアログにする。フレームに追加するとエラーになる。
	dialog.execute()  
	dialog.dispose()	
class ActionListener(unohelper.Base, XActionListener):
	def actionPerformed(self, actionevent):
		cmd = actionevent.ActionCommand
		source = actionevent.Source  # ボタンコントロールが返る。
		context = source.getContext()  # コントロールダイアログが返ってくる。
		message = context.getControl("Message")
		if cmd == "convert":
			edits = [context.getControl("Edit{}".format(i)) for i in range(1, 7)]
			edittxts = [e.getText() for e in edits]
			if ["".join((edittxts[i], edittxts[i+1])) for i in range(0, 6, 2)].count("")==2:  # 2行だけテキストボックスが空文字の時。
				pxToma, pxTomm, maTopx, maTomm, mmTopx, mmToma = createConverters(context)
				edit11, edit12, edit21, edit22, edit31, edit32 = edits
				for i in range(0, 6, 2):  # 各行について。
					x, y = edittxts[i], edittxts[i+1]
					if x or y:  # いずれかは空文字でない時。
						x = int(x) if x.isdigit() else 0  # テキストボックスの文字が数字の時は整数に変換。数字でなければ0にする。
						y = int(y) if y.isdigit() else 0
						if i==0:  # 1行目に数字が入力されている時。
							edit11.setText(x)
							edit12.setText(y)							
							maX, maY = pxToma(x, y)
							edit21.setText(maX)
							edit22.setText(maY)
							mmX, mmY = pxTomm(x, y)
							edit31.setText(mmX)
							edit32.setText(mmY)
							u = "px"
						elif i==2:  # 2行目に数字が入力されている時。
							pxX, pxY = maTopx(x, y)
							edit11.setText(pxX)
							edit12.setText(pxY)
							edit21.setText(x)
							edit22.setText(y)								
							mmX, mmY = maTomm(x, y)
							edit31.setText(mmX)
							edit32.setText(mmY)		
							u = "ma"
						elif i==4:  # 3行目に数字が入力されている時。
							pxX, pxY = mmTopx(x, y)
							edit11.setText(pxX)										
							edit12.setText(pxY)
							maX, maY = mmToma(x, y)
							edit21.setText(maX)
							edit22.setText(maY)	
							edit31.setText(x)
							edit32.setText(y)								
							u = "mm"	
						message.setText("{} converted to another unit".format(u))		
						message.getModel().setPropertyValue("TextColor", None)
						return		
			else:  # パラメーターが1つだけでない時。
				message.setText("Too many inputs or no input")
				message.getModel().setPropertyValue("TextColor", 0xFF0000)
		elif cmd == "clear":
			[context.getControl("Edit{}".format(i)).setText("") for i in range(1, 7)]
			message.setText("Pass any one of units and push Convert")
			message.getModel().setPropertyValue("TextColor", None)
	def disposing(self, eventobject):
		eventobject.Source.removeActionListener(self)
def createConverters(window):
	def maTopx(x, y):  # maをpxに変換する。
		point = window.convertPointToPixel(Point(X=x, Y=y), MeasureUnit.APPFONT)
		return point.X, point.Y
	def mmTopx(x, y):  # 1/100mmをpxに変換する。
		point = window.convertPointToPixel(Point(X=x, Y=y), MeasureUnit.MM_100TH)
		return point.X, point.Y
	def pxToma(x, y):  # pxをmaに変換する。
		point = window.convertPointToLogic(Point(X=x, Y=y), MeasureUnit.APPFONT)
		return point.X, point.Y
	def pxTomm(x, y):  # pxをmmに変換する。
		point = window.convertPointToLogic(Point(X=x, Y=y), MeasureUnit.MM_100TH)
		return point.X, point.Y
	def mmToma(x, y):  # 1/100mmをmaに変換する。
		return pxToma(*mmTopx(x, y))
	def maTomm(x, y):  # maをmmに変換する。
		return pxTomm(*maTopx(x, y))
	return pxToma, pxTomm, maTopx, maTomm, mmTopx, mmToma
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
