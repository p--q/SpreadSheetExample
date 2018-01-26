#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.style.VerticalAlignment import MIDDLE
from com.sun.star.awt import XActionListener
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
def macro(documentevent=None):  # 引数は文書のイベント駆動用。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。  
	m = 6  # 垂直マージン
	n = 5  # 行数
	name = {"PositionX": m, "Width": 42, "Height": 12, "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}  # 単位名の共通プロパティ。
	num = {"PositionX": name["PositionX"]+name["Width"], "Width": 40, "Height": name["Height"], "VerticalAlign": MIDDLE}  # 値入力欄の共通プロパティ。
	unit = {"PositionX": num["PositionX"]+num["Width"], "Width": 32, "Height": name["Height"], "NoLabel": True, "VerticalAlign": MIDDLE}  # 単位の共通プロパティ。
	button = {"Height": name["Height"]+2, "PushButtonType": 0}  # ボタンの共通プロパティ。PushButtonTypeの値はEnumではエラーになる。
	controldialog =  {"PositionX": name["PositionX"], "PositionY": 40, "Width": unit["PositionX"]+unit["Width"]+m, "Title": "Units", "Name": "ConvertUnits", "Step": 0, "Moveable": True}  # ダイアログのプロパティ。
	dialog, addControl = dialogCreator(ctx, smgr, controldialog)
	fixedline = {"PositionX": name["PositionX"], "PositionY": m, "Width": unit["PositionX"]+unit["Width"]-m, "Height": name["Height"], "Label": "Input only one of unit"}
	addControl("FixedLine", fixedline)
	name1, num1, unit1 = name.copy(), num.copy(), unit.copy()  # addControlに渡した辞書は変更されるのでコピーを渡す。
	name1["PositionY"] = num1["PositionY"] = unit1["PositionY"] = fixedline["PositionY"] + fixedline["Height"] + m    
	name1["Label"] = "Pixel: "  # 右寄せにすると右端文字が途中で切れるので最後はスペースにする。
	unit1["Label"] = "px"
	addControl("FixedText", name1)
	addControl("Edit", num1)  
	addControl("FixedText", unit1)
	name2, num2, unit2 = name.copy(), num.copy(), unit.copy()  # addControlに渡した辞書は変更されるのでコピーを渡す。
	name2["PositionY"] = num2["PositionY"] = unit2["PositionY"] = name1["PositionY"] + name1["Height"] + m  
	name2["Label"] = "Map AppFont: "  # 右寄せにすると右端文字が途中で切れるので最後はスペースにする。
	unit2["Label"] = "ma"
	addControl("FixedText", name2)
	addControl("Edit", num2)  
	addControl("FixedText", unit2)	
	name3, num3, unit3 = name.copy(), num.copy(), unit.copy()  # addControlに渡した辞書は変更されるのでコピーを渡す。
	name3["PositionY"] = num3["PositionY"] = unit3["PositionY"] = name2["PositionY"] + name2["Height"] + m  
	name3["Label"] = "Millimeter: "  # 右寄せにすると右端文字が途中で切れるので最後はスペースにする。
	unit3["Label"] = "1/100mm"
	addControl("FixedText", name3)
	addControl("Edit", num3)  
	addControl("FixedText", unit3)	
	button1, button2 = button.copy(), button.copy()
	button1["PositionY"] = button2["PositionY"] = name3["PositionY"] + name3["Height"] + m  
	button1["Width"] = 40
	button1["Label"] = "Con~vert"
	button2["Width"] = 30
	button2["Label"] = "~Clear"	
	button2["PositionX"] = unit["PositionX"] + unit["Width"] - button2["Width"]
	button1["PositionX"] = button2["PositionX"] - int(m/2) - button1["Width"]
	message = {"Name": "Message", "PositionX": int(m/2), "Width": button1["PositionX"]-int(m/2), "Height": 12, "NoLabel": True, "Align": 2, "VerticalAlign": MIDDLE}
	actionlistener = ActionListener(ctx, smgr)
	addControl("FixedText", message)
	addControl("Button", button1, {"setActionCommand": "convert" ,"addActionListener": actionlistener})
	addControl("Button", button2, {"setActionCommand": "clear" ,"addActionListener": actionlistener})
	dialog.getModel().setPropertyValue("Height", button1["PositionY"]+button1["Height"]+m)
	dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
	# ノンモダルダイアログにするとき。オートメーションでは動かない。
	showModelessly(ctx, smgr, docframe, dialog)  
	# モダルダイアログにする。フレームに追加するとエラーになる。
# 	dialog.execute()  
# 	dialog.dispose()	
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, ctx, smgr):
		self.args = ctx, smgr
# 	@enableRemoteDebugging
	def actionPerformed(self, actionevent):
		ctx, smgr = self.args
		cmd = actionevent.ActionCommand
		source = actionevent.Source  # ボタンコントロールが返る。
		context = source.getContext()  # コントロールダイアログが返ってくる。
		edit1 = context.getControl("Edit1")
		edit2 = context.getControl("Edit2")
		edit3 = context.getControl("Edit3")
		if cmd == "convert":
			e1, e2, e3 = edit1.getText(), edit2.getText(), edit3.getText()
			
			# 2つは空欄でないといけない。
			
			if e1.isdigit() and e2.isdigit() and e3.isdigit():
				pass
			else:
				message = context.getControl("Message")
				message.setText("")
	
		elif cmd == "clear":
			edit1.setText("")
			edit2.setText("")
			edit3.setText("")
	def disposing(self, eventobject):
		eventobject.Source.removeActionListener(self)
# def eventSource(event):  # イベントからコントロール、コントロールモデル、コントロール名を取得。
# 	control = event.Source  # イベントを駆動したコントロールを取得。
# 	controlmodel = control.getModel()  # コントロールモデルを取得。
# 	name = controlmodel.getPropertyValue("Name")  # コントロール名を取得。
# 	return control, controlmodel, name
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
