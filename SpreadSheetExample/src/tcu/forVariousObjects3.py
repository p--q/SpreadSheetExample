#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from wsgiref.simple_server import make_server
import webbrowser
import time
import re
from xml.etree import ElementTree as ET
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.ucb import XCommandEnvironment
from com.sun.star.rendering import RGBColor  # Struct
def macro():	
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。


	desktop = XSCRIPTCONTEXT.getDesktop() 
	doc = XSCRIPTCONTEXT.getDocument()
	controller = doc.getCurrentController()  # コントローラの取得。
	frame = controller.getFrame()  # フレームを取得。
	containerwindow = frame.getContainerWindow()
	componentwindow = frame.getComponentWindow()
	toolkit = containerwindow.getToolkit()
	
	dic_obj = {\
			"Desktop": desktop,\
			"Frame": frame,\
			"Container Window": containerwindow,\
			"Component Window": componentwindow,\
			"Toolkit": toolkit\
			}  # ツリーを出力するオブジェクトの辞書。
	dic_objs = {\
			"Container Window vs. Component Window": (containerwindow, componentwindow),\
			"Desktop vs. Frame": (desktop, frame)\
			}  # 比較ツリーを出力するオブジェクトの辞書。
	createTrees(ctx, dic_obj, dic_objs)
def createTrees(ctx, dic_obj, dic_objs):
	start = time.perf_counter()  # 開始時刻を記録。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  
	configurationprovider = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
	node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
	configurationaccess = configurationprovider.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
	libreversion = configurationaccess.getPropertyValues(('ooName', 'ooSetupVersionAboutBox'))  # LibreOfficeの名前とバージョンをタプルで返す。
	extensionmanager = ctx.getByName('/singletons/com.sun.star.deployment.ExtensionManager')
	extension = extensionmanager.getDeployedExtension("user", "pq.Tcu", "TCU.oxt", None)  # TCUの名前とバージョン番号の取得のため。
	colors = 0xFFCC99, 0xFFCCCC, 0xFF99CC, 0xFFCCFF, 0xCC99FF, 0xCCCCFF, 0x99CCFF, 0xCCFFFF, 0x99FFCC, 0xCCFFCC , 0xCCFF99  # Orange10, Red10, Pink10, Magenta10, Violet10, Blue10, SkyBlue10, Cyan10, Turquoise10, Green10, YellowGreen10	
	nodepairs = []
	[nodepairs.append(createNodes(key, tcu.wtreelines(val))) for key, val in dic_obj.items()]
	[nodepairs.append(createNodes(key, tcu.wcomparelines(*val))) for key, val in dic_objs.items()]
	n = len(colors)
	style = "\n".join(["#tabcontrol a:nth-child({0}), #tabbody div:nth-child({0}) {{background-color:rgb({1}, {2}, {3});box-shadow: 0 5px 20px rgba({1}, {2}, {3}, .5);}}"\
					.format(i, *colorToRGB(colors[i%n])) for i in range(1, len(nodepairs)+1)])
	
	
	stylenode = Elem("style", text=style)				
	tabnodes = Elem("p", {"id": "tabcontrol"})
	tabbodynodes = Elem("div", {"id": "tabbody"})
	for tabnode, tabbodynode in nodepairs:
		tabnodes.append(tabnode)
		tabbodynodes.append(tabbodynode)
	root, scriptnode = createRoot()  # scriptノードは最後に挿入したいので別に取得する。
	bodynode = root[1]  #  bodyノードを取得。
	end = time.perf_counter()  # 終了時刻を記録。
	title = "{} {}".format(*libreversion)	
	titlenode = Elem("div", {"style": "display:flex;flex-flow:row wrap;justify-content:space-between"})  # flexコンテナ。3つのアイテムをspace-betweenで配置。
	titlenode.append(Elem("div", {"style": "font-family:Tahoma, Arial, sans-serif;font-size:150%;font-weight:bold;"}, text=title))
	titlenode.append(Elem("div"))  # 真ん中のflexアイテム。
	image = "data:image/gif;base64,R0lGODlhyAAYAKIAANbW1v///97e3vf39+bm5u/v7wAAAAAAACH5BAAHAP8ALAAAAADIABgAAAPZGLpaAjDKSau9OOvNu//gJRQDYzKEQJBD675wLM90bd94ru98Pzu\
q0qkRHBqPyKRyyWw6n9Boc5AqDB1CqXbL7Xq/3sHIhAWbz+i0eim2LsbruHxOd4qFBEJ9z+/P8woAWX6EhYZQggV6h4yNjiYjcI+TlH1Ag5WZmmgDEJiboKFanqKlpk+kp6qrJ52SrLCnDoqxtaYrnZ+2u465A\
YC8wY/AAQBuwsiEdwtlyc51bWQCus/VX53HDAUABNTW309UrycE3CTg6FOK3N7aAg8h8fLz9PX29SO6CQA7"	
	formnode = Elem("div", {"id": "form", "style": "background: url({}) left top no-repeat;".format(image)})
	formnode.append(Elem("input", {"id": "query", "type": "serach", "name": "q", "placeholder": "Search the tree...", "aria-label": "Search through tree content", "accesskey": "s", "required": ""}))  # requiredは空文字で有効になる。 
	image = "data:image/gif;base64,R0lGODlhEgASALMAAIeHh////9fX18XFxbS0tO7u7qSkpN7e3pKSkvj4+MzMzL29vebm5q2trY6OjpmZmSH5BAA\
	HAP8ALAAAAAASABIAAARhMMiZpr2yqMEPvgdBCMdCKNVXjGlQLMoXKAt2NAX2eley8JZXDqMQ6AYMTGIAtHAwDNwnapysECwbgXlQbAcOByF5KQgGMM8SATAIWhR4QmEAPAZwmWsBaOiVTX+Cg4QYEQA7"  # これはなぜかタブをつけても画像が有効になる。
	formnode.append(Elem("img", {"id": "querybutton", "src": image, "alt": "", "style": "cursor:pointer;"})) 	
	titlenode.append(formnode)
	bodynode.append(titlenode)



	bodynode.append(stylenode)
	bodynode.append(tabnodes)
	bodynode.append(tabbodynodes)
	txt = "Generated by "
	tail = " {}.".format(extension.getVersion())
	bodynode.append(Elem("div", {"style": "text-align:right;margin-top:10px;"}, text=txt))
	bodynode[-1].append(Elem("a", {"href": "https://github.com/p--q/TCU"}, text=extension.getDisplayName(), tail=tail))
	txt = "Elapsed Time: {}s".format(end-start)
	bodynode.append(Elem("div", {"style": "text-align:right;"}, text=txt))	
	bodynode.append(scriptnode)
	toBrowser(root)

def colorToRGB(color):	
	red = int(color/0x10000)
	redmodulo = color % 0x10000
	green = int(redmodulo/0x100)
	blue = redmodulo % 0x100
	return red, green, blue
def createNodes(name, lines):
	i = name.replace(" ", "").replace(".", "")
	tabnode = Elem("a", {"href": "#{}".format(i)}, text=name)
	tabbodynode = Elem("div", {"id": i})
	tabbodynode.append(Elem("p", text=name))
	html = "<br/>".join(lines).replace(" ", chr(0x00A0))  # 半角スペースをノーブレークスペースに置換する。
	html = re.sub(r'(?<!\u00A0)\u00A0(?!\u00A0)', " ", html)  # タグ内にノーブレークスペースはエラーになるので連続しないノーブレークスペースを半角スペースに戻す。
	xml = "<tt style='white-space: nowrap;'>{}</tt>".format(html)
	tabbodynode.append(ET.XML(xml))
	return tabnode, tabbodynode
def createRoot():
	rt = Elem("html")
	rt.append(Elem("head"))
	rt[0].append(Elem("title", text="TCU - Tree Command for UNO"))
	rt[0].append(Elem("meta", {"meta": "UTF-8"}))
	rt.append(Elem("body"))
	style = """\
/* タブ */
#tabcontrol a {
	display: inline-block;
	padding: 1em 3.2em;
	border-radius: 1.6em;
	color: #fff;
	font-size: 18px;
	font-family: 'Lato', sans-serif;
	font-weight: 700;
	text-align: center;
	text-decoration: none;
}
/* タブにマウスポインタが載った際 */
#tabcontrol a:hover {
	text-decoration: underline;
}
/* タブの中身 */
#tabbody div {
	border: 1px solid black; /* 枠線：黒色の実線を1pxの太さで引く */
	margin-top: -1px;		/* 上側にあるタブと1pxだけ重ねるために「-1px」を指定 */
	padding: 1em;			/* 内側の余白量 */
	background-color: white; /* 背景色：白色 */
	position: relative;	  /* z-indexを調整するために必要 */
	z-index: 0;			  /* 重なり順序を「最も背面」にするため */
}
#query  {  
	width: 156px;  
	position: absolute;  
	top: 3px;  
	left: 12px;  
	border: 1px solid #FFF;  
}  
#querybutton {  
	position: absolute;  
	top: 3px;  
	left: 174px;  
}
#form { 
	width: 200px;  
	display: block;  
	height: 24px;  
	position: relative;  
margin : 0 0 0 auto;
}"""
	rt[1].append(Elem("style", text=style))
	script = """\
var tabs = document.getElementById('tabcontrol').getElementsByTagName('a');
var pages = document.getElementById('tabbody').getElementsByTagName('div');	
function changeTab() {
	var targetid = this.href.substring(this.href.indexOf('#')+1, this.href.length);  // href属性値から対象のid名を抜き出す
	for (var i=0;i<pages.length; i++) {
		if (pages[i].id!=targetid) {
			pages[i].style.display = "none";
		} else { 
			pages[i].style.display = "inline-block";  // 指定のタブページだけを表示する
		}
	}
	for (var i=0;i<tabs.length; i++) {
		tabs[i].style.zIndex = "0";
	}
	this.style.zIndex = "10";  // クリックされたタブを前面に表示する
	return false;  // ページ遷移しないようにfalseを返す
}
for (var i=0;i<tabs.length; i++) {
	tabs[i].onclick = changeTab;  // すべてのタブに対して、クリック時にchangeTab関数が実行されるよう指定する
}
tabs[0].onclick();  // 最初は先頭のタブを選択	"""
	scriptnode = Elem("script", text=script)
	return rt, scriptnode
class Wsgi:
	def __init__(self, html):
		self.resp = html
	def app(self, environ, start_response):  # WSGIアプリ。引数はWSGIサーバから渡されるデフォルト引数。
		start_response('200 OK', [ ('Content-type','text/html; charset=utf-8')])  # charset=utf-8'がないと文字化けする時がある
		yield self.resp  # デフォルトエンコードはutf-8。
	def wsgiServer(self): 
		host, port = "localhost", 8080  # サーバが受け付けるポート番号を設定。
		httpd = make_server(host, port, self.app)  # appへの接続を受け付けるWSGIサーバを生成。
		url = "http://localhost:{}".format(port)  # 出力先のurlを取得。
		webbrowser.open_new_tab(url)   # デフォルトブラウザでurlを開く。
		httpd.handle_request()  # リクエストを1回だけ受け付けたらサーバを終了させる。ローカルファイルはセキュリティの制限で開けない。
def toBrowser(root):
	html = ET.tostring(root, encoding="utf-8",  method="html")  # utf-8にエンコードする。utf-8ではなくunicodeにすると文字列になる。method="html"にしないと<script>内がhtmlエンティティになってしまう。
	server = Wsgi(html)
	server.wsgiServer()	
class Elem(ET.Element):  # キーワード引数text, tailでテキストノードを付加するxml.etree.ElementTree.Element派生クラス。
	def __init__(self, tag, attrib={},  **kwargs):  # textキーワードは文字列のみしか受け取らない。  
		txt = kwargs.pop("text", None)
		tail = kwargs.pop("tail", None)
		super().__init__(tag, attrib, **kwargs)
		if txt:
			self.text = txt
		if tail:
			self.tail = tail
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
# 		doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
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
