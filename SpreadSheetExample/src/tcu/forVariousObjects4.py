#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from wsgiref.simple_server import make_server
import webbrowser
import time
import re
from xml.etree import ElementTree as ET
from com.sun.star.beans import PropertyValue  # Struct
def macro(documentevent=None):  # 引数は文書のイベント駆動用。
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。


	desktop = XSCRIPTCONTEXT.getDesktop() 
	doc = XSCRIPTCONTEXT.getDocument()
	controller = doc.getCurrentController()  # コントローラの取得。
	frame = controller.getFrame()  # フレームを取得。
	containerwindow = frame.getContainerWindow()
	componentwindow = frame.getComponentWindow()
	toolkit = containerwindow.getToolkit()
	

	
	obj = ("Desktop", desktop),\
		("Frame", frame),\
		("Container Window", containerwindow),\
		("Component Window", componentwindow),\
		("Toolkit", toolkit) # ツリー名とオブジェクトのタプルのタプル。
	objs = ("Container Window vs. Component Window", (containerwindow, componentwindow)),\
			("Desktop vs. Frame", (desktop, frame))
	createTrees(obj, objs)
def createTrees(obj, objs):
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	start = time.perf_counter()  # 開始時刻を記録。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 
	# タブノードとタブボディノードの作成。
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  
	# Orange10, Red10, Pink10, Magenta10, Violet10, Blue10, SkyBlue10, Cyan10, Turquoise10, Green10, YellowGreen10	
	colors = 0xFFCC99, 0xFFCCCC, 0xFF99CC, 0xFFCCFF, 0xCC99FF, 0xCCCCFF, 0x99CCFF, 0xCCFFFF, 0x99FFCC, 0xCCFFCC, 0xCCFF99 
	rgbcolors = [colorToRGB(i) for i in colors]  # RGBタプルのリストにする。
	nodepairs = []  # タブノードとタブボディノードのペアを入れるリスト。
	[nodepairs.append(createNodes(n, tcu.wtreelines(j), "tcutree")) for n, j in obj]  # wtreeのタブノードとタブボティノードのペアを取得。
	[nodepairs.append(createNodes(n, tcu.wcomparelines(*j), "tcucompare")) for n, j in objs]  # wcompareのタブノードとタブボティノードのペアを取得。
	n = len(colors)  # 色数の取得。色は順繰りに使い回す。
	style = "\n".join(["#tcutab div:nth-child({0}), #tcutabbody div:nth-child({0}) {{background-color:rgb({1}, {2}, {3});box-shadow: 0 5px 5px rgba({1}, {2}, {3}, .5);}}"\
					.format(i, *rgbcolors[i%n]) for i in range(1, len(nodepairs)+1)])  # タブノードとタブボディノードの色の指定。
	stylenode = Elem("style", text=style)  # 動的なCSS。	
	tabnodes = Elem("div", {"id": "tcutab", "style": "padding:0.5em;display:flex;flex-flow:row wrap;"})  # flexコンテナ。
	tabbodynodes = Elem("div", {"id": "tcutabbody"})
	for tabnode, tabbodynode in nodepairs:  # 作成したタブノードとタブボティノードのペアについて。
		tabnodes.append(tabnode)  # タブノードを追加。
		tabbodynodes.append(tabbodynode)  # タブボディノードを追加。
	# ヘッダーノードの作成。
	configurationprovider = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
	node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
	configurationaccess = configurationprovider.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
	libreversion = configurationaccess.getPropertyValues(('ooName', 'ooSetupVersionAboutBox'))  # LibreOfficeの名前とバージョンをタプルで返す。
	titlenode = Elem("div", {"class": "tcutitle", "style": "display:flex;align-items:center;"}, sub=Elem("div", {"style": "font-family:Tahoma, Arial, sans-serif;font-size:150%;font-weight:bold;padding:10px;"}, text="{} {}".format(*libreversion)))
	inputnode = Elem("input", {"type": "serach", "name": "q", "placeholder": "Search the tree...", "aria-label": "Search through tree content", "accesskey": "s", "required": "", "style": "width:156px;position:absolute;top:3px;left:12px;border:1px solid #FFF;"})  # requiredは空文字を渡しても有効になる。
	image1 = "data:image/gif;base64,R0lGODlhEgASALMAAIeHh////9fX18XFxbS0tO7u7qSkpN7e3pKSkvj4+MzMzL29vebm5q2trY6OjpmZmSH5BAA\
HAP8ALAAAAAASABIAAARhMMiZpr2yqMEPvgdBCMdCKNVXjGlQLMoXKAt2NAX2eley8JZXDqMQ6AYMTGIAtHAwDNwnapysECwbgXlQbAcOByF5KQgGMM8SATAIWhR4QmEAPAZwmWsBaOiVTX+Cg4QYEQA7"  # これはなぜかタブをつけても画像が有効になる。			
	imagenode = Elem("img", {"src": image1, "alt": "", "style": "position:absolute;top:3px;left:174px;cursor:pointer;"})
	image2 = "data:image/gif;base64,R0lGODlhyAAYAKIAANbW1v///97e3vf39+bm5u/v7wAAAAAAACH5BAAHAP8ALAAAAADIABgAAAPZGLpaAjDKSau9OOvNu//gJRQDYzKEQJBD675wLM90bd94ru98Pzu\
q0qkRHBqPyKRyyWw6n9Boc5AqDB1CqXbL7Xq/3sHIhAWbz+i0eim2LsbruHxOd4qFBEJ9z+/P8woAWX6EhYZQggV6h4yNjiYjcI+TlH1Ag5WZmmgDEJiboKFanqKlpk+kp6qrJ52SrLCnDoqxtaYrnZ+2u465A\
YC8wY/AAQBuwsiEdwtlyc51bWQCus/VX53HDAUABNTW309UrycE3CTg6FOK3N7aAg8h8fLz9PX29SO6CQA7"	
	formnode = Elem("div", {"id": "tcuform", "style": "width:200px;display:block;height:24px;position:relative;margin:0 0 0 auto;background: url({}) left top no-repeat;".format(image2)}, subs=(inputnode, imagenode))  # フォームノード。
	resetbutton = Elem("button", {"type": "button", "style": "margin: 5px 0;border-style:none;padding:5px;border-radius:5px;font-weight:bold;outline:none;"}, text="Reset")
	subflexnode1 = Elem("div", {"style":"display:flex;flex-direction:column;"}, subs=(formnode, resetbutton))
	labelstyle = "display:block;position:relative;color:#2A3D61;font-size:14px;padding-left:1.2em;cursor:pointer;"
	inputstyle = "position: absolute;margin: auto;left:0;cursor:pointer;outline:none;"
	checkboxnode1 = Elem("label", {"style": labelstyle}, sub=Elem("input", {"type": "checkbox", "style": inputstyle}, text="Do not display the compare mode"))
	checkboxnode2 = Elem("label", {"style": labelstyle}, sub=Elem("input", {"type": "checkbox", "checked": "", "style": inputstyle}, text="Clear search terms"))
	subflexnode2 = Elem("div", {"style":"display:flex;flex-direction:column;justify-content:space-around;padding-left:5px;"}, subs=(checkboxnode1, checkboxnode2))
	headernode = Elem("div", {"id": "tcuheader", "style": "display:flex;justify-content:space-between;border-bottom:1px solid #C4CFE5;padding:0.5em 0 0.5em 0.5em"}, subs=(titlenode, Elem("div", {"style":"display:flex"}, subs=(subflexnode1, subflexnode2))))  # ヘッダーノード。flexコンテナ。space-betweenで配置。
	# フッタノードの作成。
	extensionmanager = ctx.getByName('/singletons/com.sun.star.deployment.ExtensionManager')
	extension = extensionmanager.getDeployedExtension("user", "pq.Tcu", "TCU.oxt", None)  # TCUの名前とバージョン番号の取得のため。	
	footernode = Elem("div", {"id": "tcufooter", "style": "text-align:right;margin-top:10px;color:#2A3D61;font-size:14px;line-height:22px;"})  # フッタノード。
	footernode.append(Elem("div", text="Generated by "))
	footernode[-1].append(Elem("a", {"href": "https://github.com/p--q/TCU"}, text=extension.getDisplayName(), tail=" {}".format(extension.getVersion())))
	footernode.append(Elem("div", text="Elapsed Time: {}s".format(time.perf_counter()-start)))	# 実行時間の出力。
	# 作成したノードをボディノードに追加する。
	root, scriptnode = createRoot()  # scriptノードは最後に挿入したいので別に取得する。
	bodynode = root[-1]  #  bodyノードを取得。	
	bodynode.append(stylenode)
	bodynode.append(headernode)
	bodynode.append(tabnodes)
	bodynode.append(tabbodynodes)
	bodynode.append(footernode)
	bodynode.append(scriptnode)
	
	
# 	dic = createCSS(root)
	
# 	dic = createCSS(headernode)
# 	print("\n".join(dic.keys()))
	
	dic = createXPath(root)
	print("\n".join(dic.keys()))
	
	
# 	toBrowser(root)

#####################XPath##############################	
def createXPath(root):  # CSSパスの辞書を返す。
	dic = {}  # キー:XPath、値:style属性の値、の辞書。
	style_xpath = './/*[@style]'  # sytleのあるノードを取得するXPath。
	style_nodes = root.findall(style_xpath)  # styleのあるノードをすべて取得。
	getParentNode = createParentGetterr(root)  # Elementオブジェクトから親Elementオブジェクトを取得する関数を取得。
	while style_nodes:  # styleのあるノードがある間実行。
		n = style_nodes.pop()  # スタイルのあるノードを取得。
		xpath = getElementXPath(getParentNode, n)  # ノードのCSSパスを取得。
		dic[xpath] = n.get("style")  # 同じCSSパスがでてくるのでその処理が必要。
	return dic
def getElementXPath(getParentNode, n):
	idprop = n.get("id")
	if idprop:
		return '//*[@id="{}"]'.format(idprop)
	else:
		return getElementTreeXPath(getParentNode, n)
def getElementTreeXPath(getParentNode, n):
	paths = []
	while n is not None:
		p = getParentNode(n)
		if p is not None:
			children = list(p)
			index = children.index(n) + 1
			pathindex = "[{}]".format(index) if index<len(children) else ""
			paths.append("".join((n.tag, pathindex)))
		n = p
	return "/{}".format("/".join(reversed(paths))) if paths else None
	
	
	
	
	
	



#####################XPath終わり##############################	

#####################CSS パス##############################	
def createCSS(root):  # CSSパスの辞書を返す。
# 	dic = {}  # キー:CSSパス、値:style属性の値、の辞書。
	style_xpath = './/*[@style]'  # sytleのあるノードを取得するXPath。
	style_nodes = root.findall(style_xpath)  # styleのあるノードをすべて取得。
	getParentNode = createParentGetterr(root)  # Elementオブジェクトから親Elementオブジェクトを取得する関数を取得。
	while style_nodes:  # styleのあるノードがある間実行。
		n = style_nodes.pop()  # スタイルのあるノードを取得。
		csspath = getElementCSSPath(getParentNode, n)  # ノードのCSSパスを取得。
# 		dic[csspath] = n.get("style")  # 同じCSSパスがでてくるのでその処理が必要。
# 	return dic
def getElementCSSPath(getParentNode, n):  # getParentNodeに渡したルートからのCSSパスを取得。
	paths = []
	while n is not None:  # Elementオブジェクトは子要素ない時はFalseになる(リストと同じ)なのでNoneで判断しないといけない。
		paths.append(getElementCSSSelector(n))
		n = getParentNode(n)
	return " ".join(reversed(paths))
def getElementCSSSelector(n):  # ノードnのみのセレクタを返す。
	label = n.tag.split(":")[-1].lower()  # localName、つまりタブ名を小文字で取得。コロンがあればその前は無視する。
	idprop = n.get("id")  # id属性があればそれを取得。
	if idprop:
		label = "".join((label, "#{}".format(idprop)))
	classes = n.get("class")
	if classes:
		label = "".join((label, *[".{}".format(i) for i in classes.split(" ")]))
	return label	
def createParentGetterr(root):  # root: ElementオブジェクトかElementTree。
	def getParentNode(n):  # n: Elementオブジェクト。親ノードを返す。
		idprop = n.get("id")  # id属性があればそれを取得。
		if idprop:  # id属性のあるノードのとき。xpathで必ず一つに絞れる。
			xpath = './/*[@id="{}"]/..'.format(idprop)  # 親ノードのxpathを作成。
			return root.find(xpath)  
		# id属性以外の時は複数のノードが選択される可能性があるのでfindall() or iterfind()を使う。
		tag = n.tag
		classprop = n.get("class")
		styleprop = n.get("style")
		if classprop:  # class属性があるとき。
			xpath = './/{}[@class="{}"]/..'.format(tag, classprop)
		elif styleprop:  # style属性があるとき。
			xpath = './/{}[@style="{}"]/..'.format(tag, styleprop)  # 親ノードのxpathを作成。各ノードの親ノードは一つしかないがstyle属性だけでは一つノードに絞り込めないので複数ノードが返ってくる可能性がある。
		else:  # 特定できる属性がないときはタグ名のみで検索するしかない。
			xpath = './/{}/..'.format(tag)
		for p in root.iterfind(xpath):  # 各親ノード候補に対して。
			if n in list(p):  # 子ノードが一致したのが親ノード。
				return p
		return None  # 親ノードが見つからなければNoneを返す。
	return getParentNode
#####################CSS パス 終わり##############################	


def colorToRGB(color):  # 色をRGBのタプルで返す。	16進数で渡しても10進数で計算されている。
	red = int(color/0x10000)  # 0x10000がいくつあるかがred。
	redmodulo = color % 0x10000  # redの要素を削除。
	green = int(redmodulo/0x100)  # 0x100がいくつあるかがgreen。
	blue = redmodulo % 0x100  # redとgreenの要素を削除した残りがblue。
	return red, green, blue
def createNodes(name, lines, linestype):  # linestype: wtreelines()かwcomparelines()かの区別、name: タブの表示名(ユニークでないといけない)、lines: ツリーのhtmlの行のリスト。
	tabbodyid = name.replace(" ", "").replace(".", "")  # nameからタブボディのidを作成する。空白を削除する。英数字、'_'、'-'、'.' 以外の文字はHTML4では不可。CSSセレクタ作成時に困るのでドットも削除する。	
	tabnodestyle = "display:inline-block;padding: 0.8em 1em;margin:0.2em;border-radius:1.6em;font-size:18px;font-family:'Lato',sans-serif;font-weight:700;text-align:center;cursor:pointer;color:#2A3D61;"
	tabnode = Elem("div", {"class": linestype, "style": tabnodestyle}, text=name)  # ユニークな名前のタブノードを作成。
	tabbodynode = Elem("div", {"id": tabbodyid, "class": linestype, "style": "padding:1em;border-radius:1.6em;color:#2A3D61;display:none"})  # タブボディノードにはタブの表示名から作成したclassnameをつける。
	tabbodynode.append(Elem("p", {"style": "font-size:150%;font-weight:bold;"}, text=name))  # タブボディに表示するタイトル。
	html = "<br/>".join(lines).replace(" ", chr(0x00A0))  # 半角スペースをノーブレークスペースに置換する。
	html = re.sub(r'(?<!\u00A0)\u00A0(?!\u00A0)', " ", html)  # タグ内にノーブレークスペースはエラーになるので連続しないノーブレークスペースを半角スペースに戻す。
	xml = "<code style='white-space:nowrap;'>{}</code>".format(html)  # ツリーのhtmlを完成させる。
	tabbodynode.append(ET.XML(xml))  # タブボディノードにツリーを部分木にして追加する。
	return tabnode, tabbodynode  # タブノードとタブボディノードのタプルを返す。
def createRoot():  # ルートノードを返す。
	rt = Elem("html")
	rt.append(Elem("head"))
	rt[-1].append(Elem("title", text="TCU - Tree Command for UNO"))
	rt[-1].append(Elem("meta", {"meta": "UTF-8"}))
	rt.append(Elem("body"))
	# CSSの作成。
	style = """\
button:hover {
	text-decoration: underline;  /* 下線を引く */
	background-color: #24d;
	color: #fff;
	cursor: pointer;
}
/* Firefox */
button::-moz-focus-inner {
  border: 0;  /* 選択時の点線を消す */
}
#tcutab div:hover {  /* タブにマウスポインタが乗ったとき */
	text-decoration: underline;  /* 下線を引く */
}"""
	rt[-1].append(Elem("style", text=style))
	# スクリプトノードの作成。
	script = """//  TCUモジュール
var pq_TCU = pq_TCU || function() {
	var tcu = {  // グローバルに出すオブジェクト。
		all: function() { // ここから開始する。
			g.tab.addEventListener('mousedown', eh.mouseDownTab, false); // タグノードにmousedownイベントリスナーを追加。useCaptureオプションをfalseに指定。
			g.tabbody.addEventListener('mousedown', eh.mouseDownTabBody, false); // タブボディノードにmousedownイベントリスナーを追加。useCaptureオプションをfalseに指定。
			g.form.getElementsByTagName('img')[0].addEventListener('mousedown', eh.mouseDownImg, false); // 検索ノード内のimgタグにmousedownイベントリスナーを追加。useCaptureオプションをfalseに指定。
			g.form.getElementsByTagName('input')[0].addEventListener('keydown', eh.keydownInput, false); // 検索ノード内のinputタグにkeydownイベントリスナーを追加。useCaptureオプションをfalseに指定。	
			g.tabbody.children[0].style.display = "inline-block";  // ページを開いた時は最初のタブのタブボディを表示する。
		}
	}  // end of tcu
	var g = { // モジュール内の"グローバル"変数。
		tab: document.getElementById('tcutab'),  // タブノード。
		tabbody: document.getElementById('tcutabbody'),  // タブボディノード。
		form:  document.getElementById('tcuform'),  // 検索ノード。
		tabclasses: ["tcutree", "tcucompare"]  // タブクラス名。
	};  // end of g
	var eh = {  // イベントハンドラオブジェクト。
		mouseDownTab: function(e) {  // タブノードをクリックした時。
			var target = e.target; // イベントを発生した要素を取得。タブのDOMが返ってくる。
			var tabclass = target.className  // ターゲットのクラスを取得。
			if (g.tabclasses.indexOf(tabclass)!=-1) {  // タブがクリックされたときのみ実行。そうしないとボタンを以外をクリックしても反応する。
				var tabname = target.textContent.replace(/\s+/g, "")  // タブ名を空白を除いて取得。
				var tabbodys = g.tabbody.children  // HTMLCollection(≠配列)が返る。childNodesだとTextNodeまでも返ってくる。
				for (var i=0;i<tabbodys.length;i++) {  // childrenではTextNodeを除外して取得できるが配列ではないのでforEachは使えないらしい。タブノードのHTMLCollection。
					if (tabbodys[i].id==tabname) {  // タブ名が一致する時。
						tabbodys[i].style.display = "inline-block";  // タブボディを表示する
					} else {  // タブ名が一致しない時。
						tabbodys[i].style.display = "none";  // 非表示にする。
					}
				}
			}
		},
		mouseDownTabBody: function(e) {  // タブボディノードをクリックした時。
			var target = e.target; // イベントを発生した要素を取得。
			if (target.tagName.toLowerCase()=="a") {  // aタグの時。タグ名は大文字で返ってくるらしい。
				if (target.href.startsWith("file")) {  // ローカルファイルをアンカーしている時。
					if (!target.baseURI.startsWith("file")) {  // ローカルファイルのページでない時。
						window.alert("You can not move to the local reference page for security reasons.\\n You have to save this file to the disk and reopen it.");  // \\nの\\はエスケープが必要(Pythonの文字列のため)。
					}
				}
			} 
		},
		mouseDownImg: function(e) {  // 検索ノードをクリックした時。
			var target = e.target; // イベントを発生した要素を取得。
			var input = g.form.getElementsByTagName('input')[0]  // inputタグを取得。
			eh._searchQuery(input.value)
		},
		keydownInput: function(e) { // 検索ボックスにキー入力された時。
			if (e.key=="Enter") {  // Enterキーが押された時。
				var target = e.target; // イベントを発生した要素を取得。
				eh._searchQuery(target.value)
				e.preventDefault();  // これ以上イベントを発生させない。
			}
		},
		_searchQuery: function(q) {
			g.tabbody.getElementsByTagName('code')
		
		
		
		
		}
	};  // end of eh
	return tcu; // グローバルスコープにオブジェクトを出す。
}();
pq_TCU.all();"""
	scriptnode = Elem("script", text=script)
	return rt, scriptnode  # rt[-1]がボディノード。
class Wsgi:
	def __init__(self, html):
		self.resp = html  # どこかでutf-8にエンコードしないといけない。今回は受け取るときにすでにエンコード済。
	def app(self, environ, start_response):  # WSGIアプリ。引数はWSGIサーバから渡されるデフォルト引数。
		start_response('200 OK', [ ('Content-type','text/html; charset=utf-8')])  # charset=utf-8'がないと文字化けする時がある
		yield self.resp  # デフォルトエンコードはutf-8。
	def wsgiServer(self): 
		host, port = "localhost", 8080  # サーバが受け付けるポート番号を設定。
		httpd = make_server(host, port, self.app)  # appへの接続を受け付けるWSGIサーバを生成。
		url = "http://localhost:{}".format(port)  # 出力先のurlを取得。
		webbrowser.open_new_tab(url)   # デフォルトブラウザでurlを開く。
# 		httpd.serve_forever()  # JavaScriptのデバッグ時はサーバーを立ち上がておく。
		httpd.handle_request()  # リクエストを1回だけ受け付けたらサーバを終了させる。ローカルファイルはセキュリティの制限で開けない。
def toBrowser(root):  # ブラウザにルートとなるElementオブジェクトを渡す。
	html = ET.tostring(root, encoding="utf-8",  method="html")  # utf-8にエンコードしてhtmlにする。utf-8ではなくunicodeにすると文字列になる。method="html"にしないと<script>内がhtmlエンティティになってしまう。
	server = Wsgi(html)  # エンコード済のhtmlを渡す。
	server.wsgiServer()  # htmlをデフォルトブラウザに出力。
class Elem(ET.Element):  # xml.etree.ElementTree.Element派生クラス。
	def __init__(self, tag, attrib={},  **kwargs):  # ET.Elementのアトリビュートのtextとtailはkwargsで渡す。
		txt = kwargs.pop("text", None)  # ET.Elementのtextアトリビュートを取得。
		tail = kwargs.pop("tail", None)  # ET.Elementのtailアトリビュートを取得。
		sub = kwargs.pop("sub", None)  # サブノードにするET.Elementを取得。
		subs = kwargs.pop("subs", None)  # サブノードにするET.Elementのタプルを取得。
		super().__init__(tag, attrib, **kwargs)  #  ET.Elementをインスタンス化=self。
		if txt is not None:
			self.text = txt
		if tail is not None:
			self.tail = tail
		if sub is not None:  # ET.Elementが入っていてもsubだけだとFalseになる。
			self.append(sub)
		if subs is not None:
			self.extend(subs)
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
