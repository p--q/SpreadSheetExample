#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	sheets = doc.getSheets()  # シートコレクション。
	sheet = sheets[0]  # 最初のシート。
	sheet[0, 0].clearContents(511)  # A1セルのすべてを削除。
	selection = doc.getCurrentSelection()  # 選択しているオブジェクトを取得。
	stringaddress = getRangeAddressesAsString(selection)  # セル範囲の文字列アドレスを取得。
	sheet[0, 0].setString("Selection: {}".format(stringaddress))  # A1セルに選択範囲のアドレスを出力。
	sheet[:, 0].getColumns().setPropertyValue("OptimalWidth", True)  # 列幅を最適化する。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
def getRangeAddressesAsString(rng):  # セルまたはセル範囲、セル範囲コレクションから文字列アドレスを返す。
	absolutename = rng.getPropertyValue("AbsoluteName") # セル範囲コレクションは$Sheet1.$A$4:$A$6,$Sheet1.$B$4という形式で返る。
	names = absolutename.replace("$", "").split(",")  # $を削除してセル範囲のリストにする。
	addresses = []  # 出力するアドレスを入れるリスト。
	for name in names:  # 各セル範囲について
		addresses.append(name.split(".")[-1])  # シート名を削除する。
	return ", ".join(addresses)  # コンマでつなげて出力。