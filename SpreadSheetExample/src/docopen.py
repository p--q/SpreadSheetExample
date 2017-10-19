#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import unohelper  # オートメーションには必須(必須なのはuno)。
def macro(arg):
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	sheets = doc.getSheets()  # シートコレクション。
	sheet = sheets[0]  # 最初のシート。
	sheet[0, 0].setString(str(arg))