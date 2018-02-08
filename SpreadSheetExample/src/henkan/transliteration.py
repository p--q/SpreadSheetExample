#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.i18n.TransliterationModulesNew import HALFWIDTH_FULLWIDTH, FULLWIDTH_HALFWIDTH, KATAKANA_HIRAGANA, HIRAGANA_KATAKANA, NumToCharKanjiShort_ja_JP, TextToNumLower_zh_CN, smallToLarge_ja_JP, largeToSmall_ja_JP  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.util import NumberFormat  # 定数
def halfwidth_fullwidth():  # 半角を全角に変換。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((HALFWIDTH_FULLWIDTH,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。
def fullwidth_halfwidth():  # 全角を半角に変換。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。	
def katakana_hiragana():  # かたかなをひらがなに変換。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((KATAKANA_HIRAGANA,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。	
def hiragana_katakana():  # ひらがなをカタカナに変換。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((HIRAGANA_KATAKANA,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。		
def numtocharkanjishort_ja_jp():  # アラビア数字を漢数字に変換。 
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((NumToCharKanjiShort_ja_JP,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。		
def texttonumlower_zh_cn():  # 漢数字をアラビア数字に変換。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((TextToNumLower_zh_CN,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	newtxt = transliteration.transliterate(txt, 0, len(txt), [])[0]
	if newtxt.isdigit():  # 変換結果が数字のみの時。
		selection[0, 0].setValue(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。setString()で代入すると書式を数値に変更しても文字列のままになる。
		numberformats = doc.getNumberFormats()
		formatkey = numberformats.getStandardFormat(NumberFormat.NUMBER, Locale())  # 数値の標準書式のキーを取得。
		selection[0, 0].setPropertyValue("NumberFormat", formatkey)  # セルの書式を設定。
def smalltolarge_ja_jp():  # ゃゅょっぁぃぅぇぉ、などを大きくする。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((smallToLarge_ja_JP,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。		
def largetosmall_ja_jp():  # やゆよ、などを小さくするはずだが動かない。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration 
	selection = doc.getCurrentSelection()  # 選択範囲を取得。
	txt = selection[0, 0].getString()  # 選択範囲の左上端のセルの内容を文字列として取得する。
	transliteration.loadModuleNew((largeToSmall_ja_JP,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
	selection[0, 0].setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 変換して選択セルに代入。					
