#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
def docLockControllers():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doc.lockControllers()  # コントローラをロック。
def docUnLockControllers():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doc.unlockControllers()  # コントローラのロックを解除。
def docAddActionLock():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doc.addActionLock()  # ドキュメントのアクションをロック。
def docRmoveActionLock():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doc.removeActionLock()  # ドキュメントのアクションのロックを解除。
def cellAddActionLock():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	sheet["A1"].addActionLock()  # セルのアクションをロック。
def cellRmoveActionLock():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	sheet["A1"].removeActionLock()  # セルのアクションのロックを解除。
def frameAddActionLock():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	frame = controller.getFrame()  # フレームの取得。
	frame.addActionLock()  # フレームのアクションをロック。
def frameRmoveActionLock():  
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	controller = doc.getCurrentController()  # コントローラの取得。
	frame = controller.getFrame()  # フレームの取得。
	frame.removeActionLock()  # フレームのアクションのロックを解除。	
