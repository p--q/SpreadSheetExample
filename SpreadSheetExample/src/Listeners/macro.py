def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラーの取得。
	sheets = doc.getSheets()
	
	doc.lockControllers()
	
	doc.addActionLock()
	
	sheet = sheets[0]
	controller.setActiveSheet(sheet)
	
# 	sheet = controller.getActiveSheet()


	sheet["A1"].setString("text")
# 	sheet["A1"].setPropertyValues(("IsCellBackgroundTransparent", "CellBackColor"), (False, 0xfff200))
	
# 	controller.select(sheet["A1"])
# 	controller.setActiveSheet(sheets[1])
	
# 	doc.removeActionLock()
	
# 	doc.unlockControllers()