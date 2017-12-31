def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラーの取得。
	sheets = doc.getSheets()
	controller.setActiveSheet(sheets[0])
	controller.select(sheets[0]["A1"])
	