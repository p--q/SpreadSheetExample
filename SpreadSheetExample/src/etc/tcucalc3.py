from com.sun.star.embed import ElementModes  # 定数
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
	doc = XSCRIPTCONTEXT.getDocument()
	documentstorage = doc.getDocumentStorage()  # コンポーネントからストレージを取得。
	thumbnailsstorage = documentstorage.openStorageElement("Thumbnails", ElementModes.READ)  # ドキュメント内のThumbnailsストレージを取得。
	tcu.wtree(thumbnailsstorage)