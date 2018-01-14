import math
import time
from com.sun.star.sheet import CellFlags as cf # 定数
def macro():
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	x = -3.14
	y = 3.14
	n = 30000
	d = (y-x)/n	
	
	controller = doc.getCurrentController()  # コントローラーの取得。
	sheet = controller.getActiveSheet()
	
# 	doc.addActionLock()
# 	doc.lockControllers()
	sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。
	start = time.perf_counter()
	for i in range(n):
		k = x + d*i
		sheet[i, 0].setValue(k)
		sheet[i, 1].setValue(math.sin(k))
		sheet[0, 2].setValue(i+1)
	end = time.perf_counter()
	sheet[1, 2].setString("Finished: {}s".format(end-start))
# 	doc.removeActionLock()
# 	doc.unlockControllers()
	