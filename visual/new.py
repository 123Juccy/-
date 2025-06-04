import cv2
img = cv2.imread("C:/Users/OSS360211/Desktop/digepic/pic/lena.jpg")
templ = cv2.imread("C:/Users/OSS360211/Desktop/digepic/pic/lena1.jpg")
#获取模板匹配的高，宽和通道数
height,width,c = templ.shape
#按照标准平方差匹配
result = cv2.matchTemplate(img,templ,cv2.TM_SQDIFF_NORMED)
#获取匹配的最小值，最大值，最小值坐标，最大值坐标
minvalue,maxvalue,minLoc,maxLoc = cv2.minMaxLoc(result)
resultPoint1 = minLoc
resultPoint2 = (resultPoint1[0]+width,resultPoint1[1]+height)
#在最佳匹配区域绘制红色框，线宽为2
cv2.rectangle(img,resultPoint1,resultPoint2,(0,255,0),2)
cv2.imshow("templ",templ)
cv2.imshow("img",img)
cv2.waitKey(0)
cv2.destroyAllWindows()
