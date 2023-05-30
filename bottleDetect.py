import cv2
import xlsxwriter

def colorProfiles(n):
    if n == 0 :
        name = "Pepsi"
        hsv_lower = ( 95,100,100)
        hsv_upper = (115,255,255)
        x=+1
        return (name,hsv_lower,hsv_upper)
    if n == 1 :
        name = "Coke"
        hsv_lower = ( 0,100,100)
        hsv_upper = (10,255,255)
        y=+1
        return (name,hsv_lower,hsv_upper)

frame = cv2.imread("images1.jpg")
hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
rects = {}
workbook = xlsxwriter.Workbook('store.xlsx')
worksheet = workbook.add_worksheet()


for i in range(2):
	name, hsv_lower, hsv_upper = colorProfiles(i)
	mask = cv2.inRange(hsv,hsv_lower,hsv_upper)
	conts, herirarchy = cv2.findContours(mask.copy(),cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
	biggest = sorted(conts,key = cv2.contourArea,reverse=True)[0]
	rect = cv2.boundingRect(biggest)
	x,y,w,h = rect
	cv2.rectangle(frame,(x,y),(x+w,y+h),(0,255,255),2)
	cv2.putText(frame, name, (x,y),cv2.FONT_HERSHEY_SIMPLEX,1,(0,0,255),2)


bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'ยี่ห้อ',bold)

worksheet.write(0,1,'ค้นหา',bold)

worksheet.write('A3', 'Pepsi')

worksheet.write('A4', 'Coke')

# เพิ่มข้อความโดยใช้ตัวเลขในการกำหนด Row และ Column  = Row,Column,Text
if x != 0 and y != 0 :
    worksheet.write(2, 1, 'เจอ')
    worksheet.write(3, 1, 'เจอ')
else :
    worksheet.write(2, 1, 'ไม่เจอ')
    worksheet.write(3, 1, 'ไม่เจอ')

# ปิดไฟล์ Excel
workbook.close()

cv2.imshow("Image",frame)
# cv2.imshow("MASK",mask)
cv2.waitKey(0)
cv2.destroyAllWindows()
