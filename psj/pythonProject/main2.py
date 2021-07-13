import win32com.client as win32
import pymysql
import os
# conn = pymysql.connect(host='192.168.0.104',user='root',password='1234',db='hwp',charset='utf8')
# try:
#     sql = "select * from member where bn =%s"
#     cursor = conn.cursor()
#     cursor.execute(sql, ('2'))
#     result=cursor.fetchall()
#     # print(result)
#     for row in result:
#         r1=row[0]
#         r2=row[1]
#         r3=row[2]
# except:
#     pass
# finally:
#     conn.close()

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
# hwp = win32.Dispatch("HWPFrame.HwpObject")
filename="D:/psj/pythonProject/근로자_기준소득월액변경신청서.hwp"
hwp.Open(filename)
hwp.XHwpWindows.Item(0).Visible = True

hwp.GetFieldList()
field_list = hwp.GetFieldList().split("\x02")
print(field_list)

hwp.PutFieldText("체크", "■")
hwp.PutFieldText("사업장관리번호", "11111111")
hwp.PutFieldText("명칭", "경북산업직업전문학교")
hwp.PutFieldText("전화번호", "053-784-7845")
hwp.PutFieldText("소재지", "대구광역시 동구 신천3동 68-1")
hwp.PutFieldText("년", "2021")
hwp.PutFieldText("월", "7")
hwp.PutFieldText("일", "8")

# hwp.PutFieldText("페이지2", "2페이지입니다")
# hwp.PutFieldText("페이지3", "3페이지입니다")

# hwp.SaveAs(os.path.join(os.getcwd(),"test.hwp"))


