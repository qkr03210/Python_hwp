from time import sleep
import win32com.client as win32
import win32gui
import shutil
import pymysql
import os
conn = pymysql.connect(host='192.168.0.104',user='root',password='1234',db='hwp',charset='utf8')

sql_result=[]
try:
    sql = "select * from hwp_input"
    cursor = conn.cursor()
    cursor.execute(sql)
    result=cursor.fetchall()
    for i in range(39):
        sql_result.append(result[0][i])

except:
    pass
finally:
    conn.close()
print(sql_result)
column_name=[]
conn = pymysql.connect(host='192.168.0.104',user='root',password='1234',db='hwp',charset='utf8')
try:
    sql = "show fields from hwp_input;"
    cursor = conn.cursor()
    cursor.execute(sql)
    column = cursor.fetchall()
    for row in column:
        column_name.append(row[0])
except:
    pass
finally:
    conn.close()

print(column_name)
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
# hwp = win32.Dispatch("HWPFrame.HwpObject")
# dir="D:/Python_hwp/psj/pythonProject/"

filename="근로자_자격취득신고서"
custom="이름"
# 처음 열고 복사본을 만든다
hwp.Open(os.path.join(os.getcwd(),filename+".hwp"))
hwp.SaveAs(os.path.join(os.getcwd(), filename+f"_{custom}.hwp"))  # 기존 파일명+_임의값.hwp 로 저장
hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 탭 닫기
sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)
# shutil.copyfile(r"./근로자_자격취득신고서.hwp", r"./근로자_자격취득신고서_{}.hwp")
hwp.Open(os.path.join(os.getcwd(),filename+f"_{custom}.hwp"))
# hwp.XHwpWindows.Item(0).Visible = True #화면에 보이게
hwp.XHwpWindows.Item(0).Visible = False #화면에 숨김
hwp.GetFieldList()
field_list = hwp.GetFieldList().split("\x02")
# print(field_list)


for target,data in zip(column_name,sql_result):

    if(data==None):
        hwp.PutFieldText(target, "")
        print(target, data)
    else:
        hwp.PutFieldText(target, data)
        print(target, data)

hwp.SaveAs(os.path.join(os.getcwd(), filename+f"_{custom}")+".pdf", "PDF")  # 기존 파일명+_임의값.hwp.pdf 로 저장
# hwp.SaveAs(os.path.join(os.getcwd(), filename+f"_{custom}.hwp"))  # 기존 파일명+_임의값.hwp로 다시 저장
# hwp.Quit()
hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 탭 닫기
print("문서 저장 완료")


