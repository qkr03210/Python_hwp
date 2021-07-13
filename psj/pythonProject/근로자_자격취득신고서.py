from time import sleep
import win32com.client as win32
import shutil
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
dir="D:/Python_hwp/psj/pythonProject/"
filename="근로자_자격취득신고서"
custom="이름"
# 처음 열고 복사본을 만든다
hwp.Open(dir+filename+".hwp")
hwp.SaveAs(os.path.join(os.getcwd(), filename+f"_{custom}.hwp"))  # 기존 파일명+_임의값.hwp 로 저장
hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 탭 닫기
sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)
# shutil.copyfile(r"./근로자_자격취득신고서.hwp", r"./근로자_자격취득신고서_{}.hwp")
hwp.Open(dir+filename+f"_{custom}.hwp")
hwp.XHwpWindows.Item(0).Visible = True #화면에 보이게
#hwp.XHwpWindows.Item(0).Visible = False #화면에 숨김
hwp.GetFieldList()
field_list = hwp.GetFieldList().split("\x02")
print(field_list)

hwp.PutFieldText("체크", "■")
hwp.PutFieldText("사업장관리번호", "123-45-67890-0")
hwp.PutFieldText("우편번호", "44428")

hwp.PutFieldText("명칭", "복지상사")
hwp.PutFieldText("단위사업장 명칭", "")
hwp.PutFieldText("영업소 명칭", "")
hwp.PutFieldText("소재지", "대구광역시 동구 신천3동 68-1")
hwp.PutFieldText("전화번호", "052-123-4567")
hwp.PutFieldText("팩스번호", "052-123-4567")
# 보험사무대행기관
hwp.PutFieldText("보험사무대행기관번호", "")
hwp.PutFieldText("보험사무대행기관명칭", "")
hwp.PutFieldText("하수급인 관리번호", "")
# 표 1행
hwp.PutFieldText("성명1", "김근로")
hwp.PutFieldText("주민등록번호1", "123456-1234567")
hwp.PutFieldText("국적1", "")
hwp.PutFieldText("체류자격1", "")
hwp.PutFieldText("월 소득액1", "1500000원")
hwp.PutFieldText("자격취득일1", "2018.01.01")
hwp.PutFieldText("국민연금자격취득부호1", "1")
hwp.PutFieldText("국민연금특수직종부호1", "")
hwp.PutFieldText("국민연금직역연금부호1", "")
hwp.PutFieldText("건강보험자격취득부호1", "00")
hwp.PutFieldText("건강보험감면부호1", "")
hwp.PutFieldText("건강보험회계명1", "")
hwp.PutFieldText("건강보험직종명1", "")
hwp.PutFieldText("직종부호1", "155")
hwp.PutFieldText("1주소정근로시간1", "40")
hwp.PutFieldText("계약종료연월1", "201812")
hwp.PutFieldText("보험료부과구분부호1", "")
hwp.PutFieldText("보험료부과구분사유1", "")
# 표 2행
hwp.PutFieldText("성명2", "")
hwp.PutFieldText("주민등록번호2", "")
hwp.PutFieldText("국적2", "")
hwp.PutFieldText("체류자격2", "")
hwp.PutFieldText("월 소득액2", "")
hwp.PutFieldText("자격취득일2", "")
hwp.PutFieldText("국민연금자격취득부호2", "")
hwp.PutFieldText("국민연금특수직종부호2", "")
hwp.PutFieldText("국민연금직역연금부호2", "")
hwp.PutFieldText("건강보험자격취득부호2", "")
hwp.PutFieldText("건강보험감면부호2", "")
hwp.PutFieldText("건강보험회계명2", "")
hwp.PutFieldText("건강보험직종명2", "")
hwp.PutFieldText("직종부호2", "")
hwp.PutFieldText("1주소정근로시간2", "")
hwp.PutFieldText("계약종료연월2", "")
hwp.PutFieldText("보험료부과구분부호2", "")
hwp.PutFieldText("보험료부과구분사유2", "")
# 표 3행
hwp.PutFieldText("성명3", "")
hwp.PutFieldText("주민등록번호3", "")
hwp.PutFieldText("국적3", "")
hwp.PutFieldText("체류자격3", "")
hwp.PutFieldText("월 소득액3", "")
hwp.PutFieldText("자격취득일3", "")
hwp.PutFieldText("국민연금자격취득부호3", "")
hwp.PutFieldText("국민연금특수직종부호3", "")
hwp.PutFieldText("국민연금직역연금부호3", "")
hwp.PutFieldText("건강보험자격취득부호3", "")
hwp.PutFieldText("건강보험감면부호3", "")
hwp.PutFieldText("건강보험회계명3", "")
hwp.PutFieldText("건강보험직종명3", "")
hwp.PutFieldText("직종부호3", "")
hwp.PutFieldText("1주소정근로시간3", "")
hwp.PutFieldText("계약종료연월3", "")
hwp.PutFieldText("보험료부과구분부호3", "")
hwp.PutFieldText("보험료부과구분사유3", "")
# 표 4행
hwp.PutFieldText("성명4", "")
hwp.PutFieldText("주민등록번호4", "")
hwp.PutFieldText("국적4", "")
hwp.PutFieldText("체류자격4", "")
hwp.PutFieldText("월 소득액4", "")
hwp.PutFieldText("자격취득일4", "")
hwp.PutFieldText("국민연금자격취득부호4", "")
hwp.PutFieldText("국민연금특수직종부호4", "")
hwp.PutFieldText("국민연금직역연금부호4", "")
hwp.PutFieldText("건강보험자격취득부호4", "")
hwp.PutFieldText("건강보험감면부호4", "")
hwp.PutFieldText("건강보험회계명4", "")
hwp.PutFieldText("건강보험직종명4", "")
hwp.PutFieldText("직종부호4", "")
hwp.PutFieldText("1주소정근로시간4", "")
hwp.PutFieldText("계약종료연월4", "")
hwp.PutFieldText("보험료부과구분부호4", "")
hwp.PutFieldText("보험료부과구분사유4", "")


# 날짜
hwp.PutFieldText("년", "2021")
hwp.PutFieldText("월", "7")
hwp.PutFieldText("일", "13")
# 신고인
hwp.PutFieldText("신고인", "김모씨")
hwp.PutFieldText("신고인서명", "김모씨서명")
hwp.PutFieldText("보험사무대행기관", "")
hwp.PutFieldText("보험서명", "")

# hwp.PutFieldText("페이지2", "2페이지입니다")
# hwp.PutFieldText("페이지3", "3페이지입니다")

hwp.Quit()
#hwp.SaveAs(os.path.join(os.getcwd(),"test.hwp"))


