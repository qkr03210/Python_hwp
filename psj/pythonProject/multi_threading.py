import threading
from time import sleep
import win32com.client as win32
import win32gui
import shutil
import pymysql
import os

class Main():
    def __init__(self):
        sub = SubMain()

        # th가 실행되고 th2가 실행이 되는데
        # th2는 2초 뒤에 실행이 된다.
        th = threading.Timer(0.5, sub.thread_fnc_myql())  # Timer(시간,함수)
        th.start()

        # sub.thread_stop=True


class SubMain():
    column_name = []
    conn = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')
    def __init__(self):
        print("서브메인 동작")

        try:
            sql = "show fields from hwp_input;"
            cursor = self.conn.cursor()
            cursor.execute(sql)
            column = cursor.fetchall()
            for row in column:
                self.column_name.append(row[0])
        except:
            pass

        print(self.column_name)

    def thread_fnc_myql(self):
        while True:
            conn2 = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')
            print('thread_fnc_myql 동작감지')
            sleep(0.5)
            try:
                print('thread_fnc_myql try 내부')
                sql = "select ifnull( max(idx),0 ), input_queue.index, input_queue.idx, input_queue.name from input_queue;"
                cursor = conn2.cursor()
                cursor.execute(sql)
                result = cursor.fetchall()
                print(result[0][0])
                if result[0][0] != 0:
                    conn2.close()
                    threading.Thread(target=self.thread_mysql_delete(result[0][2])).start()
                    threading.Thread(target=self.thread_fnc(result[0][1],result[0][3])).start()

            except:
                pass
            print('thread_fnc_myql try 외부')

    def thread_mysql_delete(self,target):
        print(target)
        conn2 = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')
        try:
            sql = f"delete from input_queue where idx = {target}"
            cursor = conn2.cursor()
            cursor.execute(sql)
            result = cursor.fetchall()
        except:
            pass
        finally:
            conn2.commit()
            conn2.close()

    def thread_fnc(self,number,name):
        print('찾음')
        sql_result = []
        try:
            sql = f"select * from hwp_input where idx ={number}"
            cursor = self.conn.cursor()
            cursor.execute(sql)
            result = cursor.fetchall()
            for i in range(40):
                sql_result.append(result[0][i])

        except:
            pass

        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        # hwp = win32.Dispatch("HWPFrame.HwpObject")
        # dir="D:/Python_hwp/psj/pythonProject/"

        filename = f"{sql_result[0]}"
        custom = f"{name}"
        # 처음 열고 복사본을 만든다
        hwp.Open(os.path.join(os.getcwd(), filename + ".hwp"))
        hwp.SaveAs(os.path.join(os.getcwd(), filename + f"_{custom}.hwp"))  # 기존 파일명+_임의값.hwp 로 저장
        hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 탭 닫기
        sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)
        # shutil.copyfile(r"./근로자_자격취득신고서.hwp", r"./근로자_자격취득신고서_{}.hwp")
        hwp.Open(os.path.join(os.getcwd(), filename + f"_{custom}.hwp"))
        hwp.XHwpWindows.Item(0).Visible = True #화면에 보이게
        # hwp.XHwpWindows.Item(0).Visible = False  # 화면에 숨김
        hwp.GetFieldList()
        field_list = hwp.GetFieldList().split("\x02")
        # print(field_list)

        for target, data in zip(self.column_name, sql_result):

            if (data == None):
                hwp.PutFieldText(target, "")
                print(target, data)
            else:
                hwp.PutFieldText(target, data)
                print(target, data)

        hwp.SaveAs(os.path.join(os.getcwd(), filename + f"_{custom}") + ".pdf", "PDF")  # 기존 파일명+_임의값.hwp.pdf 로 저장
        # hwp.SaveAs(os.path.join(os.getcwd(), filename+f"_{custom}.hwp"))  # 기존 파일명+_임의값.hwp로 다시 저장
        # hwp.Quit()
        hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 탭 닫기
        hwp.Quit()
        print("문서 저장 완료")

        th=threading.Timer(0.5, self.thread_fnc_myql())
        th.start()

    def thread_fnc_timer(self):
        print("쓰레드 타이머 실행")
        # 반복해서 재귀적으로 실행하는 방식
        th2 = threading.Timer(0.5, self.thread_fnc_myql())
        th2.start()


if __name__ == "__main__":
    Main()
