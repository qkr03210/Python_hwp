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
    def __init__(self):
        print("서브메인 동작")

        conn = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')
        try:
            sql = "show fields from hwp_input;"
            cursor = conn.cursor()
            cursor.execute(sql)
            column = cursor.fetchall()
            for row in column:
                self.column_name.append(row[0])
        except:
            pass
        finally:
            conn.close()
        print(self.column_name)

    def thread_fnc_myql(self):
        while True:
            print('mysql 동작감지')
            sleep(0.5)
            conn = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')
            try:
                sql = "select ifnull( max(idx),0 ), input_queue.index from input_queue;"
                cursor = conn.cursor()
                cursor.execute(sql)
                result = cursor.fetchall()
                if result[0][0] != 0:
                    print(result[0][0])
                    conn.close()
                    threading.Thread(target=self.thread_mysql_delete(result[0][0])).start()
                    threading.Thread(target=self.thread_fnc(result[0][1])).start()

            except:
                pass
            finally:
                if conn._closed == False:
                    conn.close()

    def thread_mysql_delete(self,target):
        print(target)
        conn = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')

        try:
            sql = "delete from input_queue where idx =" + target
            cursor = conn.cursor()
            cursor.execute(sql)
        except:
            pass
        finally:
            conn.close()

    def thread_fnc(self,number):
        print('찾음')
        conn = pymysql.connect(host='192.168.0.104', user='root', password='1234', db='hwp', charset='utf8')

        sql_result = []
        try:
            sql = "select * from hwp_input where idx ="+number
            cursor = conn.cursor()
            cursor.execute(sql)
            result = cursor.fetchall()
            for i in range(40):
                sql_result.append(result[0][i])

        except:
            pass
        finally:
            conn.close()

        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        # hwp = win32.Dispatch("HWPFrame.HwpObject")
        # dir="D:/Python_hwp/psj/pythonProject/"

        filename = "근로자_자격취득신고서"
        custom = "이름"
        # 처음 열고 복사본을 만든다
        hwp.Open(os.path.join(os.getcwd(), filename + ".hwp"))
        hwp.SaveAs(os.path.join(os.getcwd(), filename + f"_{custom}.hwp"))  # 기존 파일명+_임의값.hwp 로 저장
        hwp.XHwpDocuments.Item(0).Close(isDirty=False)  # 탭 닫기
        sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)
        # shutil.copyfile(r"./근로자_자격취득신고서.hwp", r"./근로자_자격취득신고서_{}.hwp")
        hwp.Open(os.path.join(os.getcwd(), filename + f"_{custom}.hwp"))
        # hwp.XHwpWindows.Item(0).Visible = True #화면에 보이게
        hwp.XHwpWindows.Item(0).Visible = False  # 화면에 숨김
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
