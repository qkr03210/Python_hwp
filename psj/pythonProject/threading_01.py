import threading
import time


class Main():
    def __init__(self):
        print("메인 동작")


        sub=SubMain()
        th = threading.Thread(target=sub.thread_fnc)
        th.start()
        #th가 실행되고 th2가 실행이 되는데
        #th2는 2초 뒤에 실행이 된다.
        th2= threading.Timer(2,sub.thread_fnc_timer()) # Timer(시간,함수)
        th2.start()

        #sub.thread_stop=True

class SubMain():
    def __init__(self):
        print("서브메인 동작")
    #   self.thread_stop=False
    def thread_fnc(self):
        while self.thread_stop is False:
            time.sleep(1)
            print("thread_fnc 실행")

    def thread_fnc_timer(self):
        print("쓰레드 타이머 실행")
        #반복해서 재귀적으로 실행하는 방식
        th2= threading.Timer(2,self.thread_fnc_timer())
        th2.start()

if __name__ == "__main__":
    Main()
        