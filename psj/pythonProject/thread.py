# https://www.youtube.com/watch?v=lKbwxUDNoWo&list=PLDtzZPtOGenaG_LeSAHpr4opgz0HebcwJ&index=14&t=934s
# 참고영상
import threading
import time

def thread_fnc1():
    while True:
        time.sleep(0.5)
        print("쓰레드에서 while문 동작")

print("프로그램 실행")

th = threading.Thread(target=thread_fnc1)
th.daemon  =True #메인 쓰레드가 종료될 경우 서브도 같이 종료가 된다.
th.start()
print("while문 이후 실행")