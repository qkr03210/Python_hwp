import os
import difflib  # 두 개 문자열간의 차이점을 분석하는 데 쓸 수 있는 외장 라이브러리입니다. 설치되어 있어요.

import pyperclip as cb  # 클립보드를 제어할 수 있는 간편한 툴입니다. pip로 설치하세요.
import win32com.client as win32  # 쓰고 계시죠?
from tkinter import *
from tkinter import messagebox
file = open("D:/psj/ExeTest/target1.txt", "r", encoding='UTF-8')
strings = file.readlines()
messagebox.showinfo("aa",strings)
file.close()