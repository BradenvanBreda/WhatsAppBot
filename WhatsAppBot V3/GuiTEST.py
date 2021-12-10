import pyautogui as pt
from time import sleep
from datetime import datetime
from datetime import timedelta
import pyperclip
import pandas
sleep(2)
#trace 153437181+27 78 006 0909
#Neg 42033779
#Pos 34003368
#pending 174068668

def calculate_age():
    temp = pyperclip.paste()
    Personal = temp.splitlines()
    sleep(1)
    length = len(Personal)
    print(length)
    if all(i.isalpha() or i.isspace() for i in Personal[0].strip()):
        print(Personal[0].strip())
    else:
        print('nope')
    if all(i.isalpha() or i.isspace() for i in Personal[1].strip()):
        print(Personal[1].strip())
    else:
        print('nope')


calculate_age()

