import pyautogui as pt
from time import sleep
from datetime import datetime
from datetime import timedelta
import pyperclip
import pandas
sleep(2)
#trace 153437181
#Neg 42033779
#Pos 34003368
#pending 174068668
def count():
    TC_DropDown = pt.locateCenterOnScreen("TC_DropDown.png", confidence=.8)
    pt.moveTo(TC_DropDown)
    pt.moveRel(30, 0)
    pt.mouseDown()
    pt.dragRel(900, 0, duration=1.5)
    sleep(1)
    pt.mouseUp()
    pt.hotkey('ctrl', 'c')
    temp = pyperclip.paste()
    Personal = temp.splitlines()
    length = len(Personal)
    if '/' in Personal[length-1]:
        print(Personal[length-1])
count()