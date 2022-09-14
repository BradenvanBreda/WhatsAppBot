import pyautogui as pt
from time import sleep
from datetime import datetime
from datetime import timedelta
sleep(0.5)
Posi = pt.locateCenterOnScreen("Today.png", confidence=.7)
x = Posi[0]
y = Posi[1]

M_B_Time = ["8:20", "8:40", "9:00", "9:20", "9:40", "10:00", "10:20", "10:40", "11:00", "11:20", "11:40"]
A_B_Time = ["13:00", "13:20", "13:40", "14:00", "14:20", "14:40", "15:00", "15:20", "15:40", "16:00", "16:20"]
print()

i=0
while(i<11):
    pt.moveTo(x+500, y+471 + i*14.5)
    sleep(0.5)
    print("*" + str(i+1) + "*" + " - @ " + M_B_Time[i] + "\n")
    i = i + 1

i=0
while(i<11):
    pt.moveTo(x+500, y+673 + i*14.5)
    sleep(0.5)
    i=i+1

