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

    Fol_Num = Look_For("Fol_Num.png")
    pt.moveTo(Fol_Num)
    pt.click()
    pyperclip.copy('174068668')
    pt.hotkey('ctrl','v')

    TC_Advanced = Look_For("TC_Advanced.png")
    pt.moveTo(TC_Advanced)
    pt.click()
    sleep(1)

    TC_Location = Look_For("TC_Location.png")
    pt.moveTo(TC_Location)
    pt.click()
    pyperclip.copy('P9')
    pt.hotkey('ctrl', 'v')
    sleep(1)

    TC_CT_Metro = Look_For("TC_WC.png")
    pt.moveTo(TC_CT_Metro)
    pt.click()
    sleep(1)

    TC_Advanced = Look_For("TC_Advanced.png")
    pt.moveTo(TC_Advanced)
    pt.click()
    sleep(1)

    TC_Search = Look_For("TC_Search.png")
    pt.moveTo(TC_Search)
    pt.click()
    sleep(1)

    TC_DropDown = Look_For("TC_DropDown.png")
    pt.moveTo(TC_DropDown)
    pt.moveRel(-15,0)
    pt.click()
    sleep(1)

    TC_GXPU = Look_For("TC_GXPU.png")
    pt.moveTo(TC_GXPU)
    pt.click()
    while True:
        Temp = pt.locateCenterOnScreen(Image, confidence=.99)
        if Temp is not None:
            break
    return Temp

    TC_Result_Page_Marker = Look_For("TC_Result_Page_Marker.png")
    pt.moveTo(TC_Result_Page_Marker)
    pt.moveRel(85, -85)
    pt.mouseDown()
    pt.dragRel(1300, 0, duration=1.5)
    sleep(1)
    pt.mouseUp()
    pt.hotkey('ctrl', 'c')
    temp = pyperclip.paste()
    Personal_Detals = temp.splitlines()
    print(Personal_Detals[2])
    print(Personal_Detals[3])
    print(Personal_Detals[4])

    while True:
        TC_Detected = pt.locateCenterOnScreen('TC_Detected.png', confidence=.8)
        if TC_Detected is not None:
            print("pos")
            break
        TC_Not_Detected = pt.locateCenterOnScreen('TC_Not_Detected.png', confidence=.8)
        if TC_Not_Detected is not None:
            print("Neg")
            break
        TC_Trace = pt.locateCenterOnScreen('TC_Trace.png', confidence=.8)
        if TC_Trace is not None:
            print("Trace")
            break
        sleep(0.5)



def Look_For(Image):

    while True:
        Temp = pt.locateCenterOnScreen(Image, confidence=.9)
        if Temp is not None:
            break
    return Temp
count()