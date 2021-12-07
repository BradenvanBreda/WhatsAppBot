#V3 - TrakCare integration

import pyautogui as pt
from time import sleep
from datetime import datetime
from datetime import timedelta
import pyperclip
import pandas
sleep(1)

def check_for_new_message():
    global df
    df = pandas.read_csv("Patient_Log.txt")
    # gui check for new message or green dots
    while True:
        sleep(1)
        Paperclip = pt.locateCenterOnScreen("Paperclip.png", confidence=.6)
        if Paperclip is not None:
            if pt.pixelMatchesColor(int(Paperclip[0] - 20), int(Paperclip[1] - 55), (255, 255, 255), tolerance=3):
                store_and_ask()

        green_dot = pt.locateCenterOnScreen("green_dot.png", confidence=.8)
        if green_dot is not None:
            pt.moveTo(green_dot)
            pt.moveRel(-100, 0)
            pt.click()
            sleep(1)
            get_phone_num()

def get_phone_num():
    global df
    global row
    Mag_dots = pt.locateCenterOnScreen("mag_dots.png", confidence=.7)
    pt.moveTo(Mag_dots)
    pt.moveRel(-200, 10)
    pt.click()
    sleep(1)
    Business_acc = pt.locateOnScreen("Business_acc.png", confidence=.7)
    if Business_acc is not None:
        pt.moveTo(Business_acc)
        pt.scroll(-500)
        Busi_Number = pt.locateCenterOnScreen("Busi_Number.png", confidence=.8)
        if Busi_Number is not None:
            pt.moveTo(Busi_Number)
            pt.moveRel(0, -35)
            pt.tripleClick()
            sleep(0.25)
            pt.hotkey('ctrl', 'c')
        else:
            return
    else:
        Number = pt.locateCenterOnScreen("Number.png", confidence=.7)
        pt.moveTo(Number)
        pt.tripleClick()
        sleep(0.25)
        pt.hotkey('ctrl', 'c')

    phone_num = pyperclip.paste()

    Close = pt.locateCenterOnScreen("Close.png", confidence=.8)
    pt.moveTo(Close)
    pt.click()
    sleep(0.25)
    df.to_csv("auto_backup.txt", index=False)

    i = len(df)
    match = False
    while i > 0: #Find matching numbers
        if str(df.iloc[i-1, 0]) == str(phone_num):
            match = True
            row = i - 1 # used for continued screening and for session indexing
            break # only latest of interest, save time
        i = i - 1
    if match == True:
        return
    else: #match == false
        Initiate_NL(phone_num)
    return

def Initiate_NL(phone_num):
    global df
    global row
    New_row = {"Cell": str(phone_num),
               "Attempts": "0",
               "Consent": "0",
               "Ticket": "0",
               "Folder": "0",
               "BirthDate": "0",
               "Name": "0",
               "Surname": "0",
               "TestDate": "0",
               "Status": "0",
               "Offset": "1",
               "8:15": "0", "9:00": "0", "9:45": "0", "10:30": "0", "11:15": "0", "13:00": "0", "13:45": "0",
               "14:30": "0", "15:15": "0", "16:00": "0",
               "Appointment": "0",
               "Outcome": "0"}
    df = df.append(New_row, ignore_index=True)
    row = len(df) - 1
    send_message('S0_Introduction.txt')
    send_message('Q1_Consent.txt')
    df.to_csv("Patient_Log.txt", index=False)

def store_and_ask():
    global df
    global row
    # closed - check number outcome and reopen,new row etc in get and check phone number
    get_message()
    new_message = pyperclip.paste().lower()

    if new_message == "pause system":
        send_message('System_Paused.txt')
        sleep(100)
    elif str(df.iloc[row, 22]) == "0":
        if new_message == "":
            send_message('S_Message_Del.txt')
        elif new_message == "stop":
            send_message('SG_Stop.txt')
            send_message('Restart.txt')
            zero_reg()
        elif new_message == "restart": #jump straight to open
            send_message('S0_Introduction.txt')
            send_message('Q1_Consent.txt')
            zero_reg()
            df.iloc[row, 22] = "0"
        elif str(df.iloc[row, 2]) == "0": #Consent?
            if new_message == "1" or new_message == "1 " or new_message == "yes" or new_message == "ja" or new_message == "ewe":
                df.iloc[row, 2] = "1"
                ask_for_ticket()
            elif new_message == "2" or new_message == "2 " or new_message == "no" or new_message == "nee" or new_message == "hayi":
                df.iloc[row, 2] = "2"
                incorrect_answer('S1_Consent.txt')
            else:
                send_message('S_Error_1_2_only.txt')
        elif str(df.iloc[row, 3]) == "0": #Ticket?
            if new_message == "1" or new_message == "1 " or new_message == "yes" or new_message == "ja" or new_message == "ewe":
                df.iloc[row, 3] = "1"
                ask_for_folder_num()
            elif new_message == "2" or new_message == "2 " or new_message == "no" or new_message == "nee" or new_message == "hayi":
                df.iloc[row, 3] = "2"
                incorrect_answer('S_ticket.txt')
            else:
                send_message('S_Error_1_2_only.txt')
        elif str(df.iloc[row, 4]) == "0":
            Fol_Num_Entry = list(new_message)
            length = len(Fol_Num_Entry)
            i = 0
            Fol_Num_Char = []
            while i < length:
                if Fol_Num_Entry[i].isnumeric():
                    Fol_Num_Char.append(Fol_Num_Entry[i])
                i = i + 1
            Fol_Num_Str = str(''.join(Fol_Num_Char))
            if Fol_Num_Str.isnumeric():
                df.iloc[row, 4] = str(Fol_Num_Str)
                Personal,Status,TestDate = TC_Info(Fol_Num_Str)
                length = len(Personal)
                if length == 5:
                    df.iloc[row, 5] = "2"
                    df.iloc[row, 6] = "2"
                    df.iloc[row, 7] = "2"
                    send_message('S_Error_1_2_only.txt')
                else:
                    #go manual

                #check personal deeets length to see if it includes a B day.go manual.

            else:
                incorrect_answer('Invalid_Fol_Num.txt')
            sleep(100)

#Too be continued
    elif str(df.iloc[row, 22]) == "1" or str(df.iloc[row, 22]) == "2":  # Closed Successfully
        if new_message == "restart" or new_message == "restart." or new_message == "restart ":
            Initiate_NL(str(df.iloc[row, 0])) #send current phone number
        else:
            send_message('Restart.txt')
    elif str(df.iloc[row, 22]) == "3":  # Closed and Unsuccessful
        if new_message == "restart" or new_message == "restart." or new_message == "restart ":
            send_message('S0_Introduction.txt')
            send_message('Q1_Consent.txt')
            zero_reg()
            df.iloc[row, 22] = "0"
        else:
            send_message('Restart.txt')
    elif str(df.iloc[row, 22]) == "4":  # Closed and Blocked
        send_message('Blocked.txt')
    else:
        print("no If entered")
    df.to_csv("Patient_Log.txt", index=False)
    return

def zero_reg():
    global df
    global row
    i = 22
    while i > 1: #does not zero attempts
        df.iloc[row, i] = "0"
        i = i - 1
    df.iloc[row, 10] = "1" #offset set to 1
    df.iloc[row, 22] = "3" #default Closed Unsuccessful
    return

def TC_Info(Fol_Num_Str):
    global df
    global row

    #Move to TC window
    mag_dots, Error = Look_For("mag_dots.png")
    if Error:
        return True
    else:
        pt.moveTo(mag_dots)
        pt.moveRel(80,0)
        pt.click()
        sleep(1)

    while True:
        TC_logonWindow = pt.locateCenterOnScreen("TC_logonWindow.png", confidence=.8)
        if TC_logonWindow is not None:
            TC_LogonUser = pt.locateCenterOnScreen("TC_LogonUser.png", confidence=.8)
            pt.moveTo(TC_LogonUser)
            pt.click()
            sleep(0.25)
            pyperclip.copy('NMBB5280')
            pt.hotkey('ctrl', 'v')
            sleep(1)
            TC_LogonPassword = pt.locateCenterOnScreen("TC_LogonPassword.png", confidence=.8)
            pt.moveTo(TC_LogonPassword)
            pt.click()
            sleep(0.25)
            pyperclip.copy('123America#NHLS')
            pt.hotkey('ctrl', 'v')
            sleep(1)
            pt.click()
            sleep(1)
            TC_LogonButton = pt.locateCenterOnScreen("TC_LogonButton.png", confidence=.8)
            pt.moveTo(TC_LogonButton)
            pt.click()
            break
        if i > 100:
            TC_LockedWindow = pt.locateCenterOnScreen("TC_LockedWindow.png", confidence=.8)
            if TC_LockedWindow is not None:
                TC_LockedPassword = pt.locateCenterOnScreen("TC_LockedPassword.png", confidence=.8)
                pt.moveTo(TC_LockedPassword)
                pt.click()
                sleep(1)
                pyperclip.copy('123America#NHLS')
                pt.hotkey('ctrl', 'v')
                sleep(1)
                pt.click()
                sleep(1)
                TC_LogonButton = pt.locateCenterOnScreen("TC_LogonButton.png", confidence=.8)
                pt.moveTo(TC_LogonButton)
                pt.click()
                break
            else:
                TC_Logout()
                return True
        i = i + 1
        sleep(1)

    # Enter Folder number
    Fol_Num,Error = Look_For("Fol_Num.png")
    if Error:
        return True
    else:
        pt.moveTo(Fol_Num)
        pt.click()
        pyperclip.copy(Fol_Num_Str)
        pt.hotkey('ctrl', 'v')
        sleep(1)

    # Open advanced Search
    TC_Advanced, Error = Look_For("TC_Advanced.png")
    if Error:
        return True
    else:
        pt.moveTo(TC_Advanced)
        pt.click()
        sleep(1)

    # Enter location
    TC_Location, Error = Look_For("TC_Location.png")
    if Error:
        return True
    else:
        pt.moveTo(TC_Location)
        pt.click()
        pyperclip.copy('D91')
        pt.hotkey('ctrl', 'v')
        sleep(1)

    # Select WC location
    TC_WC, Error = Look_For("TC_WC.png")
    if Error:
        return True
    else:
        pt.moveTo(TC_WC)
        pt.click()
        sleep(1)

    # Close advanced Search
    TC_Advanced, Error = Look_For("TC_Advanced.png")
    if Error:
        return True
    else:
        pt.moveTo(TC_Advanced)
        pt.click()
        sleep(1)

    # Click on Search
    TC_Search, Error = Look_For("TC_Search.png")
    if Error:
        return True
    else:
        pt.moveTo(TC_Search)
        pt.click()
        sleep(1)

    # Valid Folder Number Check and Collect personal detail
    while True:
        TC_DropDown = pt.locateCenterOnScreen("TC_DropDown.png", confidence=.8)
        if TC_DropDown is not None:
            break
        if i > 100:
            TC_Logout()
            send_bvb_message()
            incorrect_answer("Invalid_Fol_Num.txt")
            return True
        i = i + 1
        sleep(0.5)

    #Collect personal detail
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
    if length == 4:
        df.iloc[row, 5] = Personal[3]
        df.iloc[row, 6] = Personal[0]
        df.iloc[row, 7] = Personal[1]
    else:
        TC_Logout()
        send_bvb_message()
        send_work_message()
        send_message("Personal_Detail_Incomplete.txt")
        send_message('Manual.txt')
        zero_reg()
        return True

    # Open dropdown Action
    TC_DropDown, Error = Look_For("TC_DropDown.png")
    if Error:
        return True
    else:
        pt.moveTo(TC_DropDown)
        pt.moveRel(-15, 0)
        pt.click()
        sleep(1)

    # GXPU been conducted?
    while True:
        TC_GXPU = pt.locateCenterOnScreen("TC_GXPU.png", confidence=.8)
        if TC_GXPU is not None:
            break
        if i > 20:
            TC_Logout()
            send_bvb_message()
            incorrect_answer("No_GXPU.txt")
            return True
        i = i + 1
        sleep(1)

    # GXPU Proccesed yet?
    pt.moveTo(TC_GXPU)
    pt.click()
    while True:
        TC_Result_Page = pt.locateCenterOnScreen("TC_Result_Page.png", confidence=.8)
        if TC_Result_Page is not None: #GXPU processed
            Error = GetTestDate()
            if Error:
                return True
            Error = GetStatus()
            if Error:
                return True
            break
        if i > 30: #GXPU Not yet processed
            df.iloc[row, 8] = str(datetime.now().date())
            df.iloc[row, 9] = "4"
            break
        i = i + 1
        sleep(1)
    TC_Logout()
    return False

def GetStatus():
    global dt
    global row
    Error = False
    while True:
        TC_Detected = pt.locateCenterOnScreen('TC_Detected.png', confidence=.8)
        if TC_Detected is not None:
            df.iloc[row, 9] = "1"
            break
        TC_Not_Detected = pt.locateCenterOnScreen('TC_Not_Detected.png', confidence=.8)
        if TC_Not_Detected is not None:
            df.iloc[row, 9] = "2"
            break
        TC_Trace = pt.locateCenterOnScreen('TC_Trace.png', confidence=.8)
        if TC_Trace is not None:
            df.iloc[row, 9] = "3"
            break
        if i > 20:
            TC_Logout()
            send_bvb_message()
            send_work_message()
            send_message("GXPU_Result_Undefined.txt")
            send_message('Manual.txt')
            zero_reg()
            Error = True
            break
        sleep(1)
        i = i + 1
    return Error

def GetTestDate():
    global dt
    global row
    Error = False
    while True:
        TC_TestDate = pt.locateCenterOnScreen('TC_TestDate.png', confidence=.8)
        if TC_TestDate is not None:
            pt.moveTo(TC_TestDate)
            pt.moveRel(60, 0)
            pt.tripleClick()
            pt.hotkey('ctrl','c')
            TestDate = pyperclip.paste()
            TestdateSeperated = TestDate.split("/")
            TestdateFormated = TestdateSeperated[0] + "-" + TestdateSeperated[1] + "-" + TestdateSeperated[2]
            df.iloc[row, 8] = TestdateFormated
            break
        if i > 20:
            TC_Logout()
            send_bvb_message()
            send_work_message()
            send_message("TestDate_Undefined.txt")
            send_message('Manual.txt')
            zero_reg()
            break
        sleep(1)
        i = i + 1
    return Error


def Look_For(Image):
    global df
    global row
    i = 0
    Error = False
    while True:
        Temp = pt.locateCenterOnScreen(Image, confidence=.8)
        if Temp is not None:
            break
        if i > 100:
            Error = True
            TC_Logout()
            send_bvb_message()
            send_work_message()
            send_message("SysDown_Manual.txt")
            df.iloc[row, 22] = "5" # manual
            break
        i = i + 1
        sleep(1)
    return Temp,Error

def TC_Logout():
    while True:
        TC_URLMarker = pt.locateCenterOnScreen("TC_URLMarker.png", confidence=.8)
        if TC_URLMarker is not None:
            pt.moveTo(TC_URLMarker)
            pt.moveRel(200, 15)
            pt.tripleClick()
            pyperclip.copy('https://trakcarelabwebview.nhls.ac.za/trakcarelab/csp/system.Home.cls#/Logout')
            sleep(0.25)
            pt.hotkey('ctrl', 'v')
            sleep(0.25)
            pt.hotkey('enter')
            sleep(0.25)
            pt.moveTo(150, 35)
            pt.click()
            sleep(1)
            break
        if i > 10:
            pt.moveTo(150,35)
            pt.click()
            sleep(1)
            send_bvb_message("TC_URLMarker_NotFound.txt")
            sleep
            break
        i = i + 1
        sleep(1)
    return

def send_bvb_message(message_path):
    sleep(0.5)
    txt = open(message_path, 'r')
    temp = txt.read()
    txt.close()
    pyperclip.copy(temp)
    Paperclip = pt.locateOnScreen("Paperclip.png", confidence=.6)
    pt.moveTo(Paperclip)
    pt.moveRel(130, 0)
    pt.click()
    sleep(0.25)
    pt.hotkey('ctrl', 'v')
    pt.typewrite("\n", interval=0.01)
    return

def send_work_message():


def get_message():
    sleep(0.5)
    Paperclip = pt.locateOnScreen("Paperclip.png", confidence=.6)
    pt.moveTo(Paperclip)
    pt.moveRel(-20, -55)
    pt.tripleClick()
    sleep(0.25)
    pt.hotkey('ctrl', 'c')
    return

def send_message(message_path):
    sleep(0.5)
    txt = open(message_path, 'r')
    temp = txt.read()
    txt.close()
    pyperclip.copy(temp)
    Paperclip = pt.locateOnScreen("Paperclip.png", confidence=.6)
    pt.moveTo(Paperclip)
    pt.moveRel(130, 0)
    pt.click()
    sleep(0.25)
    pt.hotkey('ctrl', 'v')
    pt.typewrite("\n", interval=0.01)
    return

def ask_for_ticket():
    txt = open('Q_ticket.txt', 'r')
    temp = txt.read()
    txt.close()
    pyperclip.copy(temp)
    paperclip = pt.locateCenterOnScreen("paperclip.png", confidence=.8)
    pt.moveTo(paperclip)
    pt.click()
    sleep(1)
    pt.moveRel(0, -70)
    pt.click()
    sleep(3)
    Ticket_Example = pt.locateCenterOnScreen("Ticket_Example.png", confidence=.7)
    pt.moveTo(Ticket_Example)
    pt.doubleClick()
    sleep(2)
    Add_Caption = pt.locateOnScreen("caption.png", confidence=.7)
    pt.moveTo(Add_Caption)
    pt.click()
    sleep(0.25)
    pt.hotkey('ctrl', 'v')
    pt.typewrite("\n", interval=0.01)
    return

def ask_for_folder_num():
    txt = open('Q_Folder_Number.txt', 'r')
    temp = txt.read()
    txt.close()
    pyperclip.copy(temp)
    paperclip = pt.locateCenterOnScreen("paperclip.png", confidence=.8)
    pt.moveTo(paperclip)
    pt.click()
    sleep(1)
    pt.moveRel(0, -70)
    pt.click()
    sleep(3)
    Folder_Number = pt.locateCenterOnScreen("Folder_Number.png", confidence=.7)
    pt.moveTo(Folder_Number)
    pt.doubleClick()
    sleep(2)
    Caption = pt.locateOnScreen("caption.png", confidence=.7)
    pt.moveTo(Caption)
    pt.click()
    sleep(0.25)
    pt.hotkey('ctrl', 'v')
    pt.typewrite("\n", interval=0.01)
    return

def incorrect_answer(reason):
    global df
    global row
    df.iloc[row, 1] = str(int(df.iloc[row, 1]) + 1)
    if str(df.iloc[row, 1]) == "3":
        df.iloc[row, 22] = "4"
        send_message(reason)
        send_message('Blocked.txt')
    else:
        send_message(reason)
        send_message('Restart.txt')
        zero_reg()
    return

check_for_new_message()


