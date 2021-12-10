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
               "Confirmed": "0",
               "Offset": "1",
               "8:15": "0", "9:00": "0", "9:45": "0", "10:30": "0", "11:15": "0", "13:00": "0", "13:45": "0",
               "14:30": "0", "15:15": "0", "16:00": "0",
               "OptionDate": "0",
               "AppointTime": "0",
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
    elif str(df.iloc[row, 24]) == "0":
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
            df.iloc[row, 24] = "0"
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
                Error = TC_Info(Fol_Num_Str)
                if not Error:
                    # Create name message to send
                    txt = open('DOB_Correct.txt', 'r')
                    temp = txt.read()
                    txt.close()
                    MessageToSend = open("MessageToSend.txt", "w")
                    MessageToSend.write('*' + str(df.iloc[row, 5]) + '*' + "\n" + "\n" + str(temp))
                    MessageToSend.close()
                    # check age
                    dt_obj = datetime.strptime(str(df.iloc[row, 5]), '%Y-%m-%d').date()
                    today = datetime.now().date()
                    age = today.year - dt_obj.year - ((today.month, today.day) < (dt_obj.month, dt_obj.day))
                    if age >= 14:
                        if str(df.iloc[row, 9]) == 1 or str(df.iloc[row, 9]) == 2:  # GXPU pos trace
                            dt_obj = datetime.strptime(str(df.iloc[row, 8]), '%Y-%m-%d').date()
                            today = datetime.now().date()
                            Days_ago = (today - dt_obj).days
                            if Days_ago <= 13:
                                send_message('MessageToSend.txt')
                            else:
                                send_message("GXPU_Pos_Time_Exclusion.txt")
                                zero_reg()
                        elif str(df.iloc[row, 9]) == 3:  # GXPU Neg
                            dt_obj = datetime.strptime(str(df.iloc[row, 8]), '%Y-%m-%d').date()
                            today = datetime.now().date()
                            Days_ago = (today - dt_obj).days
                            if Days_ago <= 2:
                                send_message('MessageToSend.txt')
                            else:
                                send_message("GXPU_Neg_Time_Exclusion.txt")
                                zero_reg()
                        else:  # pending
                            send_message('MessageToSend.txt')
                    else:
                        send_message("Too_Young.txt")
                        zero_reg()
                else:  # errors messages and register resets are dealt with in TC_INFO
                    pass
            else:
                incorrect_answer('Invalid_Fol_Num.txt')
        elif str(df.iloc[row, 10]) == "0": #Personal detail correct
            if new_message == "1" or new_message == "1 " or new_message == "yes" or new_message == "ja" or new_message == "ewe":
                df.iloc[row, 10] = "1"
                find_available_date()
            elif new_message == "2" or new_message == "2 " or new_message == "no" or new_message == "nee" or new_message == "hayi":
                df.iloc[row, 10] = "2"
                incorrect_answer('Incorrect_FolNum.txt')
            else:
                send_message('S_Error_1_2_only.txt')
        elif str(df.iloc[row, 23]) == "0":
            if new_message == "0" or new_message == "0 ":
                df.iloc[row, 11] = int(str(df.iloc[row, 11])) + 1
                find_available_date()
            elif new_message == "11" or new_message == "11 ":
                df.iloc[row, 11] = "1"
                find_available_date()
            elif new_message == "12" or new_message == "12 ":
                send_message('S_Cancel_Options.txt')
                if str(df.iloc[row, 9]) == "4":
                    send_message('S_Waiting_Restart.txt')
                    df.iloc[row, 24] = "3"
            elif new_message.isnumeric():
                selected = int(new_message)
                check_selection(selected)
            else:
                send_message('S_Error_Num_only.txt')
    elif str(df.iloc[row, 24]) == "1" or str(df.iloc[row, 24]) == "2":  # Closed Successfully
        if new_message == "restart" or new_message == "restart." or new_message == "restart ":
            Initiate_NL(str(df.iloc[row, 0])) #send current phone number
        else:
            send_message('Restart.txt')
    elif str(df.iloc[row, 24]) == "3" or str(df.iloc[row, 24]) == "5":  # Closed and Unsuccessful
        if new_message == "restart" or new_message == "restart." or new_message == "restart ":
            send_message('S0_Introduction.txt')
            send_message('Q1_Consent.txt')
            zero_reg()
            df.iloc[row, 24] = "0"
        else:
            send_message('Restart.txt')
    elif str(df.iloc[row, 24]) == "4":  # Closed and Blocked
        send_message('Blocked.txt')
    else:
        print("no If Statement entered")
    df.to_csv("Patient_Log.txt", index=False)
    return

def zero_reg():
    global df
    global row
    i = 24
    while i > 1: #does not zero attempts
        df.iloc[row, i] = "0"
        i = i - 1
    df.iloc[row, 11] = "1" #offset set to 1
    df.iloc[row, 24] = "3" #default Closed Unsuccessful
    df.to_csv("Patient_Log.txt", index=False)
    return

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

def send_message_to(message_path,CellNum,CellImage):
    WA_SearchNum = pt.locateCenterOnScreen("WA_SearchNum.png", confidence=.9)
    pt.moveTo(WA_SearchNum)
    pt.click()
    sleep(0.25)
    pyperclip.copy(CellNum)
    pt.hotkey('ctrl', 'v')
    i = 0
    sleep(1)
    while True:
        WA_Who = pt.locateCenterOnScreen(CellImage, confidence=.9)
        if WA_Who is not None:
            pt.moveTo(WA_Who)
            pt.click()
            sleep(0.25)
            send_message(message_path)
            break
        if i > 10:
            break
        i = i + 1
        sleep(1)
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
        df.iloc[row, 24] = "4"
        send_message(reason)
        send_message('Blocked.txt')
        df.to_csv("Patient_Log.txt", index=False)
    else:
        send_message(reason)
        send_message('Restart.txt')
        zero_reg()
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

    #Login/unlock
    i=0
    while True:
        TC_logonWindow = pt.locateCenterOnScreen("TC_logonWindow.png", confidence=.9)
        if TC_logonWindow is not None:
            TC_LogonUser, Error = Look_For("TC_LogonUser.png")
            if Error:
                return True
            else:
                pt.moveTo(TC_LogonUser)
                pt.click()
                sleep(0.25)
                pyperclip.copy('NMBB5280')
                pt.hotkey('ctrl', 'v')
                sleep(1)
            TC_LogonPassword, Error = Look_For("TC_LogonPassword.png")
            if Error:
                return True
            else:
                pt.moveTo(TC_LogonPassword)
                pt.click()
                sleep(0.25)
                pyperclip.copy('123America#NHLS')
                pt.hotkey('ctrl', 'v')
                sleep(1)
                pt.click()
                sleep(1)
            TC_LogonButton, Error = Look_For("TC_LogonButton.png")
            if Error:
                return True
            else:
                pt.moveTo(TC_LogonButton)
                pt.click()
            break
        if i > 100:
            TC_LockedWindow = pt.locateCenterOnScreen("TC_LockedWindow.png", confidence=.8)
            if TC_LockedWindow is not None:
                TC_LockedPassword, Error = Look_For("TC_LockedPassword.png")
                if Error:
                    return True
                else:
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
                send_message("TC_SysDown.txt")
                MessageToSend = open("MessageToSend.txt", "w")
                MessageToSend.write("TrakCare Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                                    + "Cell#:" + str(df.iloc[row, 0]))
                MessageToSend.close()
                send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
                # send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
                zero_reg()
                df.iloc[row, 24] = "5"
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
        pyperclip.copy('P9')
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
    i = 0
    while True:
        TC_DropDown = pt.locateCenterOnScreen("TC_DropDown.png", confidence=.8)
        if TC_DropDown is not None:
            break
        if i > 100:
            TC_Logout()
            incorrect_answer("Invalid_Fol_Num.txt")
            MessageToSend = open("MessageToSend.txt", "w")
            MessageToSend.write("Invalid Folder Number Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                                + "Cell#:" + str(df.iloc[row, 0]))
            MessageToSend.close()
            send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
            # send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
            return True
        i = i + 1
        sleep(1)

    #Collect Date of Birth
    TC_DOB = pt.locateCenterOnScreen("TC_DOB.png", confidence=.8)
    pt.moveTo(TC_DOB[0], TC_DOB[1]+40)
    sleep(0.5)
    pt.tripleClick()
    sleep(0.5)
    pt.hotkey('ctrl', 'c')
    DOB = pyperclip.paste()
    if '/' in DOB:
        DOBSplit = DOB.split('/')
        DOBFormat = DOBSplit[2].strip() + '-' + DOBSplit[1] + '-' + DOBSplit[0]
        df.iloc[row, 5] = DOBFormat
    else:
        TC_Logout()
        send_message("TC_DOB_Unavailable.txt")
        MessageToSend = open("MessageToSend.txt", "w")
        MessageToSend.write("No Date of Birth Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                            + "Cell#:" + str(df.iloc[row, 0]))
        MessageToSend.close()
        send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
        send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
        zero_reg()
        df.iloc[row, 24] = "5"
        return True

    #Collect Names
    TC_NameMarkerStart = pt.locateCenterOnScreen("TC_DropDown.png", confidence=.9)
    TC_NameMarkerEnd = pt.locateOnScreen("TC_NameMarkerEnd.png", confidence=.9)
    pt.moveTo(TC_NameMarkerStart[0]+30, TC_NameMarkerStart[1])
    pt.mouseDown()
    pt.dragRel(TC_NameMarkerEnd[0] - TC_NameMarkerStart[0] - 30, 0, duration=1.5)
    sleep(1)
    pt.mouseUp()
    pt.hotkey('ctrl', 'c')
    temp = pyperclip.paste()
    Personal = temp.splitlines()
    sleep(1)
    length = len(Personal)
    if length > 0 and all(i.isalpha() or i.isspace() for i in Personal[0].strip()):
        df.iloc[row, 7] = Personal[0].strip()
    else:
        df.iloc[row, 7] = "UNKNOWN"
    if length > 1 and all(i.isalpha() or i.isspace() for i in Personal[1].strip()):
        df.iloc[row, 6] = Personal[1].strip()
    else:
        df.iloc[row, 7] = "UNKNOWN"

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
    i = 0
    while True:
        TC_GXPU = pt.locateCenterOnScreen("TC_GXPU.png", confidence=.8)
        if TC_GXPU is not None:
            break
        if i > 20:
            TC_Logout()
            incorrect_answer("No_GXPU.txt")
            MessageToSend = open("MessageToSend.txt", "w")
            MessageToSend.write("No GXPU Test Registered Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                                + "Cell#:" + str(df.iloc[row, 0]))
            MessageToSend.close()
            send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
            # send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
            return True
        i = i + 1
        sleep(1)

    # GXPU Proccesed yet?
    pt.moveTo(TC_GXPU)
    pt.click()
    i = 0
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
    i = 0
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
            send_message("GXPU_Result_Undefined.txt")
            MessageToSend = open("MessageToSend.txt", "w")
            MessageToSend.write("GXPU Result Undefined Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                                + "Cell#:" + str(df.iloc[row, 0]))
            MessageToSend.close()
            send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
            send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
            zero_reg()
            df.iloc[row, 24] = "5"
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
    i = 0
    while True:
        TC_TestDate = pt.locateCenterOnScreen('TC_TestDate.png', confidence=.8)
        if TC_TestDate is not None:
            pt.moveTo(TC_TestDate)
            pt.moveRel(70, 0)
            sleep(0.5)
            pt.tripleClick()
            sleep(0.5)
            pt.hotkey('ctrl','c')
            TestDate = pyperclip.paste()
            TestdateSeperated = TestDate.split("/")
            TestdateFormated = TestdateSeperated[2].strip() + "-" + TestdateSeperated[1] + "-" + TestdateSeperated[0]
            df.iloc[row, 8] = TestdateFormated
            break
        if i > 20:
            TC_Logout()
            send_message("TestDate_Undefined.txt")
            MessageToSend = open("MessageToSend.txt", "w")
            MessageToSend.write("Test Date Unavailable Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                                + "Cell#:" + str(df.iloc[row, 0]))
            MessageToSend.close()
            send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
            send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
            zero_reg()
            df.iloc[row, 24] = "5"
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
        if i > 50:
            Error = True
            TC_Logout()
            send_message("TC_SysDown.txt")
            MessageToSend = open("MessageToSend.txt", "w")
            MessageToSend.write("TrakCare Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n"
                                + "Cell#:" + str(df.iloc[row, 0]))
            MessageToSend.close()
            send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
            #send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
            zero_reg()
            df.iloc[row, 24] = "5"
            break
        i = i + 1
        sleep(0.5)
    return Temp,Error

def TC_Logout():
    i = 0
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
            send_message_to("TC_URLMarker_NotFound.txt","+27 72 895 9231","WA_BvBNum.png")
            break
        i = i + 1
        sleep(1)
    return

def Look_For_GC(Image):
    global df
    global row
    i = 0
    Error = False
    while True:
        Temp = pt.locateCenterOnScreen(Image, confidence=.8)
        if Temp is not None:
            break
        if i > 50:
            Error = True
            #go to whatsapp
            Back_To_WA()
            send_message("System_Down.txt")
            MessageToSend = open("MessageToSend.txt", "w")
            MessageToSend.write("Google Calender Error" + "\n" + "Folder#: " + str(df.iloc[row, 4]) + "\n" + "Cell#:" + str(df.iloc[row, 0]))
            MessageToSend.close()
            send_message_to('MessageToSend.txt', "+27 72 895 9231","WA_BvBNum.png")
            #send_message_to('MessageToSend.txt', "+27 82 075 8484","WA_WorkNum.png")
            zero_reg()
            df.iloc[row, 24] = "5"
            break
        i = i + 1
        sleep(0.5)
    return Temp,Error

def Back_To_WA():
    pt.moveTo(150, 35)
    pt.click()
    sleep(1)
    return

def find_available_date():
    global df
    global row

    Calendar_Tab, Error = Look_For_GC("Calendar_Tab.png")
    if Error:
        return
    else:
        pt.moveTo(Calendar_Tab)
        pt.click()
        sleep(1)
        pt.moveRel(0, 300)
        sleep(0.25)
        pt.scroll(-400)
        sleep(0.25)

    PosTraceRange = 14
    NegWaitRange = 3
    test_date_obj = datetime.strptime(str(df.iloc[row, 8]), '%Y-%m-%d')
    now_date_obj = datetime.strptime(str(datetime.now().date()), '%Y-%m-%d')
    now_test_diff = int((now_date_obj - test_date_obj).days)

    # get profile none limits
    if str(df.iloc[row, 9]) == "1" or str(df.iloc[row, 9]) == "3": #pos and trace
        none_limit = (PosTraceRange - 1) - now_test_diff

    elif str(df.iloc[row, 9]) == "2" or str(df.iloc[row, 9]) == "4": #Neg or Waiting
        none_limit = (NegWaitRange - 1) - now_test_diff

    else:
        none_limit = 0

    if none_limit < 1:
        Back_To_WA()
        send_message('TestDate_OutOfRange.txt')
        df.iloc[row, 24] = "3"
        # if they were waiting for the results encourage restart if postive
        if str(df.iloc[row, 9]) == "4":
            Back_To_WA()
            send_message('Waiting_Restart.txt')
            df.iloc[row, 24] = "3"
        return

    #reset options if you ask for more dates *beyond* your limit
    if int(str(df.iloc[row, 10])) > none_limit:
        df.iloc[row, 10] = 1

    while int(str(df.iloc[row, 10])) <= none_limit:
        option_date_obj = now_date_obj + timedelta(days=int(str(df.iloc[row, 10])))
        if option_date_obj.weekday() == 5:
            df.iloc[row, 10] = int(str(df.iloc[row, 10])) + 2
        elif option_date_obj.weekday() == 6:
            df.iloc[row, 10] = int(str(df.iloc[row, 10])) + 1
        else:
            #Using number of nones jog to correct day in calendar.
            Today, Error = Look_For_GC("Today.png")
            if Error:
                return
            else:
                pt.moveTo(Today)
                pt.click()
                sleep(1)

            Arrows, Error = Look_For_GC("Arrows.png")
            if Error:
                return
            else:
                x = Arrows[0]
                y = Arrows[1]
            n = 0
            while n < int(str(df.iloc[row, 10])):
                pt.moveTo(x + 12, y, duration=.05)
                pt.click()
                sleep(1)
                n = n + 1

            OptionAvialable = "0"
            filelink = open("Temp_Options.txt", "w")
            filelink.write("*English:*" + "\n" + "Please select a time slot that suits you." + "\n" + "\n")
            filelink.write("*Xhosa:*" + "\n" + "Nceda ukhethe ixesha elizokulungela." + "\n" + "\n")
            filelink.write("*Afrikaans:*" + "\n" + "Kies n tydgleuf wat u pas." + "\n"+ "\n")
            filelink.write("Date: *" + str(option_date_obj.date()) + "*" + "\n")
            filelink.write("*0*" + " - _More Options_" + "\n")
            if pt.pixelMatchesColor(int(x + 100), int(y + 225), (255, 255, 255), tolerance=5):
                #SAVE 8:30 AS AN OPTION
                filelink.write("*1*" + " - @ 8:15" + "\n")
                df.iloc[row, 12] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*1*" + " - _Booked Out_" + "\n")
                df.iloc[row, 12] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 257), (255, 255, 255), tolerance=5):
                # SAVE 9:30 AS AN OPTION
                filelink.write("*2*" + " - @ 9:00" + "\n")
                df.iloc[row, 13] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*2*" + " - _Booked Out_" + "\n")
                df.iloc[row, 13] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 289), (255, 255, 255), tolerance=5):
                # SAVE 10:30 AS AN OPTIOn
                filelink.write("*3*" + " - @ 9:45" + "\n")
                df.iloc[row, 14] = "1"
            else:
                filelink.write("*3*" + " - _Booked Out_" + "\n")
                df.iloc[row, 14] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 321), (255, 255, 255), tolerance=5):
                # SAVE 11:30 AS AN OPTION
                filelink.write("*4*" + " - @ 10:30" + "\n")
                df.iloc[row, 15] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*4*" + " - _Booked Out_" + "\n")
                df.iloc[row, 15] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 353), (255, 255, 255), tolerance=5):
                # SAVE 13:30 AS AN OPTION
                filelink.write("*5*" + " - @ 11:15" + "\n")
                df.iloc[row, 16] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*5*" + " - _Booked Out_" + "\n")
                df.iloc[row, 16] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 430), (255, 255, 255), tolerance=5):
                # SAVE 14:30 AS AN OPTION
                filelink.write("*6*" + " - @ 13:00" + "\n")
                df.iloc[row, 17] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*6*" + " - _Booked Out_" + "\n")
                df.iloc[row, 17] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 462), (255, 255, 255), tolerance=5):
                # SAVE 15:30 AS AN OPTION
                filelink.write("*7*" + " - @ 13:45" + "\n")
                df.iloc[row, 18] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*7*" + " - _Booked Out_" + "\n")
                df.iloc[row, 18] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 494), (255, 255, 255), tolerance=5):
                # SAVE 14:30 AS AN OPTION
                filelink.write("*8*" + " - @ 14:30" + "\n")
                df.iloc[row, 19] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*8*" + " - _Booked Out_" + "\n")
                df.iloc[row, 19] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 526), (255, 255, 255), tolerance=5):
                # SAVE 15:30 AS AN OPTION
                filelink.write("*9*" + " - @ 15:15" + "\n")
                df.iloc[row, 20] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*9*" + " - _Booked Out_" + "\n")
                df.iloc[row, 20] = "0"
            if pt.pixelMatchesColor(int(x + 100), int(y + 558), (255, 255, 255), tolerance=5):
                # SAVE 15:30 AS AN OPTION
                filelink.write("*10*" + " - @ 16:00" + "\n")
                df.iloc[row, 21] = "1"
                OptionAvialable = "1"
            else:
                filelink.write("*10*" + " - _Booked Out_" + "\n")
                df.iloc[row, 21] = "0"

            filelink.write("*11*" + " - _Restart Options_" + "\n")
            filelink.write("*12*" + " - _Cancel Appointment_" + "\n")
            filelink.close()

            if OptionAvialable == "0":
                #no available options
                df.iloc[row, 10] = int(str(df.iloc[row, 10])) + 1
            else:
                #send available options
                Back_To_WA()
                send_message('Temp_Options.txt')
                #OPtion date
                df.iloc[row, 22] = now_date_obj
                return
    Back_To_WA()
    send_message('S_None_Available.txt')
    df.iloc[row, 24] = "3"
    if str(df.iloc[row, 9]) == "3":
        send_message('Waiting_Restart.txt')
    return

def check_selection(selected):
    global df
    global row

    opt_date_obj = datetime.strptime(str(df.iloc[row, 22]), '%Y-%m-%d')
    now_date_obj = datetime.strptime(str(datetime.now().date()), '%Y-%m-%d')
    opt_now_diff = int((opt_date_obj - now_date_obj).days)

    if int(opt_now_diff) <= 0:
        Back_To_WA()
        send_message('S_Date_Passed.txt',)
        df.iloc[row, 11] = "1"
        find_available_date()
        return

    Today = pt.locateCenterOnScreen("Today.png", confidence=.7)
    pt.moveTo(Today)
    pt.click()
    sleep(1)

    Arrows = pt.locateCenterOnScreen("Arrows.png", confidence=.7)
    x = Arrows[0]
    y = Arrows[1]
    n = 0
    while n < opt_now_diff:
        pt.moveTo(x + 12, y, duration=.05)
        pt.click()
        sleep(1)
        n = n + 1

    if selected == 1:
        if pt.pixelMatchesColor(int(x + 100), int(y + 270), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 2:
        if pt.pixelMatchesColor(int(x + 100), int(y + 294), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 3:
        if pt.pixelMatchesColor(int(x + 100), int(y + 318), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 4:
        if pt.pixelMatchesColor(int(x + 100), int(y + 342), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 5:
        if pt.pixelMatchesColor(int(x + 100), int(y + 390), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 6:
        if pt.pixelMatchesColor(int(x + 100), int(y + 414), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 7:
        if pt.pixelMatchesColor(int(x + 100), int(y + 438), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 8:
        if pt.pixelMatchesColor(int(x + 100), int(y + 438), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 9:
        if pt.pixelMatchesColor(int(x + 100), int(y + 438), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    elif selected == 10:
        if pt.pixelMatchesColor(int(x + 100), int(y + 438), (255, 255, 255), tolerance=5):
            BookingPrep('8:15')
        else:
            Back_To_WA()
            send_message('S_No_longer_Ava.txt')
            df.iloc[row, 8] = "1"
            find_available_date()
    else:
        Back_To_WA()
        send_message('S_Not_Option.txt')
        df.iloc[row, 8] = "1"
        find_available_date()
    return

def BookingPrep(Time):
    global df
    global row
    GC_Booking, Error = Look_For_GC("GC_Booking.png")
    if Error:
        return
    else:
        pt.moveTo(GC_Booking)
        pt.click()
        sleep(1)

    GC_TimeMarker, Error = Look_For_GC("GC_TimeMarker.png")
    if Error:
        return
    else:
        pt.moveTo(GC_TimeMarker)
        pt.moveRel(0, -45)
        pt.click()
        pyperclip.copy(str(df.iloc[row, 9]) + " " + str(df.iloc[row, 0]) + " " + str(df.iloc[row, 6]) + " " + str(df.iloc[row, 7]) + " " + str(df.iloc[row, 4]))
        pt.hotkey('ctrl', 'v')
        sleep(0.5)
        pt.moveTo(GC_TimeMarker)
        pt.moveRel(-5, 40)
        pt.click()
        sleep(1)
        pyperclip.copy(Time)
        pt.hotkey('ctrl', 'v')
        sleep(0.5)
        pt.hotkey('enter')
        sleep(0.5)

    save, Error = Look_For_GC("save.png")
    if Error:
        return
    else:
        pt.moveTo(save)
        pt.click()
        sleep(1)

    Back_To_WA()
    filelink = open("Slot_Ava_Booked.txt", "w")
    filelink.write("*English:*" + "\n" + "Your appointment has been booked successfully" + "\n"+ "\n")
    filelink.write("*Xhosa:*" + "\n" + "Uzokwazi ukubayinxalenye yophando ixesha sele ulibekelwe." + "\n" + "\n")
    filelink.write("*Afrikaans:*" + "\n" + "U afspraak is suksesvol bespreek" + "\n" + "\n")
    filelink.write("Date: *" + str(df.iloc[row, 22]) + "*" + "\n" + "Time: " + Time)
    filelink.close()
    send_message('Slot_Ava_Booked.txt')
    df.iloc[row, 24] = "1"
    df.iloc[row, 23] = Time
    send_map()
    send_message('S_Ticket_Warning.txt')
    return

def send_map():
    txt = open('S_Directions.txt' 'r')
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
    Ticket_Example = pt.locateCenterOnScreen("map.png", confidence=.7)
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

check_for_new_message()


