#V3 - TrakCare integration

import pyautogui as pt
from time import sleep
from datetime import datetime
from datetime import timedelta
import pyperclip
import pandas
sleep(10)

def check_for_new_message():
    global df
    df = pandas.read_csv("Patient_Log.txt")
    # gui check for new message or green dots
    i=0
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
    Close = pt.locateCenterOnScreen("Close.png", confidence=.8)
    pt.moveTo(Close)
    pt.click()
    sleep(0.25)
    df.to_csv("auto_backup.txt", index=False)
    check_phone_num(pyperclip.paste())
    return

def check_phone_num(phone_num):
    global df
    global row
    i = len(df)
    match = False
    while i > 0: #Find matching numbers
        if str(df.iloc[i-1, 0]) == str(phone_num):
            match = True
            row = i - 1 # used for continued screening and for session indexing
            break # only latest of interest, save time
        i = i - 1

    if match == True:
        get_message()
        new_message = pyperclip.paste().lower()
        if str(df.iloc[row, 20]) == "1" or str(df.iloc[row, 20]) == "2": #Closed Successfully
            if new_message == "restart" or new_message == "restart." or new_message == "restart ":
                row = len(df)
                New_row = {"Cell": str(phone_num),
                           "Session": str(int(df.iloc[row, 1])+1),
                           "Attempts": "0",
                           "Consent": "0",
                           "Ticket": "0",
                           "Folder": "0",
                           "Date": "0",
                           "Status": "0",
                           "Offset": "1",
                           "8:15": "0", "9:00": "0", "9:45": "0","10:30": "0", "11:15": "0", "13:00": "0", "13:45": "0","14:30": "0", "15:15": "0", "16:00": "0",
                           "Appointment": "0",
                           "Outcome": "0"}
                df = df.append(New_row, ignore_index=True)
                txt = open('S0_Introduction.txt', 'r')
                temp = txt.read()
                txt.close()
                pyperclip.copy(temp)
                send_message()
                txt = open('Q1_Consent.txt', 'r')
                temp = txt.read()
                txt.close()
                pyperclip.copy(temp)
                send_message()
            else:
                txt = open('Restart.txt', 'r')
                temp = txt.read()
                txt.close()
                pyperclip.copy(temp)
                send_message()
        elif str(df.iloc[row, 20]) == "3":  # Closed and Unsuccessful
            if new_message == "restart" or new_message == "restart." or new_message == "restart ":
                df.iloc[row, 20] = "0"
                txt = open('S0_Introduction.txt', 'r')
                temp = txt.read()
                txt.close()
                pyperclip.copy(temp)
                send_message()
                txt = open('Q1_Consent.txt', 'r')
                temp = txt.read()
                txt.close()
                pyperclip.copy(temp)
                send_message()
            else:
                txt = open('Restart.txt', 'r')
                temp = txt.read()
                txt.close()
                pyperclip.copy(temp)
                send_message()
        elif str(df.iloc[row, 20]) == "4":  # Closed and Blocked
            txt = open('Blocked.txt', 'r')
            temp = txt.read()
            txt.close()
            pyperclip.copy(temp)
            send_message()
        else: # str(df.iloc[row, 20]) == "0":  #Open and in progress
            pass
    else: # match == False:
        row = len(df)
        New_row = {"Cell": str(phone_num),
                   "Session": "0",
                   "Attempts": "0",
                   "Consent": "0",
                   "Ticket": "0",
                   "Folder": "0",
                   "Date": "0",
                   "Status": "0",
                   "Offset": "1",
                   "8:15": "0", "9:00": "0", "9:45": "0", "10:30": "0", "11:15": "0", "13:00": "0", "13:45": "0",
                   "14:30": "0", "15:15": "0", "16:00": "0",
                   "Appointment": "0",
                   "Outcome": "0"}
        df = df.append(New_row, ignore_index=True)
        df.iloc[row, 20] = "0"
        txt = open('S0_Introduction.txt', 'r')
        temp = txt.read()
        txt.close()
        pyperclip.copy(temp)
        send_message()
        txt = open('Q1_Consent.txt', 'r')
        temp = txt.read()
        txt.close()
        pyperclip.copy(temp)
        send_message()
    df.to_csv("Patient_Log.txt", index=False)
    return

def store_and_ask():
    global df
    global row
    # closed - check number outcome and reopen,new row etc in get and check phone number
    get_message()
    new_message = pyperclip.paste().lower()

    if new_message == "pause system":
        temp = "System paused for 100 seconds"
        pyperclip.copy(temp)
        send_message()
        sleep(100)
    elif str(df.iloc[row, 20]) == "1" or str(df.iloc[row, 20]) == "2" or str(df.iloc[row, 20]) == "3" or str(df.iloc[row, 20]) == "4": #session closed
        get_phone_num()
    elif new_message == "":
        txt = open('S_Message_Del.txt', 'r')
        temp = txt.read()
        txt.close()
        pyperclip.copy(temp)
        send_message()
    elif new_message == "stop":
        zero_reg()
        txt = open('SG_Stop.txt', 'r')
        temp = txt.read()
        txt.close()
        pyperclip.copy(temp)
        send_message()
        txt = open('Restart.txt', 'r')
        temp = txt.read()
        txt.close()
        pyperclip.copy(temp)
        send_message()


def zero_reg():
    global df
    global row
    i = 20
    while i > 2:
        df.iloc[row, i] = "0"
        i = i - 1
    df.iloc[row, 8] = "1" #offset set to 1
    df.iloc[row, 20] = "3"  # Outcome to Unsuccsesful
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

def send_message():
    sleep(0.5)
    Paperclip = pt.locateOnScreen("Paperclip.png", confidence=.6)
    pt.moveTo(Paperclip)
    pt.moveRel(130, 0)
    pt.click()
    sleep(0.25)
    pt.hotkey('ctrl', 'v')
    pt.typewrite("\n", interval=0.01)
    return

check_for_new_message()


