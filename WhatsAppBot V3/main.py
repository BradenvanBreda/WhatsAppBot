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
               "Date": "0",
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
    elif str(df.iloc[row, 19]) == "0":
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
            df.iloc[row, 19] = "0"
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
            Fol_num_Str = str(''.join(Fol_Num_Char))
            if Fol_num_Str.isnumeric():
                # check it
            else:
                incorrect_answer('Invalid_Fol_Num.txt')
            sleep(100)

#Too be continued
    elif str(df.iloc[row, 19]) == "1" or str(df.iloc[row, 19]) == "2":  # Closed Successfully
        if new_message == "restart" or new_message == "restart." or new_message == "restart ":
            Initiate_NL(str(df.iloc[row, 0])) #send current phone number
        else:
            send_message('Restart.txt')
    elif str(df.iloc[row, 19]) == "3":  # Closed and Unsuccessful
        if new_message == "restart" or new_message == "restart." or new_message == "restart ":
            send_message('S0_Introduction.txt')
            send_message('Q1_Consent.txt')
            zero_reg()
            df.iloc[row, 19] = "0"
        else:
            send_message('Restart.txt')
    elif str(df.iloc[row, 19]) == "4":  # Closed and Blocked
        send_message('Blocked.txt')
    else:
        print("no If entered")
    df.to_csv("Patient_Log.txt", index=False)
    return

def zero_reg():
    global df
    global row
    i = 19
    while i > 1: #does not zero attempts
        df.iloc[row, i] = "0"
        i = i - 1
    df.iloc[row, 7] = "1" #offset set to 1
    df.iloc[row, 19] = "3" #default Closed Unsuccessful
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
        df.iloc[row, 19] = "4"
        send_message(reason)
        send_message('Blocked.txt')
    else:
        send_message(reason)
        send_message('Restart.txt')
        zero_reg()
    return

check_for_new_message()


