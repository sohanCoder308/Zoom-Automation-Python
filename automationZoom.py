import os
import pandas as pd
from pandas.core.indexes.base import Index
import pyautogui
import time
from datetime import datetime

def signIn(meeting_id,password):

    #Open's Zoom Application from the specified location
    os.startfile("C:/Users/vikas/AppData/Roaming/Zoom/bin/Zoom.exe")
    time.sleep(3)

    #Click's join button
    joinbtn=pyautogui.locateAllOnScreen("joinameeting.png")
    for jBtn in joinbtn:
        pyautogui.moveTo(jBtn)
        pyautogui.click()
        time.sleep(1)

    #Type the meeting id
    meetingidbtn=pyautogui.locateCenterOnScreen("meetingid.png")
    pyautogui.moveTo(meetingidbtn)
    pyautogui.write(meeting_id)
    time.sleep(3)

    #To turn of video and audio
    mediaBtn=pyautogui.locateAllOnScreen("media.png")
    for btn in mediaBtn:
        pyautogui.moveTo(btn)
        pyautogui.click()
        time.sleep(1)

    #To join
    join=pyautogui.locateAllOnScreen("join.png")
    for joinB in join:
        pyautogui.moveTo(joinB)
        pyautogui.click()
        time.sleep(2)

    #Enter's passcode to join meeting
    passcode=pyautogui.locateCenterOnScreen("meetingPasscode.png")
    pyautogui.moveTo(passcode)
    pyautogui.write(password)
    time.sleep(1)

    #Click's on join button
    joinmeeting=pyautogui.locateAllOnScreen("joinmeeting.png")
    for joinMe in joinmeeting:
        pyautogui.moveTo(joinMe)
        pyautogui.click()
        time.sleep(1)


df = pd.read_excel('timings.xlsx')

while True:
    # To get the current time
    now = datetime.now().strftime("%H:%M")
    if now in str(df['Timings']):
        mylist=df["Timings"]
        mylist=[i.strftime("%H:%M") for i in mylist]
        c= [i for i in range(len(mylist)) if mylist[i]==now]
        row = df.loc[c] 
        print(row)
        meeting_id = str(row.iloc[0,1])  
        password = str(row.iloc[0,2])  
        time.sleep(5)
        signIn(meeting_id, password)
        time.sleep(2)
        print('signed in')
        break