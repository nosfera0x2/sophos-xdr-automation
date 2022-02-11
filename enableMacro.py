# ----------------------------------------------------------------------------
# Name:         Open Email Attachment Function
# Description:  Python Script to Automate XDR Threat Intel
# Author:       Spencer Brown
# URL:          
# Date:         02/09/2022
# ----------------------------------------------------------------------------

from pywinauto import Application, Desktop
import time
import psutil
import os 
import wmi 
import win32gui
import win32con
from extractTitle import extractTitle


def returnMicrosoftPID():
    process_name = "WINWORD.EXE"
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            wordPID = proc.pid
            return (wordPID)

microsoftWord = returnMicrosoftPID()

def openMacro():
    app = Application(backend="uia").connect(title_re='.*DeltaFlightItinerary')
    #print(app.windows())
    window = app.windows()
    wordWindow = str(extractTitle(window[0]))
    print(wordWindow)
    type(wordWindow)
    print("I'm looking for the window..." + wordWindow)
    guiWindow = win32gui.FindWindow(None,wordWindow)
    win32gui.ShowWindow(guiWindow,win32con.SW_SHOWMAXIMIZED)
    #time.sleep(5)
    print(app.windows())
    main_dlg = app[wordWindow]
    main_dlg.set_focus()
    main_dlg.print_control_identifiers()
    main_dlg.EnableEditingButton.click_input() #enable editing, pivots to enable content
    time.sleep(5)
    #print('this is the second set of identifiers:')
    #main_dlg.print_control_identifiers()
    main_dlg.EnableContentButton.click_input() #enable content button
    
openMacro()
