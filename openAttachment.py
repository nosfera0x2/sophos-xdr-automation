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

def returnOutlookPID():
    process_name = "OUTLOOK.EXE"
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            outlookPID = proc.pid
            return (outlookPID)

outlook = returnOutlookPID()

def returnMicrosoftPID():
    process_name = "WINWORD.EXE"
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            wordPID = proc.pid
            return (wordPID)

microsoftWord = returnMicrosoftPID()


def checkIfOutlookRunning(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;

def killOutlook():
    ti = 0
    name = 'OUTLOOK.EXE'
    f = wmi.WMI()

    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()
            ti +=1
    if ti == 0:
        print('Process not found!')


def openEmailAttachment():
    #start the Outlook application
   
    app = Application(backend="uia").start(r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK')

    time.sleep(10)
    #Main application is presented as 'Outlook Today - Outlook'
    #Define the main dialogue
    main_dlg = app['Outlook Today - Outlook']
    #Define the bth2 tree item in the left pane of outlook
    bth2_tree = main_dlg.child_window(title="BTH2", control_type="TreeItem")
    #Define the '2' subfolder in the bth2 folder.
    treeItem_two = bth2_tree.child_window(title="2", control_type="TreeItem")
    #Presents the email available in the 2 subfolder
    treeItem_two.click_input(double=True)
    time.sleep(5)
    #The title of the outlook application has now changed after selected the '2' subfolder
    #Redefine the main dialogue, as 'second dialogue'
    second_dlg = app['2 - Sophos - Outlook']
    time.sleep(5)
    #Define the email, and open it. 
    email = second_dlg.child_window(title="With Attachments, Subject Your flight has been successfully booked!, Received Mon 1/31, Size 136 KB, Flag Status Unflagged, ", control_type="DataItem")
    email.click_input(double=True)
    time.sleep(5)
     
    #The email is now opened, and defined as an interactive dialogue 
    #The attachment is identified and opened. 
    email_dlg = app['Your flight has been successfully booked! - Message (HTML) ']
    attachment = email_dlg.child_window(title="Attachments", control_type="Pane")
    attachment.Button25.click_input(double=True)

def openEmailOutlookRunning():
    app = Application(backend="uia").connect(process=outlook, visible_only=False)
    time.sleep(5)
    #need to make window visible for wrapper object
    print("I've detected outlook running, here are the windows:")
    print(app.windows())
    nameOfWindow = 'Outlook Today - Outlook'
    print("I'm now using win32gui to make the window visible")
    window= win32gui.FindWindow(None,nameOfWindow)
    win32gui.ShowWindow(window, win32con.SW_SHOWMAXIMIZED)
    time.sleep(5)
    #Main application is presented as 'Outlook Today - Outlook'
    #Define the main dialogue
    main_dlg = app['Outlook Today - Outlook']
    main_dlg.set_focus()
    time.sleep(5)
    #Define the bth2 tree item in the left pane of outlook
    bth2_tree = main_dlg.child_window(title="BTH2", control_type="TreeItem")
    #Define the '2' subfolder in the bth2 folder.
    treeItem_two = bth2_tree.child_window(title="2", control_type="TreeItem")
    #Presents the email available in the 2 subfolder
    treeItem_two.click_input(double=True)
    time.sleep(5)
    #The title of the outlook application has now changed after selected the '2' subfolder
    #Redefine the main dialogue, as 'second dialogue'
    second_dlg = app['2 - Sophos - Outlook']
    time.sleep(5)
    #Define the email, and open it. 
    email = second_dlg.child_window(title="With Attachments, Subject Your flight has been successfully booked!, Received Mon 1/31, Size 136 KB, Flag Status Unflagged, ", control_type="DataItem")
    email.click_input(double=True)
    time.sleep(5)
     
    #The email is now opened, and defined as an interactive dialogue 
    #The attachment is identified and opened. 
    email_dlg = app['Your flight has been successfully booked! - Message (HTML) ']
    attachment = email_dlg.child_window(title="Attachments", control_type="Pane")
    attachment.Button25.click_input(double=True)

if __name__ == '__main__':
   if (checkIfOutlookRunning('OUTLOOK.EXE') == True):
        openEmailOutlookRunning()
   else:
        openEmailAttachment()
