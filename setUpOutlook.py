from calendar import c
import psutil
import win32gui
import win32con
from pywinauto import Application, Desktop
import time
import os


def setUpOutlook():
    app=Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE')
    #currentWindows=app.windows()
    time.sleep(2)
    dlg=app['Welcome to Microsoft Outlook 2016']
    dlg.Next.click_input()
    dlg=app['Opening - Microsoft Outlook']    
    dlg.NoRadioButton.click_input()
    dlg.Next.click_input()
    dlg.CheckBox.click()
    dlg.Finish.click_input()
    time.sleep(5)
    dlg=app['Accounts']
    dlg.CloseButton.click_input()
    time.sleep(3)
    dlg=app['Outlook Today - Outlook']
    time.sleep(5)
    dlg.AcceptButton.click_input()
    time.sleep(5)
    os.system('cmd /k "Powershell.exe -ExecutionPolicy Unrestricted -file C:\\Users\\spencer\\ProvisioningFiles\\scripts\\pstfile.ps1"')
setUpOutlook() 