from pywinauto import Application, Desktop
import time
import os 
import psutil
import io
import sys
import re

def setUpOutlook():
    app=Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE')
    #currentWindows=app.windows()
    time.sleep(5)
    dlg=app['Welcome to Microsoft Outlook 2016']
    dlg.Next.click_input()
    dlg=app['Opening - Microsoft Outlook']    
    dlg.NoRadioButton.click_input()
    dlg.Next.click_input()
    dlg.CheckBox.click()
    dlg.Finish.click_input()
    time.sleep(30)
    dlg=app['Accounts']
    dlg.CloseButton.click_input()
    time.sleep(3)
    dlg=app['Outlook Today - Outlook']
    time.sleep(5)
    dlg.AcceptButton.click_input()
    time.sleep(10)
    os.system('cmd /k "Powershell.exe -ExecutionPolicy Unrestricted -file C:\\threat\\pstfile.ps1"')
    time.sleep(2)
    os.system('cmd /k "touch c:\threat\outlookIsSetup"')

def returnPID(process):
    process_name = process
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            processID = proc.pid
            return (processID)
            
def processRunCheck(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;
            
def killProcess(process):
    ti = 0
    name = process
    print('The process I am seeking to kill is: ' + str(name))
    for proc in psutil.process_iter():
    #check whether the process name matches
        if proc.name() == str(name):
            proc.kill()
            print('I have found, and killed ' + str(name))
        ti += 1
        if ti == 0:
            print('There are no running instances of ' + str(name))

outlook = returnPID("OUTLOOK.EXE")

def startEmailAttachment():
    
    app = Application(backend="uia").connect(title_re='Outlook Today - Outlook')
    time.sleep(5) #Need time for the activate office dialogue to appear. 
    accountsDLG = app['Accounts']
    accountsDLG.CloseButton.click_input()
    mainDLG= app['Outlook Today - Outlook'] #Main application is presented as 'Outlook Today - Outlook'
    mainDLG.sophosTreeItem.click_input(double=True)#Expand the Sophos profile tree
    sophosDLG = app['Sophos - Outlook']
    bthDLG = sophosDLG['BTH2TreeItem']
    bthDLG.click_input(double=True)#expand the BTH2 menu tree
    bthDLG = app['BTH2 - Sophos - Outlook']
    bthDLG['2TreeItem'].click_input()
    time.sleep(5)
    twoDLG = app['2 - Sophos - Outlook'] #New window name
    time.sleep(5)
    twoDLG['DeltaFlightItinerary.docm85 KB1 of 1 attachmentsUse alt + down arrow to open the options menu'].click_input(double=True) #opens the attachment


def enableMacro():
    app = Application(backend="uia").connect(title_re='.*DeltaFlightItinerary*')
    time.sleep(5)
    mainDLG = app.window(title_re="DeltaFlightItinerary*")
    mainDLG['Enter a product key instead'].click_input()
    productKeyDLG = mainDLG['Enter your product keyPane']
    productKeyDLG.CloseButton.click_input()
    mainDLG.EnableEditingButton.click_input() #Enable editing, pivots to enable content
    time.sleep(3)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)


if __name__ == '__main__':
    if (os.path.isfile(r"C:\threat\outlookIsSetup") == False):
       setUpOutlook()
    else:
        killProcess('WINWORD.EXE')

        if (processRunCheck('outlook.exe') == False):
            setUpOutlook()
            time.sleep(10)
            startEmailAttachment()
            time.sleep(10)
            enableMacro()
        
        else:
            print('Outlook is currently running.')
            print('I am now going to kill the Outlook Process.')
            #killProcess('OUTLOOK.EXE')
            time.sleep(5)
            #print('Outlook will now start.')
            startEmailAttachment()
            time.sleep(5)
            enableMacro()
