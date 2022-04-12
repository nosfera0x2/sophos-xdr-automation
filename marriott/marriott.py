from pywinauto import Application, Desktop
import time
import os 
import psutil
import io
import sys
import re

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

def processRunCheck(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;

def marriottEmailAttachment():
    
    app = Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK')
    time.sleep(5) #Need time for the activate office dialogue to appear. 
    accountsDLG = app['Accounts']
    time.sleep(3)
    accountsDLG.CloseButton.click_input()
    mainDLG= app['Outlook Today - Outlook'] #Main application is presented as 'Outlook Today - Outlook'
    mainDLG.sophosTreeItem.click_input(double=True)#Expand the Sophos profile tree
    sophosDLG = app['Sophos - Outlook']
    tmDLG=sophosDLG['z.TM AUTO'].click_input(double=True)
    print(app.windows())
    marriottDLG=app['z.TM AUTO - Sophos - Outlook']
    time.sleep(3)
    marriottDLG.print_control_identifiers()
    time.sleep(5)
    marriottDLG['BonusPointsOffer.docm776 KB1 of 1 attachmentsUse alt + down arrow to open the options menu'].click_input(double=True) #opens the attachment
	
#TO ENABLE THE MACRO
def marriottMacro():
    app = Application(backend="uia").connect(title_re='.*BonusPointsOffer*')
    time.sleep(5)
    mainDLG = app.window(title_re="BonusPointsOffer*")
    mainDLG['Enter a product key instead'].click_input()
    productKeyDLG = mainDLG['Enter your product keyPane']
    productKeyDLG.CloseButton.click_input()
    mainDLG.EnableEditingButton.click_input() #Enable editing, pivots to enable content
    time.sleep(3)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)

if __name__ == '__main__':
    if (os.path.isfile(r"C:\threat\outlookIsSetup") == False):
       print('Outlook is not properly set up')
    else:
        killProcess('WINWORD.EXE')

        if (processRunCheck('outlook.exe') == False):
          
            time.sleep(10)
            marriottEmailAttachment()
            time.sleep(10)
            marriottMacro()
        
        else:
            print('Outlook is currently running.')
            print('I am now going to kill the Outlook Process.')
            killProcess('OUTLOOK.EXE')
            time.sleep(5)
            print('Outlook will now start.')
            marriottEmailAttachment()
            time.sleep(5)
            marriottMacro()
