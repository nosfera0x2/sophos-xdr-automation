import pywinauto
import time 
import sys

outlook_start = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE'
outlook_connect = "Outlook Today - Outlook"


class pywinApp:

    def __init__(self,start=None,connect=None):
        self.start = start
        self.connect = connect

    def startapp(self):
        app = pywinauto.Application(backend="uia").start(self.start)
        return app

    def connectapp(self):  
        app = pywinauto.Application(backend="uia").connect(title_re=self.connect)
        return app
    

microsoft_outlook = pywinApp(start=outlook_start,connect=outlook_connect)

outlook_app = microsoft_outlook.connectapp()
outlook=outlook_app['Outlook Today - Outlook']
filepath = "C:\\Windows\\Temp\\test.txt"
the_wizard = 'Microsoft Office Activation Wizard'


def check_window(filepath,string,dialogue_object):
    temp = sys.stdout
    sys.stdout = open(filepath,'w')
    print(dialogue_object.children())
    sys.stdout = temp
    with open(filepath, 'r') as file:
        content = file.read()
        if string in content:
            print('Correct Dialogue Found!')
            return True
        else:
            print('Dialogue does not exist')
            return False
        
wizard = outlook['Microsoft Office Activation Wizard']
if check_window(filepath,the_wizard,outlook) == True:
    print("I will now close the window!")
    wizard.child_window(title="Close", control_type="Button").click_input()
else:
    print("This didn't work!")
