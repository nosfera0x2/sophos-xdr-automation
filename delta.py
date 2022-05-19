import pywinauto
import time 
import sys

outlook_path = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE'
outlook_connect = "Outlook Today - Outlook"


class pywinApp:

    def __init__(self,path):
        self.path = path
       

    def startapp(self):
        app = pywinauto.Application(backend="uia").start(self.path,class_name="Outlook")
        return app

    def connectapp(self):  
        app = pywinauto.Application(backend="uia").connect(path=self.path,visible_only=False,class_name="Outlook")
        return app
    

microsoft_outlook = pywinApp(outlook_path)

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
            print(f'Dialogue: {string} in {dialogue_object} located.')
            return True
        else:
            print(f'Dialogue: {string} in {dialogue_object} not found.')
            dialogue_object.print_control_identifiers()
            return False

def check_id(filepath,string,dialogue_object):
    temp = sys.stdout
    sys.stdout = open(filepath,'w',encoding="utf8")
    print(dialogue_object.print_control_identifiers())
    sys.stdout = temp
    with open(filepath, 'r') as file:
        content = file.read()
        if string in content:
            print(f'Dialogue: {string} in {dialogue_object} located.')
            return True
        else:
            print(f'Dialogue: {string} in {dialogue_object} not found.')
            return False
        
wizard = outlook['Microsoft Office Activation Wizard']
if check_window(filepath,the_wizard,outlook) == True:
    print("I will now close the window!")
    wizard.child_window(title="Close", control_type="Button").click_input()
else:
    print("This didn't work!")

outlook.child_window(title="Sophos", control_type="TreeItem").click_input(double=True)
sophos = outlook.child_window(title="Sophos", control_type="TreeItem")


def revert_tree():
    global filepath
    global sophos
    string = 'child_window(title="Inbox", control_type="TreeItem")'
    dlg = outlook_app.window(title_re=".*Sophos*.")
    if check_id(filepath,string,dlg) == True:
        dlg.child_window(title="Inbox", control_type="TreeItem").click_input(double=True)
        print("Tree is reset")
    else:
        print("oops")
        pass

revert_tree()



