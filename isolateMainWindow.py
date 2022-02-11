from pywinauto import Application, Desktop
import time
import psutil
import os 
import wmi 
import win32gui
import win32con 
import re 

def checkIfOutlookRunning(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;

def returnOutlookPID():
    process_name = "OUTLOOK.EXE"
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            outlookPID = proc.pid
            return (outlookPID)

outlook = returnOutlookPID()

def openEmailOutlookRunning():
    app = Application(backend="uia").connect(process=outlook, visible_only=False)
    #time.sleep(5)
    #need to make window visible for wrapper object
    #print("I've detected outlook running, here are the windows:")
    currentWindow= app.windows()
    #print(currentWindow[0])
    outlookString = str(currentWindow[0])
    #print(type(outlookString))
    #print(re.findall(r"'(.*?)'",outlookString))
    #targetString = str(re.findall(r"'(.*?)'",outlookString))
    print(outlookString)
    pattern = r"'(.*?[^\\])'"
    targetString = str(re.findall(pattern,outlookString))
    targetString1 = targetString.replace("[","")
    targetString2 = targetString1.replace("]","")
    targetStringFinal = targetString2.replace("'","")
    print("The name of the window you are on is:" + targetStringFinal)
    window = win32gui.FindWindow(None,targetStringFinal)
    win32gui.ShowWindow(window, win32con.SW_SHOWMAXIMIZED)
    time.sleep(5)
    #print("Here are the current windows:" + str(app.windows()))
    #print("I should be able to connect by app[targetStringFinal]")
    print("Here are the windows I now have available:" + str(app.windows()))
    test_dlg = app[targetStringFinal]
    test_dlg.set_focus()
    time.sleep(3)
    print("This is where I will be printing the control identifiers:")
    #test_dlg.print_control_identifiers()
    sophos_pane = test_dlg['Pane17']
    sophos_folder = sophos_pane.child_window(title="Mail Folders", control_type="Tree")
    sophos_folder.print_control_identifiers()
    sophos_subfolder = sophos_folder.child_window(title="Sophos", control_type="TreeItem")
    sophos_subfolder.print_control_identifiers()
    sophos_subfolder.click_input(double=True)
    #test_dlg.Sophos.click_input()
    #test_dlg.Pane17.click_input(double=True)
    #time.sleep(5)
    print(app.windows())


    

if __name__ == '__main__':
   if (checkIfOutlookRunning('OUTLOOK.EXE') == True):
        openEmailOutlookRunning()
