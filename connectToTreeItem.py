from pywinauto import Application, Desktop
import time
from returnPID import returnPID
from processRunCheck import processRunCheck
from extractTitle import extractTitle
import win32gui
import win32con

 


outlook = returnPID("OUTLOOK.EXE")

def connectToTreeItem():
    app = Application(backend="uia").connect(process=outlook, visible_only=False)
    currentWindow= app.windows()
    outlookString = str(currentWindow[0])
    currentOutlookWindow = extractTitle(outlookString)
    print("The name of the window you are on is: " + currentOutlookWindow)
    window = win32gui.FindWindow(None,currentOutlookWindow)
    win32gui.ShowWindow(window, win32con.SW_SHOWMAXIMIZED)
    print("Here are the windows I now have available:" + str(app.windows()))
    outlook_dlg = app[currentOutlookWindow]
    outlook_dlg.set_focus()
    outlook_dlg.BTH2.click_input()
    #test_dlg.print_control_identifiers()
    outlook_dlg['2TreeItem'].click_input()
