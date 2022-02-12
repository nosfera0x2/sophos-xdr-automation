# ----------------------------------------------------------------------------
# Name:         Enable Macro in Word Document Function
# Description:  Python Script to Automate XDR Threat Intel
# Author:       Spencer Brown
# URL:          
# Date:         02/09/2022
# ----------------------------------------------------------------------------

from pywinauto import Application, Desktop
import time
import win32gui
import win32con
from extractTitle import extractTitle

def enableMacro():
    #Connects to open microsoft word winow using regex, since window name is dynamic.
    app = Application(backend="uia").connect(title_re='.*DeltaFlightItinerary')
    window = app.windows() #app.windows() returns a list object.
    wordWindow = str(extractTitle(window[0])) #Extract title from list, and convert to string
    print("I'm looking for the window..." + wordWindow)
    guiWindow = win32gui.FindWindow(None,wordWindow) #Find our window with win32gui
    win32gui.ShowWindow(guiWindow,win32con.SW_SHOWMAXIMIZED) #Maximize the window. 
    main_dlg = app[wordWindow]
    main_dlg.set_focus() #Make window visible for clicking inputs. 
    main_dlg.EnableEditingButton.click_input() #Enable editing, pivots to enable content
    time.sleep(3)
    main_dlg.EnableContentButton.click_input() #Enable content button
