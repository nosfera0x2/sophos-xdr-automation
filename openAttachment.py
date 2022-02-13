# ----------------------------------------------------------------------------
# Name:         Open Email Attachment Function
# Description:  Python Script to Automate XDR Threat Intel
# Author:       Spencer Brown
# URL:          
# Date:         02/09/2022
# ----------------------------------------------------------------------------

from pywinauto import Application, Desktop
import time
import os 
from processRunCheck import processRunCheck
from killProcess import killProcess
from enableMacro import enableMacro


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
    email = second_dlg['Attachment With Attachments, Subject Your flight has been successfully booked!, Received Mon 1/31, Size 136 KB, Flag Status Unflagged, DataItem']
    email.click_input(double=True)
    email_dlg = app['Your flight has been successfully booked! - Message (HTML) ']
    #defines the attachment pane in the opened email message
    attachment = email_dlg['AttachmentsPane']
    attachment.AttachmentoptionsButton.click_input()
    email_dlg.ContextMenu.Open.click_input()


if __name__ == '__main__':

    killProcess('WINWORD.EXE')

    if (processRunCheck('outlook.exe') == False):
        openEmailAttachment()
        time.sleep(10)
        enableMacro()
        
    else:
        print('Outlook is currently running.')
        print('I am now going to kill the Outlook Process.')
        killProcess('OUTLOOK.EXE')
        time.sleep(10)
        print('Outlook will now start.')
        openEmailAttachment()
        time.sleep(10)
        enableMacro()
      
