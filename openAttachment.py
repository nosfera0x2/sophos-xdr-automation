# ----------------------------------------------------------------------------
# Name:         
# Description:  Python Script to Automate XDR Threat Intel
# Author:       Spencer Brown
# URL:          
# Date:         02/09/2022
# ----------------------------------------------------------------------------

from pywinauto import Application, Desktop
import time

#start the Outlook application
app = Application(backend="uia").start(r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK')
time.sleep(10)
#lets see what windows are available, and assign it to a dialogue
print(app.windows())
#Main application is presented as 'Outlook Today - Outlook'
#Define the main dialogue
main_dlg = app['Outlook Today - Outlook']
#To see if it works, lets see if we can print out the control identifiers.
#main_dlg.print_control_identifiers()
#success
bth2_tree = main_dlg.child_window(title="BTH2", control_type="TreeItem")
#bth2_tree.print_control_identifiers()
treeItem_two = bth2_tree.child_window(title="2", control_type="TreeItem")
treeItem_two.click_input(double=True)
time.sleep(5)
second_dlg = app['2 - Sophos - Outlook']
time.sleep(5)
#second_dlg.print_control_identifiers()
email = second_dlg.child_window(title="With Attachments, Subject Your flight has been successfully booked!, Received Mon 1/31, Size 136 KB, Flag Status Unflagged, ", control_type="DataItem")
email.click_input(double=True)
time.sleep(5)
print(app.windows())
email_dlg = app['Your flight has been successfully booked! - Message (HTML) ']
email_dlg.print_control_identifiers()
attachment = email_dlg.child_window(title="Attachments", control_type="Pane")
attachment.Button25.click_input(double=True)
