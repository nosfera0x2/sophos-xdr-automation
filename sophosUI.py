import psutil
from pywinauto import Application, Desktop
import time


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

def tamperCheck():
    app = Application(backend="uia").connect(process=sophosUI, visible_only=False)
    dlg = app['Sophos Endpoint Agent']
    checkbox = dlg.child_window(auto_id="OverrideSettingsCheckBox", control_type="CheckBox").wrapper_object()
    value = checkbox.get_toggle_state()
    return value


sophosUI = returnPID("Sophos UI.exe")

def SophosUI():
    app = Application(backend="uia").connect(path="explorer.exe")
    sys_tray = app.window(class_name="Shell_TrayWnd")
    sys_tray.child_window(title='Sophos Endpoint Agent').click_input(double=True)
    app = Application(backend="uia").connect(process=sophosUI, visible_only=False)
    dlg = app['Sophos Endpoint Agent']

    dlg['SettingsRadioButton'].click_input()
    time.sleep(3)
    if (tamperCheck() == 1):
        dlg.child_window(auto_id="OverrideSettingsCheckBox", control_type="CheckBox").click_input()
        dlg.child_window(title="Close", auto_id="NavBarCloseButton", control_type="Button").click_input()
    else:
        dlg.child_window(auto_id="OverrideSettingsCheckBox", control_type="CheckBox").click_input()
        dlg.child_window(auto_id="Enable Deep Learning", control_type="Button").click_input()
        dlg.child_window(auto_id="Ransomware Detection", control_type="Button").click_input()
        dlg.child_window(auto_id="Exploit Mitigation", control_type="Button").click_input()
        dlg.child_window(auto_id="Malicious Behavior Detection", control_type="Button").click_input()
        dlg.child_window(auto_id="AMSI Protection", control_type="Button").click_input()
        dlg.child_window(title="Close", auto_id="NavBarCloseButton", control_type="Button").click_input()

SophosUI()
