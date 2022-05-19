import sys

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

def current_window(app):
    try:
        string = str(app.windows()[0])
        remove_first = string[24:]
        window = remove_first[:-8]
        return window
    except:
        print("Less or More than 1 window found.")

def revert_tree(app,check_file):

    dlg = app[currentWindow(app)]
    if check_id(check_file,"Atomic Red Team",dlg) == True:
        dlg = app[currentWindow(app)]
        dlg.child_window(title="Sophos", control_type="TreeItem").collapse()
        print("Tree is reset")
    else:
        print("Tree is already collapsed")
        pass
