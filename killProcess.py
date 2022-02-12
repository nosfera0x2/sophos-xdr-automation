import wmi

def killProcess(process):
    ti = 0
    name = process
    print('The process I am seeking to kill is: ' + str(name))
    f = wmi.WMI()

    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()
            print('I have found, and killed ' + str(name))
            ti +=1
    if ti == 0:
        print('There are no running instances of ' + str(name))
