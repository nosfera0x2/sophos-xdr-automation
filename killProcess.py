def killProcess(process):
    ti = 0
    name = process
    f = wmi.WMI()

    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()
            ti +=1
    if ti == 0:
        print('Process not found!')
