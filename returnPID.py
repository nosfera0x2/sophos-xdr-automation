import psutil

def returnPID(process):
    process_name = process
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            processID = proc.pid
            return (processID)
