import os
import win32serviceutil
import win32service
import win32event
import servicemanager
import listener

class ListenerService(win32serviceutil.ServiceFramework):
    _svc_name_ = "ListenerService"
    _svc_display_name_ = "Listener Service"
    _svc_description_ = "Listens for the trigger phrase and runs the voice assistant"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.isAlive = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.isAlive = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        listener.main()

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ListenerService)
