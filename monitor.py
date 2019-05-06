import win32serviceutil
import win32service
import win32event
import os
import logging
import inspect
import sys
import servicemanager


class IFaceAttReaderMonitor(win32serviceutil.ServiceFramework):
    _svc_name_ = "IFaceAttReaderMonitor"
    _svc_display_name_ = "IFaceAttReader Monitor service"
    _svc_description_ = "This is a python service monitor code "

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.logger = self._getLogger()
        self.run = True

    def _getLogger(self):
        logger = logging.getLogger('[IFaceAttReaderMonitor]')

        this_file = inspect.getfile(inspect.currentframe())
        dirpath = os.path.abspath(os.path.dirname(this_file))
        handler = logging.FileHandler(os.path.join(dirpath, "monitor-service.log"))

        formatter = logging.Formatter('%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
        handler.setFormatter(formatter)

        logger.addHandler(handler)
        logger.setLevel(logging.INFO)

        return logger

    def SvcDoRun(self):
        import wmi
        import time 
        c = wmi.WMI()
        while True:
          for service in c.Win32_Service(Name="IFaceAttReader"):
            result, = service.StartService()
            if result == 0:
              self.logger.info("Service " + service.Name + " started")
            elif result == 10:
              self.logger.info("service is running.")
            else:
              self.logger.info("Some problem" + str(result))
            break
          else:
            self.logger.info("Service not found")
          time.sleep(300)

    def SvcStop(self):
        self.logger.info("monitor service is stop....")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.run = False


if __name__ == '__main__':
        win32serviceutil.HandleCommandLine(IFaceAttReaderMonitor)
