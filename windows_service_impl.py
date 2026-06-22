# -*- coding: utf-8 -*-
"""
Реализация Windows Service для Loodsman API
"""

import sys
import os
import win32serviceutil
import win32service
import win32event
import servicemanager

# Добавляем путь к проекту
sys.path.append(os.path.dirname(__file__))

class LoodsmanAPIService(win32serviceutil.ServiceFramework):
    _svc_name_ = "LoodsmanAPI"
    _svc_display_name_ = "Loodsman API Service"
    _svc_description_ = "FastAPI сервис для выгрузки данных из ЛОЦМАН"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.process = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        if self.process:
            self.process.terminate()
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        import servicemanager
        import subprocess
        
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        
        # Запускаем uvicorn
        cmd = [
            sys.executable, "-m", "uvicorn", "app.main:app",
            "--host", "0.0.0.0", "--port", "8000"
        ]
        
        self.process = subprocess.Popen(cmd)
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(LoodsmanAPIService)
