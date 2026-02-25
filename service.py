import threading
import traceback

import servicemanager
import uvicorn
import win32event
import win32service
import win32serviceutil

from app.config import load_config
from app.main import app


class PrintServerDashboardService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PrintServerDashboard"
    _svc_display_name_ = "Print Server Dashboard"
    _svc_description_ = "Servidor web para monitoramento de impressao e contadores."

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.server = None
        self.server_thread = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        if self.server:
            self.server.should_exit = True
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg("Print Server Dashboard: service started")
        try:
            cfg = load_config()
            uv_cfg = uvicorn.Config(
                app=app,
                host=cfg.server_host,
                port=cfg.server_port,
                proxy_headers=True,
                log_level="info",
                log_config=None,
            )
            self.server = uvicorn.Server(uv_cfg)
            self.server_thread = threading.Thread(target=self.server.run, daemon=True)
            self.server_thread.start()

            while True:
                if win32event.WaitForSingleObject(self.stop_event, 1000) == win32event.WAIT_OBJECT_0:
                    break
                if self.server_thread and not self.server_thread.is_alive():
                    break
        except Exception:
            servicemanager.LogErrorMsg(
                "Print Server Dashboard: service error\n" + traceback.format_exc()
            )
        finally:
            if self.server:
                self.server.should_exit = True
            if self.server_thread and self.server_thread.is_alive():
                self.server_thread.join(timeout=10)
            servicemanager.LogInfoMsg("Print Server Dashboard: service stopped")


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(PrintServerDashboardService)
