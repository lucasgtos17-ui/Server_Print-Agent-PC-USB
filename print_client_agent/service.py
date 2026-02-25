import threading
import traceback
import time
import os
from datetime import datetime

import win32event
import win32service
import win32serviceutil


def _log(msg: str) -> None:
    try:
        base = os.getenv("PROGRAMDATA") or r"C:\ProgramData"
        log_dir = os.path.join(base, "PrintClientAgent")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, "service.log")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()} {msg}\n")
    except Exception:
        pass


class PrintClientAgentService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PrintClientAgent"
    _svc_display_name_ = "Print Client Agent"
    _svc_description_ = "Coleta jobs da impressora local e envia ao servidor."

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.thread_stop = threading.Event()

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.thread_stop.set()
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        _log("service started")
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        try:
            # Keep service alive even when configuration is missing/invalid.
            while not self.thread_stop.is_set():
                try:
                    from config import load_config
                    from agent import run_agent
                    cfg = load_config()
                    run_agent(cfg, stop_event=self.thread_stop)
                except Exception:
                    _log("service worker error\n" + traceback.format_exc())
                    # Backoff before retrying load/run cycle.
                    for _ in range(10):
                        if self.thread_stop.is_set():
                            break
                        time.sleep(1)
        except Exception:
            _log("service error\n" + traceback.format_exc())
        finally:
            _log("service stopped")


if __name__ == "__main__":
    try:
        win32serviceutil.HandleCommandLine(PrintClientAgentService)
    except Exception:
        _log("HandleCommandLine fatal error\n" + traceback.format_exc())
        raise
