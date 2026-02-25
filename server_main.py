import os
import subprocess
import sys
import time
import webbrowser

import uvicorn

from app.main import app, cfg


def _service_exe_path() -> str:
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "PrintServerDashboardService.exe")


def _service_exists() -> bool:
    result = subprocess.run(
        ["sc.exe", "query", "PrintServerDashboard"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.returncode == 0


def _service_running() -> bool:
    result = subprocess.run(
        ["sc.exe", "query", "PrintServerDashboard"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.returncode == 0 and "RUNNING" in result.stdout


def _try_install_and_start_service() -> bool:
    svc_exe = _service_exe_path()
    if not os.path.exists(svc_exe):
        return False

    if not _service_exists():
        subprocess.run([svc_exe, "install"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        time.sleep(1)

    if not _service_running():
        subprocess.run([svc_exe, "start"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        time.sleep(2)

    return _service_running()


def _run_embedded_server() -> None:
    uvicorn.run(
        app,
        host=cfg.server_host,
        port=cfg.server_port,
        proxy_headers=True,
        log_level="info",
        log_config=None,
    )


if __name__ == "__main__":
    # EXE behavior: launcher installs/starts service automatically.
    if getattr(sys, "frozen", False):
        if _try_install_and_start_service():
            webbrowser.open(f"http://{cfg.server_host}:{cfg.server_port}/")
            raise SystemExit(0)
    _run_embedded_server()
