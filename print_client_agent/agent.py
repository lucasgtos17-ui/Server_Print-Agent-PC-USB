import socket
import time
from typing import Dict, List
from urllib.parse import urlparse

import requests
import win32print

from config import AgentConfig, load_config


AGENT_VERSION = "1.2.0"


def list_printers() -> List[str]:
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags)
    return [p[2] for p in printers]


def _job_hash(job: Dict) -> str:
    return f"{job.get('JobId')}|{job.get('Submitted')}|{job.get('pPrinterName')}|{job.get('pDocument')}"


def poll_printer(printer_name: str) -> List[Dict]:
    handle = win32print.OpenPrinter(printer_name)
    try:
        jobs = win32print.EnumJobs(handle, 0, -1, 1)
        return jobs
    finally:
        win32print.ClosePrinter(handle)


def get_printer_model(printer_name: str) -> str:
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            info = win32print.GetPrinter(handle, 2) or {}
            return str(info.get("pDriverName", "") or "").strip()
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        return ""


def get_default_printer() -> str:
    try:
        return str(win32print.GetDefaultPrinter() or "").strip()
    except Exception:
        return ""


def resolve_printer_name(cfg: AgentConfig) -> str:
    if bool(getattr(cfg, "monitor_default_printer", False)):
        return get_default_printer()
    return str(cfg.printer_name or "").strip()


def _local_ip_for_server(server_url: str) -> str:
    try:
        parsed = urlparse(server_url)
        host = parsed.hostname
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
        if not host:
            return ""
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect((host, port))
            return str(s.getsockname()[0])
        finally:
            s.close()
    except Exception:
        return ""


def send_jobs(server_url: str, jobs: List[Dict], printer_name: str, printer_model: str, client_ip: str) -> None:
    if not jobs:
        return

    host = socket.gethostname()
    payload = []
    for j in jobs:
        total_pages = j.get("TotalPages", 0) or j.get("PagesPrinted", 0) or 0
        payload.append(
            {
                "job_id": j.get("JobId"),
                "user": j.get("pUserName", ""),
                "printer": printer_name,
                "document": j.get("pDocument", ""),
                "pages": total_pages,
                "copies": j.get("Copies", 1) or 1,
                "submitted": str(j.get("Submitted", "")),
                "client_host": host,
                "client_ip": client_ip,
                "printer_model": printer_model,
                "agent_id": f"{host}|{printer_name}",
                "agent_version": AGENT_VERSION,
            }
        )

    url = server_url.rstrip("/") + "/api/client-jobs"
    requests.post(url, json=payload, timeout=5)


def send_heartbeat(server_url: str, printer_name: str, printer_model: str, client_ip: str) -> None:
    host = socket.gethostname()
    payload = {
        "agent_id": f"{host}|{printer_name}",
        "host": host,
        "client_ip": client_ip,
        "printer_name": printer_name,
        "printer_model": printer_model,
        "agent_version": AGENT_VERSION,
    }
    url = server_url.rstrip("/") + "/api/agents/heartbeat"
    requests.post(url, json=payload, timeout=5)


def run_agent(cfg: AgentConfig, stop_event=None) -> None:
    if not cfg.server_url:
        raise ValueError("server_url must be configured")
    if not bool(getattr(cfg, "monitor_default_printer", False)) and not cfg.printer_name:
        raise ValueError("printer_name must be configured")

    seen = set()
    client_ip = _local_ip_for_server(cfg.server_url)
    last_heartbeat = 0.0
    current_printer = ""
    current_model = ""

    while True:
        if stop_event is not None and stop_event.is_set():
            break
        try:
            target_printer = resolve_printer_name(cfg)
            if not target_printer:
                time.sleep(max(2, cfg.poll_interval_sec))
                continue

            if target_printer != current_printer:
                current_printer = target_printer
                current_model = get_printer_model(current_printer)

            now = time.time()
            if now - last_heartbeat >= 30:
                send_heartbeat(cfg.server_url, current_printer, current_model, client_ip)
                last_heartbeat = now

            jobs = poll_printer(current_printer)
            new_jobs = []
            for j in jobs:
                key = _job_hash(j)
                if key in seen:
                    continue
                seen.add(key)
                new_jobs.append(j)
            if new_jobs:
                send_jobs(cfg.server_url, new_jobs, current_printer, current_model, client_ip)
        except Exception:
            pass
        time.sleep(max(2, cfg.poll_interval_sec))


if __name__ == "__main__":
    cfg = load_config()
    run_agent(cfg)
