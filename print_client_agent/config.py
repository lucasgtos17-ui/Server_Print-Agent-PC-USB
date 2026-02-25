import json
import os
import sys
from dataclasses import dataclass
from typing import Optional


@dataclass
class AgentConfig:
    server_url: str
    printer_name: str
    poll_interval_sec: int
    start_with_windows: bool
    monitor_default_printer: bool = False


def app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def default_config_path() -> str:
    local = os.getenv("LOCALAPPDATA")
    if local:
        cfg_dir = os.path.join(local, "PrintClientAgent")
        os.makedirs(cfg_dir, exist_ok=True)
        return os.path.join(cfg_dir, "config.json")
    return os.path.join(app_base_dir(), "config.json")


def load_config(config_path: Optional[str] = None) -> AgentConfig:
    cfg_path = config_path or default_config_path()
    if not os.path.exists(cfg_path):
        fallback = os.path.join(app_base_dir(), "config.json")
        if os.path.exists(fallback):
            cfg_path = fallback
        else:
            raise FileNotFoundError(f"Config file not found: {cfg_path}")

    with open(cfg_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    return AgentConfig(
        server_url=data.get("server_url", ""),
        printer_name=data.get("printer_name", ""),
        poll_interval_sec=int(data.get("poll_interval_sec", 5)),
        start_with_windows=bool(data.get("start_with_windows", False)),
        monitor_default_printer=bool(data.get("monitor_default_printer", False)),
    )


def save_config(config: AgentConfig, config_path: Optional[str] = None) -> None:
    cfg_path = config_path or default_config_path()
    os.makedirs(os.path.dirname(cfg_path), exist_ok=True)
    payload = {
        "server_url": config.server_url,
        "printer_name": config.printer_name,
        "poll_interval_sec": config.poll_interval_sec,
        "start_with_windows": config.start_with_windows,
        "monitor_default_printer": bool(config.monitor_default_printer),
    }
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    # Keep a local copy next to the executable/script so service accounts
    # without the same LOCALAPPDATA can still read configuration.
    fallback = os.path.join(app_base_dir(), "config.json")
    try:
        with open(fallback, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except Exception:
        pass
