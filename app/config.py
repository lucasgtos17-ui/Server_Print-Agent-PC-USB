import json
import os
from dataclasses import dataclass
from typing import Optional


@dataclass
class AppConfig:
    papercut_log_dir: str
    papercut_log_glob: str
    papercut_xmlrpc_url: Optional[str]
    papercut_auth_token: Optional[str]
    papercut_verify_tls: bool
    db_path: str
    server_host: str
    server_port: int
    default_days: int
    printer_poll_enabled: bool
    printer_poll_interval_sec: int


def _env(name: str, default: Optional[str] = None) -> Optional[str]:
    value = os.getenv(name)
    return value if value is not None and value != "" else default


def load_config(config_path: Optional[str] = None) -> AppConfig:
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    cfg_path = config_path or os.path.join(base_dir, "config.json")

    data = {}
    if os.path.exists(cfg_path):
        with open(cfg_path, "r", encoding="utf-8") as f:
            data = json.load(f)

    papercut_log_dir = _env("PAPERCUT_LOG_DIR", data.get("papercut_log_dir", ""))
    papercut_log_glob = _env("PAPERCUT_LOG_GLOB", data.get("papercut_log_glob", "printlog_*.log"))
    papercut_xmlrpc_url = _env("PAPERCUT_XMLRPC_URL", data.get("papercut_xmlrpc_url"))
    papercut_auth_token = _env("PAPERCUT_AUTH_TOKEN", data.get("papercut_auth_token"))
    papercut_verify_tls = str(_env("PAPERCUT_VERIFY_TLS", str(data.get("papercut_verify_tls", True)))).lower() == "true"
    db_path = _env("DB_PATH", data.get("db_path", "data\\papercut.db"))
    server_host = _env("SERVER_HOST", data.get("server_host", "0.0.0.0"))
    server_port = int(_env("SERVER_PORT", str(data.get("server_port", 8088))))
    default_days = int(_env("DEFAULT_DAYS", str(data.get("default_days", 7))))
    printer_poll_enabled = str(_env("PRINTER_POLL_ENABLED", str(data.get("printer_poll_enabled", True)))).lower() == "true"
    printer_poll_interval_sec = int(_env("PRINTER_POLL_INTERVAL_SEC", str(data.get("printer_poll_interval_sec", 300))))

    return AppConfig(
        papercut_log_dir=papercut_log_dir,
        papercut_log_glob=papercut_log_glob,
        papercut_xmlrpc_url=papercut_xmlrpc_url,
        papercut_auth_token=papercut_auth_token,
        papercut_verify_tls=papercut_verify_tls,
        db_path=db_path,
        server_host=server_host,
        server_port=server_port,
        default_days=default_days,
        printer_poll_enabled=printer_poll_enabled,
        printer_poll_interval_sec=printer_poll_interval_sec,
    )
