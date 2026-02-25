import os
from datetime import datetime
from typing import Dict, Iterable, List, Optional


FIELDS = [
    "date",
    "time",
    "user",
    "full_name",
    "printer",
    "server",
    "document",
    "pages",
    "copies",
    "paper_size",
    "language",
    "job_size_kb",
    "cost",
    "client",
    "grayscale",
    "duplex",
    "paper_height_mm",
    "paper_width_mm",
    "color_pages",
    "cost_adjustment",
    "job_type",
]


def parse_printlog_line(line: str) -> Optional[Dict[str, str]]:
    line = line.strip("\n")
    if not line or line.startswith("#"):
        return None
    parts = line.split("\t")
    if len(parts) < 8:
        return None

    record = {}
    for idx, name in enumerate(FIELDS):
        record[name] = parts[idx] if idx < len(parts) else ""

    date_str = record.get("date", "")
    time_str = record.get("time", "")
    timestamp = None
    try:
        timestamp = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
    except Exception:
        try:
            timestamp = datetime.strptime(f"{date_str} {time_str}", "%Y/%m/%d %H:%M:%S")
        except Exception:
            timestamp = None

    record["timestamp"] = timestamp
    return record


def iter_printlog_file(path: str) -> Iterable[Dict[str, str]]:
    if not os.path.exists(path):
        return []

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            rec = parse_printlog_line(line)
            if rec:
                yield rec


def iter_printlog_files(paths: List[str]) -> Iterable[Dict[str, str]]:
    for p in paths:
        for rec in iter_printlog_file(p):
            yield rec
