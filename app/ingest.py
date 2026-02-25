import argparse
import glob
import os
from datetime import datetime, timedelta

from app.config import load_config
from app.log_parser import iter_printlog_files
from app.storage import init_db, upsert_jobs


def _select_files(log_dir: str, pattern: str, since_days: int) -> list[str]:
    if not log_dir:
        return []
    full_pattern = os.path.join(log_dir, pattern)
    files = glob.glob(full_pattern)
    if since_days <= 0:
        return sorted(files)

    cutoff = datetime.now() - timedelta(days=since_days)
    selected = []
    for f in files:
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(f))
            if mtime >= cutoff:
                selected.append(f)
        except Exception:
            continue
    return sorted(selected)


def main() -> None:
    parser = argparse.ArgumentParser(description="Ingest PaperCut print logs into SQLite")
    parser.add_argument("--config", default=None, help="Path to config.json")
    parser.add_argument("--log-dir", default=None, help="PaperCut print log directory")
    parser.add_argument("--glob", default=None, help="Log file glob pattern")
    parser.add_argument("--since-days", type=int, default=7, help="Only ingest files modified in the last N days")
    args = parser.parse_args()

    cfg = load_config(args.config)
    log_dir = args.log_dir or cfg.papercut_log_dir
    pattern = args.glob or cfg.papercut_log_glob

    if not log_dir:
        raise SystemExit("papercut_log_dir not configured")

    init_db(cfg.db_path)
    files = _select_files(log_dir, pattern, args.since_days)
    if not files:
        print("No log files found")
        return

    inserted = upsert_jobs(cfg.db_path, iter_printlog_files(files))
    print(f"Inserted {inserted} job(s)")


if __name__ == "__main__":
    main()
