import os
import sqlite3
from datetime import datetime, timedelta
from typing import Any, Dict, Iterable, List, Optional, Tuple


def _connect(db_path: str) -> sqlite3.Connection:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_report_exclusions_table(cur: sqlite3.Cursor) -> None:
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS report_exclusions (
            kind TEXT,
            value TEXT,
            note TEXT,
            updated_at TEXT,
            PRIMARY KEY (kind, value)
        )
        """
    )


def init_db(db_path: str) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_hash TEXT UNIQUE,
            timestamp TEXT,
            user TEXT,
            full_name TEXT,
            printer TEXT,
            server TEXT,
            document TEXT,
            pages INTEGER,
            copies INTEGER,
            paper_size TEXT,
            language TEXT,
            job_size_kb INTEGER,
            cost REAL,
            client TEXT,
            grayscale TEXT,
            duplex TEXT,
            paper_height_mm TEXT,
            paper_width_mm TEXT,
            color_pages INTEGER,
            cost_adjustment TEXT,
            job_type TEXT
        )
        """
    )
    # Migrations
    cur.execute("PRAGMA table_info(jobs)")
    cols = {row[1] for row in cur.fetchall()}
    if "source" not in cols:
        cur.execute("ALTER TABLE jobs ADD COLUMN source TEXT")
    if "client_host" not in cols:
        cur.execute("ALTER TABLE jobs ADD COLUMN client_host TEXT")
    if "job_id" not in cols:
        cur.execute("ALTER TABLE jobs ADD COLUMN job_id TEXT")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS user_departments (
            user TEXT PRIMARY KEY,
            department TEXT,
            source TEXT,
            updated_at TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS printer_models (
            printer TEXT PRIMARY KEY,
            model TEXT,
            source TEXT,
            updated_at TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS departments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            updated_at TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS printer_departments (
            printer TEXT PRIMARY KEY,
            department_id INTEGER,
            updated_at TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS printer_sources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            ip TEXT,
            brand TEXT,
            model TEXT,
            serial TEXT,
            location TEXT,
            counter_url TEXT,
            enabled INTEGER DEFAULT 1,
            last_error TEXT,
            updated_at TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS printer_counters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            printer_name TEXT,
            ip TEXT,
            brand TEXT,
            model TEXT,
            timestamp TEXT,
            total_print INTEGER,
            total_copy INTEGER,
            total_scan INTEGER
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS client_agents (
            agent_id TEXT PRIMARY KEY,
            host TEXT,
            printer_name TEXT,
            printer_model TEXT,
            serial TEXT,
            location TEXT,
            ip TEXT,
            version TEXT,
            updated_at TEXT
        )
        """
    )
    cur.execute("PRAGMA table_info(printer_sources)")
    ps_cols = {row[1] for row in cur.fetchall()}
    if "serial" not in ps_cols:
        cur.execute("ALTER TABLE printer_sources ADD COLUMN serial TEXT")
    if "location" not in ps_cols:
        cur.execute("ALTER TABLE printer_sources ADD COLUMN location TEXT")
    cur.execute("PRAGMA table_info(client_agents)")
    ca_cols = {row[1] for row in cur.fetchall()}
    if "serial" not in ca_cols:
        cur.execute("ALTER TABLE client_agents ADD COLUMN serial TEXT")
    if "location" not in ca_cols:
        cur.execute("ALTER TABLE client_agents ADD COLUMN location TEXT")
    _ensure_report_exclusions_table(cur)
    conn.commit()
    conn.close()


def _to_int(value: Any) -> Optional[int]:
    try:
        return int(float(value))
    except Exception:
        return None


def _to_float(value: Any) -> Optional[float]:
    try:
        return float(value)
    except Exception:
        return None


def _normalize_since(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    s = str(value).strip()
    if not s:
        return None
    if len(s) == 10:
        return f"{s}T00:00:00"
    return s


def _normalize_until(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    s = str(value).strip()
    if not s:
        return None
    if len(s) == 10:
        return f"{s}T23:59:59.999999"
    return s


def _get_exclusions(conn: sqlite3.Connection) -> Dict[str, set]:
    cur = conn.cursor()
    out = {"printer": set(), "agent": set()}
    try:
        cur.execute("SELECT kind, value FROM report_exclusions")
    except sqlite3.OperationalError:
        # Backward-compatible with old databases created before this table existed.
        return out
    for row in cur.fetchall():
        kind = str(row[0] or "").strip().lower()
        value = str(row[1] or "").strip()
        if kind in out and value:
            out[kind].add(value)
    return out


def _add_exclusion_where(clauses: List[str], params: List[Any], ex: Dict[str, set], alias: str = "") -> None:
    prefix = f"{alias}." if alias else ""
    if ex["printer"]:
        placeholders = ",".join(["?"] * len(ex["printer"]))
        clauses.append(f"{prefix}printer NOT IN ({placeholders})")
        params.extend(sorted(ex["printer"]))
    if ex["agent"]:
        placeholders = ",".join(["?"] * len(ex["agent"]))
        clauses.append(
            f"(({prefix}source != 'client') OR ((COALESCE({prefix}client_host,'') || '|' || COALESCE({prefix}printer,'')) NOT IN ({placeholders})))"
        )
        params.extend(sorted(ex["agent"]))


def _job_hash(rec: Dict[str, Any]) -> str:
    parts = [
        str(rec.get("source") or ""),
        str(rec.get("client_host") or ""),
        str(rec.get("job_id") or ""),
        str(rec.get("timestamp") or ""),
        rec.get("user", ""),
        rec.get("printer", ""),
        rec.get("document", ""),
        str(rec.get("pages", "")),
        str(rec.get("copies", "")),
    ]
    return "|".join(parts)


def upsert_jobs(db_path: str, records: Iterable[Dict[str, Any]]) -> int:
    conn = _connect(db_path)
    cur = conn.cursor()
    inserted = 0

    for rec in records:
        ts = rec.get("timestamp")
        ts_str = ts.isoformat() if isinstance(ts, datetime) else (ts or "")
        job_hash = _job_hash({**rec, "timestamp": ts_str})

        cur.execute(
            """
            INSERT OR IGNORE INTO jobs (
                job_hash, timestamp, user, full_name, printer, server,
                document, pages, copies, paper_size, language, job_size_kb,
                cost, client, grayscale, duplex, paper_height_mm, paper_width_mm,
                color_pages, cost_adjustment, job_type, source, client_host, job_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                job_hash,
                ts_str,
                rec.get("user", ""),
                rec.get("full_name", ""),
                rec.get("printer", ""),
                rec.get("server", ""),
                rec.get("document", ""),
                _to_int(rec.get("pages")),
                _to_int(rec.get("copies")),
                rec.get("paper_size", ""),
                rec.get("language", ""),
                _to_int(rec.get("job_size_kb")),
                _to_float(rec.get("cost")),
                rec.get("client", ""),
                rec.get("grayscale", ""),
                rec.get("duplex", ""),
                rec.get("paper_height_mm", ""),
                rec.get("paper_width_mm", ""),
                _to_int(rec.get("color_pages")),
                rec.get("cost_adjustment", ""),
                rec.get("job_type", ""),
                rec.get("source", ""),
                rec.get("client_host", ""),
                rec.get("job_id", ""),
            ),
        )
        if cur.rowcount:
            inserted += 1

    conn.commit()
    conn.close()
    return inserted


def query_summary(db_path: str, days: int = 7) -> Dict[str, Any]:
    conn = _connect(db_path)
    cur = conn.cursor()
    since = datetime.now() - timedelta(days=days)
    ex = _get_exclusions(conn)

    where_parts = ["timestamp >= ?"]
    params: List[Any] = [since.isoformat()]
    _add_exclusion_where(where_parts, params, ex)
    where_sql = " AND ".join(where_parts)

    cur.execute(
        f"""
        SELECT COUNT(*) AS jobs, COALESCE(SUM(pages * COALESCE(copies,1)), 0) AS total_pages
        FROM jobs
        WHERE {where_sql}
        """,
        params,
    )
    totals = dict(cur.fetchone())

    cur.execute(
        f"""
        SELECT substr(timestamp, 1, 10) AS day, COALESCE(SUM(pages * COALESCE(copies,1)), 0) AS pages
        FROM jobs
        WHERE {where_sql}
        GROUP BY day
        ORDER BY day ASC
        """,
        params,
    )
    by_day = [dict(r) for r in cur.fetchall()]

    cur.execute(
        f"""
        SELECT user, COALESCE(SUM(pages * COALESCE(copies,1)), 0) AS pages
        FROM jobs
        WHERE {where_sql}
        GROUP BY user
        ORDER BY pages DESC
        LIMIT 10
        """,
        params,
    )
    top_users = [dict(r) for r in cur.fetchall()]

    cur.execute(
        f"""
        SELECT printer, COALESCE(SUM(pages * COALESCE(copies,1)), 0) AS pages
        FROM jobs
        WHERE {where_sql}
        GROUP BY printer
        ORDER BY pages DESC
        LIMIT 10
        """,
        params,
    )
    top_printers = [dict(r) for r in cur.fetchall()]

    conn.close()
    return {
        "totals": totals,
        "by_day": by_day,
        "top_users": top_users,
        "top_printers": top_printers,
    }


def query_jobs(
    db_path: str,
    limit: int = 50,
    user: Optional[str] = None,
    printer: Optional[str] = None,
    since: Optional[str] = None,
    until: Optional[str] = None,
) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    ex = _get_exclusions(conn)

    clauses = []
    params: List[Any] = []

    if user:
        clauses.append("user = ?")
        params.append(user)
    if printer:
        clauses.append("printer = ?")
        params.append(printer)
    since = _normalize_since(since)
    until = _normalize_until(until)

    if since:
        clauses.append("timestamp >= ?")
        params.append(since)
    if until:
        clauses.append("timestamp <= ?")
        params.append(until)
    _add_exclusion_where(clauses, params, ex)

    where = "WHERE " + " AND ".join(clauses) if clauses else ""

    sql = f"""
        SELECT * FROM jobs
        {where}
        ORDER BY timestamp DESC
        LIMIT ?
    """
    params.append(limit)

    cur.execute(sql, params)
    rows = [dict(r) for r in cur.fetchall()]
    return rows


def insert_client_jobs(db_path: str, records: Iterable[Dict[str, Any]]) -> int:
    normalized = []
    for rec in records:
        normalized.append(
            {
                "timestamp": rec.get("submitted") or rec.get("timestamp") or "",
                "user": rec.get("user", ""),
                "printer": rec.get("printer", ""),
                "document": rec.get("document", ""),
                "pages": rec.get("pages", 0),
                "copies": rec.get("copies", 1),
                "source": "client",
                "client_host": rec.get("client_host", ""),
                "job_id": rec.get("job_id", ""),
            }
        )
    return upsert_jobs(db_path, normalized)


def upsert_client_agent(
    db_path: str,
    agent_id: str,
    host: str,
    printer_name: str,
    printer_model: str = "",
    serial: str = "",
    location: str = "",
    ip: str = "",
    version: str = "",
) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO client_agents (agent_id, host, printer_name, printer_model, serial, location, ip, version, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(agent_id) DO UPDATE SET
            host=excluded.host,
            printer_name=excluded.printer_name,
            printer_model=excluded.printer_model,
            serial=excluded.serial,
            location=excluded.location,
            ip=excluded.ip,
            version=excluded.version,
            updated_at=excluded.updated_at
        """,
        (
            agent_id,
            host,
            printer_name,
            printer_model,
            serial,
            location,
            ip,
            version,
            datetime.now().isoformat(),
        ),
    )
    conn.commit()
    conn.close()


def list_client_agents(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT agent_id, host, printer_name, printer_model, serial, location, ip, version, updated_at
        FROM client_agents
        ORDER BY updated_at DESC
        """
    )
    rows = [dict(r) for r in cur.fetchall()]
    return rows


def update_client_agent(
    db_path: str,
    agent_id: str,
    host: str,
    printer_name: str,
    printer_model: str = "",
    serial: str = "",
    location: str = "",
    ip: str = "",
    version: str = "",
) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE client_agents
        SET host = ?, printer_name = ?, printer_model = ?, serial = ?, location = ?, ip = ?, version = ?, updated_at = ?
        WHERE agent_id = ?
        """,
        (host, printer_name, printer_model, serial, location, ip, version, datetime.now().isoformat(), agent_id),
    )
    conn.commit()
    conn.close()


def delete_client_agent(db_path: str, agent_id: str) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM client_agents WHERE agent_id = ?", (agent_id,))
    conn.commit()
    conn.close()


def get_client_agent(db_path: str, agent_id: str) -> Optional[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT agent_id, host, printer_name, printer_model, serial, location, ip, version, updated_at
        FROM client_agents
        WHERE agent_id = ?
        """,
        (agent_id,),
    )
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def upsert_user_department(db_path: str, user: str, department: str, source: str = "manual") -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO user_departments (user, department, source, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(user) DO UPDATE SET department=excluded.department, source=excluded.source, updated_at=excluded.updated_at
        """,
        (user, department, source, datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()


def upsert_printer_model(db_path: str, printer: str, model: str, source: str = "manual") -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO printer_models (printer, model, source, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(printer) DO UPDATE SET model=excluded.model, source=excluded.source, updated_at=excluded.updated_at
        """,
        (printer, model, source, datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()


def list_user_departments(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT user, department, source, updated_at FROM user_departments ORDER BY user ASC")
    rows = [dict(r) for r in cur.fetchall()]
    return rows


def list_printer_models(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT printer, model, source, updated_at FROM printer_models ORDER BY printer ASC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def query_report(
    db_path: str,
    since: Optional[str] = None,
    until: Optional[str] = None,
    group_by: str = "user",
) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    ex = _get_exclusions(conn)

    clauses = []
    params: List[Any] = []
    since = _normalize_since(since)
    until = _normalize_until(until)

    if since:
        clauses.append("j.timestamp >= ?")
        params.append(since)
    if until:
        clauses.append("j.timestamp <= ?")
        params.append(until)
    _add_exclusion_where(clauses, params, ex, alias="j")
    where = "WHERE " + " AND ".join(clauses) if clauses else ""

    group_map = {
        "user": "j.user",
        "department": "COALESCE(ud.department, d.name, 'Nao definido')",
        "printer": "j.printer",
        "model": "COALESCE(pm.model, 'Nao definido')",
    }
    group_expr = group_map.get(group_by, "j.user")

    sql = f"""
        SELECT
            {group_expr} AS group_name,
            COUNT(*) AS jobs,
            COALESCE(SUM(j.pages * COALESCE(j.copies,1)), 0) AS pages
        FROM jobs j
        LEFT JOIN user_departments ud ON ud.user = j.user
        LEFT JOIN printer_models pm ON pm.printer = j.printer
        LEFT JOIN printer_departments pdm ON pdm.printer = j.printer
        LEFT JOIN departments d ON d.id = pdm.department_id
        {where}
        GROUP BY group_name
        ORDER BY pages DESC
    """
    cur.execute(sql, params)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def query_job_printer_readings(
    db_path: str,
    since: Optional[str] = None,
    until: Optional[str] = None,
) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    ex = _get_exclusions(conn)

    since_n = _normalize_since(since)
    until_n = _normalize_until(until)

    clauses: List[str] = []
    params: List[Any] = []
    if until_n:
        clauses.append("j.timestamp <= ?")
        params.append(until_n)
    _add_exclusion_where(clauses, params, ex, alias="j")
    where = "WHERE " + " AND ".join(clauses) if clauses else ""

    initial_expr = "0"
    initial_params: List[Any] = []
    if since_n:
        initial_expr = "CASE WHEN j.timestamp < ? THEN COALESCE(j.pages,0) * COALESCE(j.copies,1) ELSE 0 END"
        initial_params.append(since_n)

    sql = f"""
        SELECT
            j.printer AS printer_name,
            COALESCE(SUM({initial_expr}), 0) AS reading_initial,
            COALESCE(SUM(COALESCE(j.pages,0) * COALESCE(j.copies,1)), 0) AS reading_final
        FROM jobs j
        {where}
        GROUP BY j.printer
        HAVING COALESCE(SUM(COALESCE(j.pages,0) * COALESCE(j.copies,1)), 0) > 0
        ORDER BY j.printer
    """
    cur.execute(sql, initial_params + params)
    rows = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT name, serial FROM printer_sources")
    src_map = {str(r["name"]): str(r["serial"] or "").strip() for r in cur.fetchall()}

    cur.execute("SELECT printer_name, serial, updated_at FROM client_agents ORDER BY updated_at DESC")
    ag_map: Dict[str, str] = {}
    for r in cur.fetchall():
        key = str(r["printer_name"] or "").strip()
        if key and key not in ag_map:
            ag_map[key] = str(r["serial"] or "").strip()

    cur.execute("SELECT printer, model FROM printer_models")
    model_map = {str(r["printer"]): str(r["model"] or "").strip() for r in cur.fetchall()}
    conn.close()

    out: List[Dict[str, Any]] = []
    for r in rows:
        name = str(r.get("printer_name") or "").strip()
        if not name:
            continue
        initial = int(r.get("reading_initial", 0) or 0)
        final = int(r.get("reading_final", 0) or 0)
        serial = src_map.get(name) or ag_map.get(name) or model_map.get(name) or "Não definido"
        out.append(
            {
                "printer_name": name,
                "serial": serial,
                "reading_initial": initial,
                "reading_final": final,
                "difference": max(0, final - initial),
            }
        )
    out.sort(key=lambda x: int(x.get("difference", 0)), reverse=True)
    return out


def upsert_printer_source(
    db_path: str,
    name: str,
    ip: str,
    brand: str,
    model: str,
    serial: str,
    location: str,
    counter_url: str,
    enabled: bool = True,
) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO printer_sources (name, ip, brand, model, serial, location, counter_url, enabled, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (name, ip, brand, model, serial, location, counter_url, 1 if enabled else 0, datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()


def list_departments(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, name, updated_at FROM departments ORDER BY name ASC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def create_department(db_path: str, name: str) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO departments (name, updated_at) VALUES (?, ?)",
        (str(name).strip(), datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()


def update_department(db_path: str, department_id: int, name: str) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "UPDATE departments SET name = ?, updated_at = ? WHERE id = ?",
        (str(name).strip(), datetime.now().isoformat(), int(department_id)),
    )
    conn.commit()
    conn.close()


def delete_department(db_path: str, department_id: int) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM printer_departments WHERE department_id = ?", (int(department_id),))
    cur.execute("DELETE FROM departments WHERE id = ?", (int(department_id),))
    conn.commit()
    conn.close()


def upsert_printer_department(db_path: str, printer: str, department_id: int) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO printer_departments (printer, department_id, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(printer) DO UPDATE SET department_id=excluded.department_id, updated_at=excluded.updated_at
        """,
        (str(printer).strip(), int(department_id), datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()


def delete_printer_department(db_path: str, printer: str) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM printer_departments WHERE printer = ?", (str(printer).strip(),))
    conn.commit()
    conn.close()


def list_printer_departments(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT pd.printer, pd.department_id, d.name AS department_name, pd.updated_at
        FROM printer_departments pd
        LEFT JOIN departments d ON d.id = pd.department_id
        ORDER BY pd.printer ASC
        """
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def list_known_printers(db_path: str) -> List[str]:
    conn = _connect(db_path)
    cur = conn.cursor()
    values = set()
    cur.execute("SELECT DISTINCT printer FROM jobs WHERE COALESCE(printer,'') <> ''")
    values.update(str(r[0]) for r in cur.fetchall())
    cur.execute("SELECT DISTINCT name FROM printer_sources WHERE COALESCE(name,'') <> ''")
    values.update(str(r[0]) for r in cur.fetchall())
    cur.execute("SELECT DISTINCT printer_name FROM client_agents WHERE COALESCE(printer_name,'') <> ''")
    values.update(str(r[0]) for r in cur.fetchall())
    cur.execute("SELECT DISTINCT printer FROM printer_models WHERE COALESCE(printer,'') <> ''")
    values.update(str(r[0]) for r in cur.fetchall())
    conn.close()
    return sorted(v for v in values if v)


def update_printer_source(
    db_path: str,
    source_id: int,
    name: str,
    ip: str,
    brand: str,
    model: str,
    serial: str,
    location: str,
    counter_url: str,
    enabled: bool = True,
) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE printer_sources
        SET name = ?, ip = ?, brand = ?, model = ?, serial = ?, location = ?, counter_url = ?, enabled = ?, updated_at = ?
        WHERE id = ?
        """,
        (
            name,
            ip,
            brand,
            model,
            serial,
            location,
            counter_url,
            1 if enabled else 0,
            datetime.now().isoformat(),
            int(source_id),
        ),
    )
    conn.commit()
    conn.close()


def delete_printer_source(db_path: str, source_id: int) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM printer_sources WHERE id = ?", (int(source_id),))
    conn.commit()
    conn.close()


def get_printer_source(db_path: str, source_id: int) -> Optional[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, name, ip, brand, model, serial, location, counter_url, enabled, last_error, updated_at
        FROM printer_sources
        WHERE id = ?
        """,
        (int(source_id),),
    )
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def list_printer_sources(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, name, ip, brand, model, serial, location, counter_url, enabled, last_error, updated_at
        FROM printer_sources
        ORDER BY id DESC
        """
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def set_printer_source_error(db_path: str, source_id: int, error: Optional[str]) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE printer_sources
        SET last_error = ?, updated_at = ?
        WHERE id = ?
        """,
        (error, datetime.now().isoformat(), source_id),
    )
    conn.commit()
    conn.close()


def insert_printer_counter(
    db_path: str,
    printer_name: str,
    ip: str,
    brand: str,
    model: str,
    total_print: int,
    total_copy: int,
    total_scan: int,
    timestamp: Optional[str] = None,
) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO printer_counters (printer_name, ip, brand, model, timestamp, total_print, total_copy, total_scan)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            printer_name,
            ip,
            brand,
            model,
            timestamp or datetime.now().isoformat(),
            _to_int(total_print) or 0,
            _to_int(total_copy) or 0,
            _to_int(total_scan) or 0,
        ),
    )
    conn.commit()
    conn.close()


def list_latest_counters(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            c.id,
            c.printer_name,
            c.ip,
            c.brand,
            c.model,
            c.timestamp,
            c.total_print,
            c.total_copy,
            c.total_scan,
            COALESCE(ps.serial, '') AS serial,
            COALESCE(ps.location, '') AS location
        FROM printer_counters c
        JOIN (
            SELECT printer_name, MAX(timestamp) AS ts
            FROM printer_counters
            GROUP BY printer_name
        ) last
        ON c.printer_name = last.printer_name AND c.timestamp = last.ts
        LEFT JOIN printer_sources ps ON ps.name = c.printer_name
        ORDER BY c.printer_name ASC
        """
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def query_counter_report(
    db_path: str,
    since: Optional[str] = None,
    until: Optional[str] = None,
    group_by: str = "printer",
    metric: str = "print",
) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    ex = _get_exclusions(conn)

    since = _normalize_since(since)
    until = _normalize_until(until)

    clauses = []
    params: List[Any] = []
    if since:
        clauses.append("timestamp >= ?")
        params.append(since)
    if until:
        clauses.append("timestamp <= ?")
        params.append(until)
    if ex["printer"]:
        placeholders = ",".join(["?"] * len(ex["printer"]))
        clauses.append(f"printer_name NOT IN ({placeholders})")
        params.extend(sorted(ex["printer"]))
    where = "WHERE " + " AND ".join(clauses) if clauses else ""

    cur.execute(
        f"""
        SELECT printer_name, ip, brand, model, timestamp, total_print, total_copy, total_scan
        FROM printer_counters
        {where}
        ORDER BY printer_name ASC, timestamp ASC
        """,
        params,
    )
    rows = [dict(r) for r in cur.fetchall()]

    # Last reading before initial period start (baseline).
    baseline_map: Dict[str, Dict[str, Any]] = {}
    if since:
        prev_day_map: Dict[str, Dict[str, Any]] = {}
        try:
            since_dt = datetime.fromisoformat(str(since))
            prev_start = (since_dt - timedelta(days=1)).isoformat()
            pre_day_clauses = ["timestamp >= ?", "timestamp < ?"]
            pre_day_params: List[Any] = [prev_start, since]
            if ex["printer"]:
                placeholders = ",".join(["?"] * len(ex["printer"]))
                pre_day_clauses.append(f"printer_name NOT IN ({placeholders})")
                pre_day_params.extend(sorted(ex["printer"]))
            pre_day_where = "WHERE " + " AND ".join(pre_day_clauses)
            cur.execute(
                f"""
                SELECT c.*
                FROM printer_counters c
                JOIN (
                    SELECT printer_name, MAX(timestamp) AS ts
                    FROM printer_counters
                    {pre_day_where}
                    GROUP BY printer_name
                ) p
                ON c.printer_name = p.printer_name AND c.timestamp = p.ts
                """,
                pre_day_params,
            )
            prev_day_map = {str(r["printer_name"]): dict(r) for r in cur.fetchall()}
        except Exception:
            prev_day_map = {}

        pre_clauses = ["timestamp < ?"]
        pre_params: List[Any] = [since]
        if ex["printer"]:
            placeholders = ",".join(["?"] * len(ex["printer"]))
            pre_clauses.append(f"printer_name NOT IN ({placeholders})")
            pre_params.extend(sorted(ex["printer"]))
        pre_where = "WHERE " + " AND ".join(pre_clauses)
        cur.execute(
            f"""
            SELECT c.*
            FROM printer_counters c
            JOIN (
                SELECT printer_name, MAX(timestamp) AS ts
                FROM printer_counters
                {pre_where}
                GROUP BY printer_name
            ) p
            ON c.printer_name = p.printer_name AND c.timestamp = p.ts
            """,
            pre_params,
        )
        baseline_map = {str(r["printer_name"]): dict(r) for r in cur.fetchall()}
        baseline_map.update(prev_day_map)

    # Metadata from configured IP printers.
    cur.execute("SELECT name, model, serial, location FROM printer_sources")
    src_map = {str(r["name"]): dict(r) for r in cur.fetchall()}

    # Metadata from agents.
    cur.execute("SELECT printer_name, printer_model, serial, location, host FROM client_agents")
    ag_map = {str(r["printer_name"]): dict(r) for r in cur.fetchall()}

    per_printer: Dict[str, Dict[str, Any]] = {}
    for r in rows:
        name = str(r.get("printer_name") or "")
        if name not in per_printer:
            per_printer[name] = {"first_in_range": r, "last_in_range": r}
        else:
            per_printer[name]["last_in_range"] = r

    metric_key = {
        "print": "total_print",
        "copy": "total_copy",
        "scan": "total_scan",
    }.get(metric, "total_print")

    grouped: Dict[str, Dict[str, Any]] = {}
    for name, v in per_printer.items():
        first_in_range = v["first_in_range"]
        last = v["last_in_range"]
        initial = baseline_map.get(name, first_in_range)

        initial_value = int(initial.get(metric_key, 0) or 0)
        final_value = int(last.get(metric_key, 0) or 0)
        delta = max(0, final_value - initial_value)

        src = src_map.get(name, {})
        ag = ag_map.get(name, {})
        resolved_model = (
            str(src.get("model") or "").strip()
            or str(ag.get("printer_model") or "").strip()
            or str(last.get("model") or "").strip()
        )
        resolved_serial = str(src.get("serial") or "").strip() or str(ag.get("serial") or "").strip() or resolved_model
        resolved_location = (
            str(src.get("location") or "").strip()
            or str(ag.get("location") or "").strip()
            or str(ag.get("host") or "").strip()
            or "Não definido"
        )
        resolved_brand = str(last.get("brand") or "").strip() or "Não definido"

        if group_by == "model":
            key = resolved_model or "Não definido"
        elif group_by == "brand":
            key = resolved_brand
        elif group_by == "serial":
            key = resolved_serial or "Não definido"
        else:
            key = name or "Não definido"

        if key not in grouped:
            grouped[key] = {
                "group_name": key,
                "jobs": 0,
                "pages": 0,
                "reading_initial": 0,
                "reading_final": 0,
                "difference": 0,
                "printer_name": "",
                "brand": "",
                "model": "",
                "serial": "",
                "location": "",
                "metric": metric,
            }
        grouped[key]["pages"] += int(delta)
        grouped[key]["jobs"] += 1
        grouped[key]["reading_initial"] += initial_value
        grouped[key]["reading_final"] += final_value
        grouped[key]["difference"] += int(delta)
        grouped[key]["printer_name"] = key if group_by == "printer" else (grouped[key]["printer_name"] or name)
        grouped[key]["brand"] = resolved_brand
        grouped[key]["model"] = resolved_model or "Não definido"
        grouped[key]["serial"] = resolved_serial or "Não definido"
        grouped[key]["location"] = resolved_location

    out = list(grouped.values())
    out.sort(key=lambda x: x["difference"], reverse=True)
    conn.close()
    return out


def query_counter_daily(
    db_path: str,
    since: Optional[str] = None,
    until: Optional[str] = None,
) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    ex = _get_exclusions(conn)

    since_n = _normalize_since(since)
    until_n = _normalize_until(until)

    clauses = []
    params: List[Any] = []
    if until_n:
        clauses.append("timestamp <= ?")
        params.append(until_n)
    if ex["printer"]:
        placeholders = ",".join(["?"] * len(ex["printer"]))
        clauses.append(f"printer_name NOT IN ({placeholders})")
        params.extend(sorted(ex["printer"]))
    where = "WHERE " + " AND ".join(clauses) if clauses else ""

    cur.execute(
        f"""
        SELECT printer_name, timestamp, total_print, total_copy
        FROM printer_counters
        {where}
        ORDER BY printer_name ASC, timestamp ASC
        """,
        params,
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    start_dt = None
    end_dt = None
    try:
        if since_n:
            start_dt = datetime.fromisoformat(since_n)
        if until_n:
            end_dt = datetime.fromisoformat(until_n)
    except Exception:
        start_dt = None
        end_dt = None

    per_day: Dict[str, int] = {}
    last_by_printer: Dict[str, Tuple[int, int]] = {}
    for r in rows:
        name = str(r.get("printer_name") or "")
        ts_raw = str(r.get("timestamp") or "")
        try:
            ts = datetime.fromisoformat(ts_raw)
        except Exception:
            continue

        curr_print = int(r.get("total_print", 0) or 0)
        curr_copy = int(r.get("total_copy", 0) or 0)
        prev_print, prev_copy = last_by_printer.get(name, (curr_print, curr_copy))
        delta = max(0, (curr_print - prev_print) + (curr_copy - prev_copy))
        last_by_printer[name] = (curr_print, curr_copy)

        if start_dt and ts < start_dt:
            continue
        if end_dt and ts > end_dt:
            continue
        day = ts.strftime("%Y-%m-%d")
        per_day[day] = per_day.get(day, 0) + delta

    out = [{"day": d, "pages": per_day[d]} for d in sorted(per_day.keys())]
    return out


def query_recent_counter_events(db_path: str, limit: int = 50) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    ex = _get_exclusions(conn)

    where = ""
    params: List[Any] = []
    if ex["printer"]:
        placeholders = ",".join(["?"] * len(ex["printer"]))
        where = f"WHERE printer_name NOT IN ({placeholders})"
        params.extend(sorted(ex["printer"]))

    sql = f"""
        WITH seq AS (
          SELECT
            printer_name,
            timestamp,
            COALESCE(total_print, 0) AS total_print,
            COALESCE(total_copy, 0) AS total_copy,
            LAG(COALESCE(total_print, 0)) OVER (PARTITION BY printer_name ORDER BY timestamp) AS prev_print,
            LAG(COALESCE(total_copy, 0)) OVER (PARTITION BY printer_name ORDER BY timestamp) AS prev_copy
          FROM printer_counters
          {where}
        )
        SELECT
          timestamp,
          printer_name,
          MAX(0, (total_print - COALESCE(prev_print, total_print))) AS delta_print,
          MAX(0, (total_copy - COALESCE(prev_copy, total_copy))) AS delta_copy
        FROM seq
        WHERE (total_print - COALESCE(prev_print, total_print)) > 0
           OR (total_copy - COALESCE(prev_copy, total_copy)) > 0
        ORDER BY timestamp DESC
        LIMIT ?
    """
    cur.execute(sql, params + [int(limit)])
    rows = []
    for r in cur.fetchall():
        dprint = int(r["delta_print"] or 0)
        dcopy = int(r["delta_copy"] or 0)
        doc = f"Contador IP: +{dprint} impressão, +{dcopy} cópia"
        rows.append(
            {
                "timestamp": str(r["timestamp"] or ""),
                "user": "-",
                "printer": str(r["printer_name"] or ""),
                "pages": dprint + dcopy,
                "copies": 1,
                "document": doc,
            }
        )
    conn.close()
    return rows


def list_report_exclusions(db_path: str) -> List[Dict[str, Any]]:
    conn = _connect(db_path)
    cur = conn.cursor()
    _ensure_report_exclusions_table(cur)
    cur.execute("SELECT kind, value, note, updated_at FROM report_exclusions ORDER BY kind, value")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def upsert_report_exclusion(db_path: str, kind: str, value: str, note: str = "") -> None:
    k = str(kind or "").strip().lower()
    if k not in ("printer", "agent"):
        raise ValueError("kind must be printer or agent")
    v = str(value or "").strip()
    if not v:
        raise ValueError("value is required")
    conn = _connect(db_path)
    cur = conn.cursor()
    _ensure_report_exclusions_table(cur)
    cur.execute(
        """
        INSERT INTO report_exclusions (kind, value, note, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(kind, value) DO UPDATE SET note=excluded.note, updated_at=excluded.updated_at
        """,
        (k, v, str(note or "").strip(), datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()


def delete_report_exclusion(db_path: str, kind: str, value: str) -> None:
    conn = _connect(db_path)
    cur = conn.cursor()
    _ensure_report_exclusions_table(cur)
    cur.execute("DELETE FROM report_exclusions WHERE kind = ? AND value = ?", (str(kind).strip().lower(), str(value).strip()))
    conn.commit()
    conn.close()
