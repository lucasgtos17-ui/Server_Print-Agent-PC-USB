from fastapi import FastAPI, Query, Body
from fastapi.responses import HTMLResponse, Response
from typing import Optional
import threading
import time

from app.config import load_config
from app.ingest import _select_files
from app.log_parser import iter_printlog_files
from app.printer_scraper import fetch_counters
from app.storage import (
    create_department,
    delete_client_agent,
    delete_department,
    delete_printer_department,
    delete_printer_source,
    delete_report_exclusion,
    get_client_agent,
    get_printer_source,
    init_db,
    insert_printer_counter,
    insert_client_jobs,
    list_departments,
    list_known_printers,
    list_client_agents,
    list_latest_counters,
    list_printer_departments,
    list_report_exclusions,
    list_printer_models,
    list_printer_sources,
    list_user_departments,
    query_counter_report,
    query_counter_daily,
    query_jobs,
    query_job_printer_readings,
    query_recent_counter_events,
    query_report,
    query_summary,
    set_printer_source_error,
    update_department,
    update_client_agent,
    update_printer_source,
    upsert_printer_department,
    upsert_report_exclusion,
    upsert_client_agent,
    upsert_jobs,
    upsert_printer_model,
    upsert_printer_source,
    upsert_user_department,
)

app = FastAPI(title="Print Server Dashboard")

cfg = load_config()

_poll_thread_started = False


@app.on_event("startup")
def startup() -> None:
    init_db(cfg.db_path)
    if cfg.papercut_log_dir:
        files = _select_files(cfg.papercut_log_dir, cfg.papercut_log_glob, cfg.default_days)
        upsert_jobs(cfg.db_path, iter_printlog_files(files))
    if cfg.printer_poll_enabled:
        _start_printer_poll_thread()


def _scan_all_printers():
    sources = list_printer_sources(cfg.db_path)
    results = []
    for src in sources:
        if not src.get("enabled"):
            continue
        try:
            counters = fetch_counters(src.get("counter_url", ""), src.get("brand", ""))
            insert_printer_counter(
                cfg.db_path,
                printer_name=src.get("name", ""),
                ip=src.get("ip", ""),
                brand=src.get("brand", ""),
                model=src.get("model", ""),
                total_print=counters.get("print", 0),
                total_copy=counters.get("copy", 0),
                total_scan=counters.get("scan", 0),
            )
            set_printer_source_error(cfg.db_path, int(src.get("id", 0)), None)
            results.append({"printer": src.get("name", ""), "ok": True})
        except Exception as e:
            set_printer_source_error(cfg.db_path, int(src.get("id", 0)), str(e))
            results.append({"printer": src.get("name", ""), "ok": False, "error": str(e)})
    return results


def _poll_loop():
    while True:
        try:
            _scan_all_printers()
        except Exception:
            pass
        time.sleep(max(30, cfg.printer_poll_interval_sec))


def _start_printer_poll_thread():
    global _poll_thread_started
    if _poll_thread_started:
        return
    _poll_thread_started = True
    t = threading.Thread(target=_poll_loop, daemon=True)
    t.start()


@app.get("/api/summary")
def api_summary(days: int = Query(default=None)):
    d = days if days is not None else cfg.default_days
    return query_summary(cfg.db_path, d)


@app.get("/api/jobs")
def api_jobs(
    limit: int = 50,
    user: Optional[str] = None,
    printer: Optional[str] = None,
    since: Optional[str] = None,
    until: Optional[str] = None,
):
    return query_jobs(cfg.db_path, limit=limit, user=user, printer=printer, since=since, until=until)


@app.get("/api/user-departments")
def api_user_departments():
    return list_user_departments(cfg.db_path)


@app.post("/api/user-departments")
def api_user_departments_upsert(payload: dict = Body(...)):
    user = str(payload.get("user", "")).strip()
    department = str(payload.get("department", "")).strip()
    source = str(payload.get("source", "manual")).strip() or "manual"
    if not user or not department:
        return {"ok": False, "error": "user and department are required"}
    upsert_user_department(cfg.db_path, user, department, source)
    return {"ok": True}


@app.get("/api/printer-models")
def api_printer_models():
    return list_printer_models(cfg.db_path)


@app.post("/api/printer-models")
def api_printer_models_upsert(payload: dict = Body(...)):
    printer = str(payload.get("printer", "")).strip()
    model = str(payload.get("model", "")).strip()
    source = str(payload.get("source", "manual")).strip() or "manual"
    if not printer or not model:
        return {"ok": False, "error": "printer and model are required"}
    upsert_printer_model(cfg.db_path, printer, model, source)
    return {"ok": True}


@app.get("/api/departments")
def api_departments():
    return list_departments(cfg.db_path)


@app.post("/api/departments")
def api_departments_create(payload: dict = Body(...)):
    name = str(payload.get("name", "")).strip()
    if not name:
        return {"ok": False, "error": "name is required"}
    try:
        create_department(cfg.db_path, name)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.put("/api/departments/{department_id}")
def api_departments_update(department_id: int, payload: dict = Body(...)):
    name = str(payload.get("name", "")).strip()
    if not name:
        return {"ok": False, "error": "name is required"}
    try:
        update_department(cfg.db_path, department_id, name)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.delete("/api/departments/{department_id}")
def api_departments_delete(department_id: int):
    delete_department(cfg.db_path, department_id)
    return {"ok": True}


@app.get("/api/printers-known")
def api_printers_known():
    return list_known_printers(cfg.db_path)


@app.get("/api/printer-departments")
def api_printer_departments():
    return list_printer_departments(cfg.db_path)


@app.post("/api/printer-departments")
def api_printer_departments_upsert(payload: dict = Body(...)):
    printer = str(payload.get("printer", "")).strip()
    department_id = payload.get("department_id")
    if not printer or not department_id:
        return {"ok": False, "error": "printer and department_id are required"}
    upsert_printer_department(cfg.db_path, printer, int(department_id))
    return {"ok": True}


@app.delete("/api/printer-departments")
def api_printer_departments_delete(printer: str):
    p = str(printer or "").strip()
    if not p:
        return {"ok": False, "error": "printer is required"}
    delete_printer_department(cfg.db_path, p)
    return {"ok": True}


@app.post("/api/client-jobs")
def api_client_jobs(payload: list = Body(...)):
    if not isinstance(payload, list):
        return {"ok": False, "error": "payload must be a list"}
    # Auto-register agents/printers coming from client payload
    for rec in payload:
        host = str(rec.get("client_host", "")).strip()
        printer_name = str(rec.get("printer", "")).strip()
        printer_model = str(rec.get("printer_model", "")).strip()
        printer_serial = str(rec.get("printer_serial", "")).strip()
        location = str(rec.get("location", "")).strip()
        agent_id = str(rec.get("agent_id", "")).strip() or (f"{host}|{printer_name}" if host and printer_name else "")
        if agent_id and host and printer_name:
            upsert_client_agent(
                cfg.db_path,
                agent_id=agent_id,
                host=host,
                printer_name=printer_name,
                printer_model=printer_model,
                serial=printer_serial,
                location=location,
                ip=str(rec.get("client_ip", "")).strip(),
                version=str(rec.get("agent_version", "")).strip(),
            )
            if printer_model:
                upsert_printer_model(cfg.db_path, printer_name, printer_model, "agent")
    inserted = insert_client_jobs(cfg.db_path, payload)
    return {"ok": True, "inserted": inserted}


@app.post("/api/agents/heartbeat")
def api_agents_heartbeat(payload: dict = Body(...)):
    host = str(payload.get("host", "")).strip()
    printer_name = str(payload.get("printer_name", "")).strip()
    agent_id = str(payload.get("agent_id", "")).strip() or (f"{host}|{printer_name}" if host and printer_name else "")
    if not agent_id or not host or not printer_name:
        return {"ok": False, "error": "agent_id/host/printer_name are required"}
    printer_model = str(payload.get("printer_model", "")).strip()
    printer_serial = str(payload.get("printer_serial", "")).strip()
    location = str(payload.get("location", "")).strip()
    upsert_client_agent(
        cfg.db_path,
        agent_id=agent_id,
        host=host,
        printer_name=printer_name,
        printer_model=printer_model,
        serial=printer_serial,
        location=location,
        ip=str(payload.get("client_ip", "")).strip(),
        version=str(payload.get("agent_version", "")).strip(),
    )
    if printer_model:
        upsert_printer_model(cfg.db_path, printer_name, printer_model, "agent")
    return {"ok": True}


@app.get("/api/agents")
def api_agents():
    return list_client_agents(cfg.db_path)


@app.put("/api/agents/{agent_id}")
def api_agents_update(agent_id: str, payload: dict = Body(...)):
    current = get_client_agent(cfg.db_path, agent_id)
    if not current:
        return {"ok": False, "error": "agent not found"}
    host = str(payload.get("host", current.get("host", ""))).strip()
    printer_name = str(payload.get("printer_name", current.get("printer_name", ""))).strip()
    printer_model = str(payload.get("printer_model", current.get("printer_model", ""))).strip()
    serial = str(payload.get("serial", current.get("serial", ""))).strip()
    location = str(payload.get("location", current.get("location", ""))).strip()
    ip = str(payload.get("ip", current.get("ip", ""))).strip()
    version = str(payload.get("version", current.get("version", ""))).strip()
    if not host or not printer_name:
        return {"ok": False, "error": "host and printer_name are required"}
    update_client_agent(cfg.db_path, agent_id, host, printer_name, printer_model, serial, location, ip, version)
    if printer_model:
        upsert_printer_model(cfg.db_path, printer_name, printer_model, "agent")
    return {"ok": True}


@app.delete("/api/agents/{agent_id}")
def api_agents_delete(agent_id: str):
    delete_client_agent(cfg.db_path, agent_id)
    return {"ok": True}


@app.get("/api/printer-sources")
def api_printer_sources():
    return list_printer_sources(cfg.db_path)


@app.post("/api/printer-sources")
def api_printer_sources_upsert(payload: dict = Body(...)):
    source_id = payload.get("id")
    name = str(payload.get("name", "")).strip()
    ip = str(payload.get("ip", "")).strip()
    brand = str(payload.get("brand", "")).strip()
    model = str(payload.get("model", "")).strip()
    serial = str(payload.get("serial", "")).strip()
    location = str(payload.get("location", "")).strip()
    counter_url = str(payload.get("counter_url", "")).strip()
    if not name or not ip or not counter_url:
        return {"ok": False, "error": "name, ip and counter_url are required"}
    if source_id:
        update_printer_source(cfg.db_path, int(source_id), name, ip, brand, model, serial, location, counter_url, True)
        return {"ok": True, "updated": True}
    upsert_printer_source(cfg.db_path, name, ip, brand, model, serial, location, counter_url, True)
    return {"ok": True, "created": True}


@app.delete("/api/printer-sources/{source_id}")
def api_printer_sources_delete(source_id: int):
    delete_printer_source(cfg.db_path, source_id)
    return {"ok": True}


@app.post("/api/printer-sources/{source_id}/test")
def api_printer_source_test(source_id: int):
    src = get_printer_source(cfg.db_path, source_id)
    if not src:
        return {"ok": False, "error": "printer source not found"}
    try:
        counters = fetch_counters(src.get("counter_url", ""), src.get("brand", ""))
        insert_printer_counter(
            cfg.db_path,
            printer_name=src.get("name", ""),
            ip=src.get("ip", ""),
            brand=src.get("brand", ""),
            model=src.get("model", ""),
            total_print=counters.get("print", 0),
            total_copy=counters.get("copy", 0),
            total_scan=counters.get("scan", 0),
        )
        set_printer_source_error(cfg.db_path, source_id, None)
        return {"ok": True, "counters": counters}
    except Exception as e:
        set_printer_source_error(cfg.db_path, source_id, str(e))
        return {"ok": False, "error": str(e)}


@app.post("/api/printer-scan")
def api_printer_scan():
    results = _scan_all_printers()
    return {"ok": True, "results": results}


@app.get("/api/printer-counters")
def api_printer_counters():
    return list_latest_counters(cfg.db_path)


@app.get("/api/exclusions")
def api_exclusions():
    return list_report_exclusions(cfg.db_path)


@app.post("/api/exclusions")
def api_exclusions_upsert(payload: dict = Body(...)):
    kind = str(payload.get("kind", "")).strip().lower()
    value = str(payload.get("value", "")).strip()
    note = str(payload.get("note", "")).strip()
    if kind not in ("printer", "agent"):
        return {"ok": False, "error": "kind must be printer or agent"}
    if not value:
        return {"ok": False, "error": "value is required"}
    try:
        upsert_report_exclusion(cfg.db_path, kind, value, note)
    except Exception as e:
        return {"ok": False, "error": str(e)}
    return {"ok": True}


@app.delete("/api/exclusions")
def api_exclusions_delete(kind: str = Query(...), value: str = Query(...)):
    delete_report_exclusion(cfg.db_path, kind, value)
    return {"ok": True}


@app.get("/report-counters")
def report_counters_export(
    format: str = "csv",
    group_by: str = "printer",
    metric: str = "print",
    since: Optional[str] = None,
    until: Optional[str] = None,
):
    fmt = (format or "csv").lower()
    group = (group_by or "printer").lower()
    metric = (metric or "print").lower()
    rows = query_counter_report(cfg.db_path, since=since, until=until, group_by=group, metric=metric)
    group_labels = {
        "printer": "Impressora",
        "brand": "Marca",
        "model": "Modelo",
        "serial": "Serial",
    }
    metric_labels = {
        "print": "Impressões",
        "copy": "Cópias",
        "scan": "Scans",
    }
    headers = ["grupo", "impressora", "marca", "modelo", "serial", "local", "leitura_inicial", "leitura_final", "diferenca"]

    if fmt == "csv":
        import csv
        import io

        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(headers)
        for r in rows:
            writer.writerow(
                [
                    r.get("group_name", ""),
                    r.get("printer_name", ""),
                    r.get("brand", ""),
                    r.get("model", ""),
                    r.get("serial", ""),
                    r.get("location", ""),
                    r.get("reading_initial", 0),
                    r.get("reading_final", 0),
                    r.get("difference", 0),
                ]
            )
        data = buf.getvalue().encode("utf-8")
        return Response(
            content=data,
            media_type="text/csv",
            headers={"Content-Disposition": "attachment; filename=relatorio-contadores.csv"},
        )

    if fmt == "xlsx":
        import io
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "Relatório"
        ws.append(headers)
        for r in rows:
            ws.append(
                [
                    r.get("group_name", ""),
                    r.get("printer_name", ""),
                    r.get("brand", ""),
                    r.get("model", ""),
                    r.get("serial", ""),
                    r.get("location", ""),
                    r.get("reading_initial", 0),
                    r.get("reading_final", 0),
                    r.get("difference", 0),
                ]
            )
        out = io.BytesIO()
        wb.save(out)
        return Response(
            content=out.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=relatorio-contadores.xlsx"},
        )

    if fmt == "pdf":
        import io
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas

        out = io.BytesIO()
        c = canvas.Canvas(out, pagesize=landscape(A4))
        width, height = landscape(A4)
        x = 24
        y = height - 50
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x, y, "Relatório de Contadores")
        y -= 20
        c.setFont("Helvetica", 10)
        c.drawString(x, y, f"Agrupar por: {group_labels.get(group, group)}")
        y -= 20
        c.drawString(x, y, f"Métrica: {metric_labels.get(metric, metric)}")
        y -= 20
        c.drawString(x, y, f"Período: {since or '-'} até {until or '-'}")
        y -= 30
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x, y, "Grupo")
        c.drawString(x + 110, y, "Impressora")
        c.drawString(x + 220, y, "Marca")
        c.drawString(x + 290, y, "Modelo")
        c.drawString(x + 390, y, "Serial")
        c.drawString(x + 490, y, "Local")
        c.drawRightString(x + 600, y, "Inicial")
        c.drawRightString(x + 680, y, "Final")
        c.drawRightString(x + 760, y, "Diferença")
        y -= 14
        c.setFont("Helvetica", 8)
        for r in rows:
            if y < 60:
                c.showPage()
                y = height - 50
                c.setFont("Helvetica-Bold", 8)
                c.drawString(x, y, "Grupo")
                c.drawString(x + 110, y, "Impressora")
                c.drawString(x + 220, y, "Marca")
                c.drawString(x + 290, y, "Modelo")
                c.drawString(x + 390, y, "Serial")
                c.drawString(x + 490, y, "Local")
                c.drawRightString(x + 600, y, "Inicial")
                c.drawRightString(x + 680, y, "Final")
                c.drawRightString(x + 760, y, "Diferença")
                y -= 14
                c.setFont("Helvetica", 8)
            c.drawString(x, y, str(r.get("group_name", ""))[:18])
            c.drawString(x + 110, y, str(r.get("printer_name", ""))[:18])
            c.drawString(x + 220, y, str(r.get("brand", ""))[:12])
            c.drawString(x + 290, y, str(r.get("model", ""))[:16])
            c.drawString(x + 390, y, str(r.get("serial", ""))[:16])
            c.drawString(x + 490, y, str(r.get("location", ""))[:16])
            c.drawRightString(x + 600, y, str(r.get("reading_initial", 0)))
            c.drawRightString(x + 680, y, str(r.get("reading_final", 0)))
            c.drawRightString(x + 760, y, str(r.get("difference", 0)))
            y -= 12
        c.save()
        return Response(
            content=out.getvalue(),
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=relatorio-contadores.pdf"},
        )

    return {"ok": False, "error": "format must be csv, xlsx, or pdf"}


def _build_report_rows(since: Optional[str], until: Optional[str], group_by: str):
    return query_report(cfg.db_path, since=since, until=until, group_by=group_by)


@app.get("/report")
def report_export(
    format: str = "csv",
    group_by: str = "user",
    since: Optional[str] = None,
    until: Optional[str] = None,
):
    fmt = (format or "csv").lower()
    group = (group_by or "user").lower()
    rows = _build_report_rows(since, until, group)
    # Fallback: if there are no job logs, use printer counter deltas.
    if not rows and group in ("printer", "model"):
        by = "printer" if group == "printer" else "model"
        r_print = query_counter_report(cfg.db_path, since=since, until=until, group_by=by, metric="print")
        r_copy = query_counter_report(cfg.db_path, since=since, until=until, group_by=by, metric="copy")
        merged = {}
        for r in r_print:
            key = str(r.get("group_name", ""))
            merged[key] = {"group_name": key, "jobs": int(r.get("jobs", 0) or 0), "pages": int(r.get("difference", 0) or 0)}
        for r in r_copy:
            key = str(r.get("group_name", ""))
            if key not in merged:
                merged[key] = {"group_name": key, "jobs": int(r.get("jobs", 0) or 0), "pages": 0}
            merged[key]["pages"] += int(r.get("difference", 0) or 0)
            merged[key]["jobs"] = max(int(merged[key]["jobs"]), int(r.get("jobs", 0) or 0))
        rows = sorted(merged.values(), key=lambda x: int(x.get("pages", 0)), reverse=True)
    group_labels = {
        "user": "Usuário",
        "department": "Setor",
        "printer": "Impressora",
        "model": "Modelo",
    }

    # Relatório detalhado por impressora com leitura anterior/final e diferença.
    detailed_printer_rows = None
    if group == "printer":
        r_print = query_counter_report(cfg.db_path, since=since, until=until, group_by="printer", metric="print")
        r_copy = query_counter_report(cfg.db_path, since=since, until=until, group_by="printer", metric="copy")
        r_jobs = query_job_printer_readings(cfg.db_path, since=since, until=until)
        merged = {}
        for src in (r_print, r_copy):
            for r in src:
                key = str(r.get("printer_name") or r.get("group_name") or "").strip()
                if not key:
                    continue
                item = merged.setdefault(
                    key,
                    {
                        "impressora": key,
                        "serial": str(r.get("serial") or "").strip() or "Não definido",
                        "leitura_anterior": 0,
                        "leitura_final": 0,
                        "diferenca": 0,
                    },
                )
                item["leitura_anterior"] += int(r.get("reading_initial", 0) or 0)
                item["leitura_final"] += int(r.get("reading_final", 0) or 0)
                item["diferenca"] += int(r.get("difference", 0) or 0)
                if item["serial"] == "Não definido":
                    item["serial"] = str(r.get("serial") or "").strip() or "Não definido"

        # Fallback para impressoras de agent (ou outras) sem contador IP:
        # usa acumulado de jobs para leitura anterior/final.
        for r in r_jobs:
            key = str(r.get("printer_name") or "").strip()
            if not key or key in merged:
                continue
            merged[key] = {
                "impressora": key,
                "serial": str(r.get("serial") or "").strip() or "Não definido",
                "leitura_anterior": int(r.get("reading_initial", 0) or 0),
                "leitura_final": int(r.get("reading_final", 0) or 0),
                "diferenca": int(r.get("difference", 0) or 0),
            }

        leitura_anterior_data = "-"
        if since:
            try:
                from datetime import datetime, timedelta

                leitura_anterior_data = (datetime.fromisoformat(str(since)[:10]) - timedelta(days=1)).strftime("%Y-%m-%d")
            except Exception:
                leitura_anterior_data = str(since)
        leitura_final_data = str(until or "-")

        detailed_printer_rows = []
        for _, v in sorted(merged.items(), key=lambda kv: int(kv[1].get("diferenca", 0)), reverse=True):
            detailed_printer_rows.append(
                {
                    "impressora": v["impressora"],
                    "serial": v["serial"],
                    "leitura_anterior_data": leitura_anterior_data,
                    "leitura_anterior": v["leitura_anterior"],
                    "leitura_final_data": leitura_final_data,
                    "leitura_final": v["leitura_final"],
                    "diferenca": v["diferenca"],
                }
            )

    headers = ["group", "jobs", "pages"]
    if detailed_printer_rows is not None:
        headers = [
            "impressora",
            "serial",
            "leitura_anterior_data",
            "leitura_anterior",
            "leitura_final_data",
            "leitura_final",
            "diferenca",
        ]

    if fmt == "csv":
        import csv
        import io

        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(headers)
        if detailed_printer_rows is not None:
            for r in detailed_printer_rows:
                writer.writerow(
                    [
                        r.get("impressora", ""),
                        r.get("serial", ""),
                        r.get("leitura_anterior_data", ""),
                        r.get("leitura_anterior", 0),
                        r.get("leitura_final_data", ""),
                        r.get("leitura_final", 0),
                        r.get("diferenca", 0),
                    ]
                )
        else:
            for r in rows:
                writer.writerow([r.get("group_name", ""), r.get("jobs", 0), r.get("pages", 0)])
        data = buf.getvalue().encode("utf-8")
        return Response(
            content=data,
            media_type="text/csv",
            headers={"Content-Disposition": "attachment; filename=relatorio.csv"},
        )

    if fmt == "xlsx":
        import io
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "Relatório"
        ws.append(headers)
        if detailed_printer_rows is not None:
            for r in detailed_printer_rows:
                ws.append(
                    [
                        r.get("impressora", ""),
                        r.get("serial", ""),
                        r.get("leitura_anterior_data", ""),
                        r.get("leitura_anterior", 0),
                        r.get("leitura_final_data", ""),
                        r.get("leitura_final", 0),
                        r.get("diferenca", 0),
                    ]
                )
        else:
            for r in rows:
                ws.append([r.get("group_name", ""), r.get("jobs", 0), r.get("pages", 0)])
        out = io.BytesIO()
        wb.save(out)
        return Response(
            content=out.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=relatorio.xlsx"},
        )

    if fmt == "pdf":
        import io
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas

        page_size = landscape(A4) if detailed_printer_rows is not None else A4
        out = io.BytesIO()
        c = canvas.Canvas(out, pagesize=page_size)
        width, height = page_size
        x = 24 if detailed_printer_rows is not None else 40
        y = height - 50
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x, y, "Relatório de Impressão")
        y -= 20
        c.setFont("Helvetica", 10)
        c.drawString(x, y, f"Agrupar por: {group_labels.get(group, group)}")
        y -= 20
        c.drawString(x, y, f"Período: {since or '-'} até {until or '-'}")
        y -= 30
        if detailed_printer_rows is not None:
            col_printer = x
            col_serial = x + 175
            col_data_ant = x + 305
            col_leitura_ant = x + 410
            col_data_final = x + 520
            col_leitura_final = x + 635
            col_dif = x + 770
            c.setFont("Helvetica-Bold", 8)
            c.drawString(col_printer, y, "Impressora")
            c.drawString(col_serial, y, "Serial")
            c.drawString(col_data_ant, y, "Data ant.")
            c.drawRightString(col_leitura_ant, y, "Leitura ant.")
            c.drawString(col_data_final, y, "Data final")
            c.drawRightString(col_leitura_final, y, "Leitura final")
            c.drawRightString(col_dif, y, "Diferença")
            y -= 14
            c.setFont("Helvetica", 8)
            for r in detailed_printer_rows:
                if y < 60:
                    c.showPage()
                    y = height - 50
                    c.setFont("Helvetica-Bold", 8)
                    c.drawString(col_printer, y, "Impressora")
                    c.drawString(col_serial, y, "Serial")
                    c.drawString(col_data_ant, y, "Data ant.")
                    c.drawRightString(col_leitura_ant, y, "Leitura ant.")
                    c.drawString(col_data_final, y, "Data final")
                    c.drawRightString(col_leitura_final, y, "Leitura final")
                    c.drawRightString(col_dif, y, "Diferença")
                    y -= 14
                    c.setFont("Helvetica", 8)
                c.drawString(col_printer, y, str(r.get("impressora", ""))[:30])
                c.drawString(col_serial, y, str(r.get("serial", ""))[:22])
                c.drawString(col_data_ant, y, str(r.get("leitura_anterior_data", ""))[:10])
                c.drawRightString(col_leitura_ant, y, str(r.get("leitura_anterior", 0)))
                c.drawString(col_data_final, y, str(r.get("leitura_final_data", ""))[:10])
                c.drawRightString(col_leitura_final, y, str(r.get("leitura_final", 0)))
                c.drawRightString(col_dif, y, str(r.get("diferenca", 0)))
                y -= 12
        else:
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x, y, "Grupo")
            c.drawString(x + 260, y, "Jobs")
            c.drawString(x + 340, y, "Páginas")
            y -= 14
            c.setFont("Helvetica", 10)
            for r in rows:
                if y < 60:
                    c.showPage()
                    y = height - 50
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(x, y, "Grupo")
                    c.drawString(x + 260, y, "Jobs")
                    c.drawString(x + 340, y, "Páginas")
                    y -= 14
                    c.setFont("Helvetica", 10)
                c.drawString(x, y, str(r.get("group_name", ""))[:40])
                c.drawRightString(x + 300, y, str(r.get("jobs", 0)))
                c.drawRightString(x + 380, y, str(r.get("pages", 0)))
                y -= 12
        c.save()
        return Response(
            content=out.getvalue(),
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=relatorio.pdf"},
        )

    return {"ok": False, "error": "format must be csv, xlsx, or pdf"}


@app.get("/", response_class=HTMLResponse)
def home():
    def _fmt_ts(value: str) -> str:
        try:
            from datetime import datetime

            dt = datetime.fromisoformat(str(value))
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            return str(value or "")

    summary = query_summary(cfg.db_path, cfg.default_days)
    jobs = query_jobs(cfg.db_path, limit=50)

    totals = summary.get("totals", {})
    total_jobs = totals.get("jobs", 0)
    total_pages = totals.get("total_pages", 0)
    by_day_data = summary.get("by_day", [])
    top_users_data = summary.get("top_users", [])
    top_printers_data = summary.get("top_printers", [])

    if int(total_jobs or 0) == 0 and int(total_pages or 0) == 0:
        from datetime import datetime, timedelta

        since_day = (datetime.now() - timedelta(days=cfg.default_days)).strftime("%Y-%m-%d")
        until_day = datetime.now().strftime("%Y-%m-%d")
        rp = query_counter_report(cfg.db_path, since=since_day, until=until_day, group_by="printer", metric="print")
        rc = query_counter_report(cfg.db_path, since=since_day, until=until_day, group_by="printer", metric="copy")
        by_name = {}
        for r in rp:
            by_name[str(r.get("group_name") or "")] = int(r.get("difference", 0) or 0)
        for r in rc:
            key = str(r.get("group_name") or "")
            by_name[key] = by_name.get(key, 0) + int(r.get("difference", 0) or 0)

        top_printers_data = [
            {"printer": k, "pages": v}
            for k, v in sorted(by_name.items(), key=lambda x: x[1], reverse=True)
            if k
        ][:10]
        by_day_data = query_counter_daily(cfg.db_path, since=since_day, until=until_day)
        total_pages = sum(int(v) for v in by_name.values())
        total_jobs = total_pages
        top_users_data = []
    max_day_pages = max([d.get("pages", 0) for d in by_day_data], default=0) or 1

    chart_w = 760
    chart_h = 270
    pad_l, pad_r, pad_t, pad_b = 22, 18, 20, 44
    inner_w = max(10, chart_w - pad_l - pad_r)
    inner_h = max(10, chart_h - pad_t - pad_b)
    chart_points = []
    if by_day_data:
        n = len(by_day_data)
        for i, d in enumerate(by_day_data):
            v = float(d.get("pages", 0) or 0)
            bar_w = max(12.0, min(48.0, inner_w / max(1, n) * 0.65))
            if n == 1:
                x = pad_l + (inner_w / 2.0)
                bar_w = min(80.0, inner_w * 0.35)
            else:
                x = pad_l + (inner_w * i / (n - 1))
            y = pad_t + (inner_h * (1.0 - (v / max_day_pages)))
            chart_points.append((x, y, bar_w, str(d.get("day", ""))[5:], int(v)))
    chart_bars = "".join(
        f'<rect x="{(x - (w/2)):.2f}" y="{y:.2f}" width="{w:.2f}" height="{max(2.0, pad_t + inner_h - y):.2f}" rx="5" class="chart-bar"><title>{day} - {val} páginas</title></rect>'
        for x, y, w, day, val in chart_points
    )
    chart_xlabels = "".join(
        f'<text x="{x:.2f}" y="{chart_h - 10}" class="chart-xlabel" text-anchor="middle">{day}</text>'
        for x, _, _, day, _ in chart_points
    )
    chart_values = "".join(
        f'<text x="{x:.2f}" y="{max(10.0, y - 6):.2f}" class="chart-vlabel" text-anchor="middle">{val}</text>'
        for x, y, _, _, val in chart_points
    )
    total_period = int(sum(int(d.get("pages", 0) or 0) for d in by_day_data))
    avg_period = int(total_period / max(1, len(by_day_data))) if by_day_data else 0
    peak = max(by_day_data, key=lambda z: int(z.get("pages", 0) or 0), default={"day": "-", "pages": 0})
    daily_chart_svg = f"""
    <svg class="daily-chart" viewBox="0 0 {chart_w} {chart_h}" preserveAspectRatio="none" role="img" aria-label="Gráfico diário de páginas">
      <line x1="{pad_l}" y1="{pad_t + inner_h}" x2="{pad_l + inner_w}" y2="{pad_t + inner_h}" class="chart-axis" />
      <line x1="{pad_l}" y1="{pad_t}" x2="{pad_l}" y2="{pad_t + inner_h}" class="chart-axis" />
      <line x1="{pad_l}" y1="{pad_t + inner_h * 0.66:.2f}" x2="{pad_l + inner_w}" y2="{pad_t + inner_h * 0.66:.2f}" class="chart-grid" />
      <line x1="{pad_l}" y1="{pad_t + inner_h * 0.33:.2f}" x2="{pad_l + inner_w}" y2="{pad_t + inner_h * 0.33:.2f}" class="chart-grid" />
      {chart_bars}
      {chart_values}
      {chart_xlabels}
    </svg>
    """

    if not jobs:
        jobs = query_recent_counter_events(cfg.db_path, limit=50)

    rows = "".join(
        f"<tr><td>{_fmt_ts(j.get('timestamp',''))}</td><td>{j.get('user','')}</td><td>{j.get('printer','')}</td><td class=\"num\">{j.get('pages','')}</td><td class=\"num\">{j.get('copies','')}</td><td>{j.get('document','')}</td></tr>"
        for j in jobs
    )

    by_day = "".join(
        f"""
        <tr>
          <td>{d.get('day','')}</td>
          <td>
            <div class="bar-row">
              <div class="bar" style="width:{int((d.get('pages',0)/max_day_pages)*100)}%"></div>
            </div>
          </td>
          <td class="num">{d.get('pages','')}</td>
        </tr>
        """
        for d in by_day_data
    )

    top_users = "".join(
        f"<tr><td>{u.get('user','')}</td><td>{u.get('pages','')}</td></tr>"
        for u in top_users_data
    )

    top_printers = "".join(
        f"<tr><td>{p.get('printer','')}</td><td>{p.get('pages','')}</td></tr>"
        for p in top_printers_data
    )

    latest_counters = list_latest_counters(cfg.db_path)
    counter_rows = "".join(
        f"<tr><td>IP</td><td>{c.get('printer_name','')}</td><td>{c.get('ip','')}</td><td>{c.get('brand','')}</td><td>{c.get('model','')}</td><td class=\"num\">{c.get('total_print',0)}</td><td class=\"num\">{c.get('total_copy',0)}</td><td class=\"num\">{c.get('total_scan',0)}</td><td>{_fmt_ts(c.get('timestamp',''))}</td></tr>"
        for c in latest_counters
    )

    html = f"""
    <html>
      <head>
        <title>Print Server Dashboard</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <style>
          :root {{
            --bg: #0f1115;
            --panel: #151a22;
            --panel-2: #1b2230;
            --text: #e7edf5;
            --muted: #98a4b3;
            --accent: #51c1ff;
            --accent-2: #3ee6a5;
            --warning: #ffb65c;
            --border: #232c3b;
            --shadow: 0 12px 30px rgba(0,0,0,0.25);
          }}
          * {{ box-sizing: border-box; }}
          body {{
            margin: 0;
            font-family: "IBM Plex Sans", "Segoe UI", Arial, sans-serif;
            color: var(--text);
            background:
              radial-gradient(1200px 800px at 10% -10%, #233045 0%, transparent 60%),
              radial-gradient(900px 600px at 90% -20%, #1f3b3a 0%, transparent 65%),
              var(--bg);
          }}
          .wrap {{ max-width: 1200px; margin: 0 auto; padding: 28px 24px 40px; }}
          header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            margin-bottom: 24px;
          }}
          .title {{
            display: flex;
            flex-direction: column;
            gap: 6px;
          }}
          .title h1 {{
            margin: 0;
            font-size: 26px;
            letter-spacing: 0.3px;
          }}
          .subtitle {{
            color: var(--muted);
            font-size: 14px;
          }}
          .pill {{
            padding: 8px 12px;
            background: linear-gradient(120deg, #1d2635, #1a2030);
            border: 1px solid var(--border);
            border-radius: 999px;
            font-size: 12px;
            color: var(--muted);
          }}
          .cards {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 16px;
            margin-bottom: 22px;
          }}
          .card {{
            background: linear-gradient(160deg, var(--panel), var(--panel-2));
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 16px;
            box-shadow: var(--shadow);
          }}
          .card h3 {{ margin: 0 0 8px; font-size: 13px; color: var(--muted); }}
          .card .value {{ font-size: 26px; font-weight: 700; }}
          .card .trend {{ font-size: 12px; color: var(--muted); margin-top: 4px; }}
          .grid {{
            display: grid;
            grid-template-columns: 1.4fr 1fr;
            gap: 18px;
          }}
          .panel {{
            background: linear-gradient(160deg, var(--panel), var(--panel-2));
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 18px;
            box-shadow: var(--shadow);
          }}
          .panel h2 {{ margin: 0 0 12px; font-size: 16px; }}
          .panel h3 {{ margin: 12px 0 8px; font-size: 13px; color: var(--muted); }}
          table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
          }}
          th, td {{
            padding: 10px 8px;
            border-bottom: 1px solid var(--border);
          }}
          th {{ text-align: left; color: var(--muted); font-weight: 600; }}
          .num {{ text-align: right; }}
          .bar-row {{
            height: 8px;
            background: #111722;
            border-radius: 999px;
            overflow: hidden;
            border: 1px solid #1f2a3a;
          }}
          .bar {{
            height: 100%;
            background: linear-gradient(90deg, var(--accent), var(--accent-2));
            border-radius: 999px;
          }}
          .daily-chart-wrap {{
            margin-top: 14px;
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 14px 14px 10px;
            background: linear-gradient(180deg, rgba(255,255,255,0.01), rgba(0,0,0,0.06));
          }}
          .daily-chart {{
            width: 100%;
            height: 270px;
            display: block;
          }}
          .chart-meta {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-top: 8px;
          }}
          .chart-pill {{
            font-size: 11px;
            color: var(--muted);
            border: 1px solid var(--border);
            border-radius: 999px;
            padding: 4px 10px;
            background: rgba(255,255,255,0.02);
          }}
          .chart-axis {{ stroke: #3a4559; stroke-width: 1; }}
          .chart-grid {{ stroke: #2a3446; stroke-width: 1; stroke-dasharray: 2 5; }}
          .chart-bar {{ fill: url(#gradArea); stroke: rgba(81,193,255,0.6); stroke-width: 0.8; }}
          .chart-vlabel {{
            fill: #b9c5d6;
            font-size: 11px;
            font-family: "IBM Plex Sans", "Segoe UI", Arial, sans-serif;
          }}
          .chart-xlabel {{
            fill: var(--muted);
            font-size: 12px;
            font-weight: 600;
            font-family: "IBM Plex Sans", "Segoe UI", Arial, sans-serif;
          }}
          .table-wrap {{
            max-height: 360px;
            overflow: auto;
            border: 1px solid var(--border);
            border-radius: 12px;
          }}
          .badge {{
            display: inline-block;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 11px;
            color: #0b1220;
            background: var(--accent);
            font-weight: 700;
          }}
          .form-row {{
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 10px;
          }}
          .form-row input, .form-row select {{
            width: 100%;
            padding: 8px 10px;
            background: #0f1622;
            color: var(--text);
            border: 1px solid var(--border);
            border-radius: 8px;
          }}
          .btn {{
            background: linear-gradient(135deg, var(--accent), var(--accent-2));
            border: none;
            color: #06121a;
            padding: 10px 14px;
            border-radius: 10px;
            font-weight: 700;
            cursor: pointer;
          }}
          @media (max-width: 1000px) {{
            .cards {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            .grid {{ grid-template-columns: 1fr; }}
          }}
          @media (max-width: 600px) {{
            .cards {{ grid-template-columns: 1fr; }}
          }}
        </style>
      </head>
      <body>
        <div class="wrap">
          <header>
            <div class="title">
              <h1>Print Server Dashboard</h1>
              <div class="subtitle">Servidor de Impressão - Em tempo Real - Dashbord</div>
            </div>
            <div style="display:flex; gap:10px; align-items:center;">
              <a class="pill" href="/configuracoes">Configurações</a>
              <a class="pill" href="/relatorios">Relatórios</a>
              <div class="pill">Servidor ativo • Porta {cfg.server_port}</div>
            </div>
          </header>

          <section class="cards">
            <div class="card">
              <h3>Jobs Processados</h3>
              <div class="value">{total_jobs}</div>
              <div class="trend">Últimos {cfg.default_days} dias</div>
            </div>
            <div class="card">
              <h3>Páginas Totais</h3>
              <div class="value">{total_pages}</div>
              <div class="trend">Inclui cópias</div>
            </div>
            <div class="card">
              <h3>Top Usuário</h3>
              <div class="value">{(top_users_data[:1] or [{}])[0].get("user","-")}</div>
              <div class="trend">Maior volume</div>
            </div>
            <div class="card">
              <h3>Top Impressora</h3>
              <div class="value">{(top_printers_data[:1] or [{}])[0].get("printer","-")}</div>
              <div class="trend">Maior volume</div>
            </div>
          </section>

          <section class="grid">
            <div class="panel">
              <h2>Páginas por Dia</h2>
              <table>
                <tr><th>Dia</th><th>Volume</th><th class="num">Páginas</th></tr>
                {by_day}
              </table>
              <div class="daily-chart-wrap">
                <svg width="0" height="0" style="position:absolute">
                  <defs>
                    <linearGradient id="gradArea" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stop-color="#51c1ff" stop-opacity="0.30" />
                      <stop offset="100%" stop-color="#51c1ff" stop-opacity="0.02" />
                    </linearGradient>
                  </defs>
                </svg>
                {daily_chart_svg}
                <div class="chart-meta">
                  <span class="chart-pill">Total no período: {total_period}</span>
                  <span class="chart-pill">Média por dia: {avg_period}</span>
                  <span class="chart-pill">Pico: {peak.get('day','-')} ({peak.get('pages',0)} páginas)</span>
                </div>
              </div>
            </div>
            <div class="panel">
              <h2>Ranking</h2>
              <div class="table-wrap">
                <table>
                  <tr><th>Usuário</th><th class="num">Páginas</th></tr>
                  {top_users}
                </table>
              </div>
              <div style="height:12px;"></div>
              <div class="table-wrap">
                <table>
                  <tr><th>Impressora</th><th class="num">Páginas</th></tr>
                  {top_printers}
                </table>
              </div>
            </div>
          </section>

          <section class="panel" style="margin-top:18px;">
            <h2>Impressoras IP (Contadores)</h2>
            <div class="subtitle">Adicionar IPs e URL de contadores por modelo</div>
            <form method="post" action="/api/printer-sources" onsubmit="return submitJson(event, this);">
              <div class="form-row">
                <input name="name" placeholder="nome (ex: Brother-90)" required />
                <input name="ip" placeholder="IP (ex: 192.168.150.90)" required />
                <input name="brand" placeholder="marca (ex: Brother)" />
              </div>
              <div style="height:10px;"></div>
              <div class="form-row">
                <input name="model" placeholder="modelo (ex: DCP-8080DN)" />
                <input name="counter_url" placeholder="URL de contadores (ex: http://IP/etc/mnt_info.html?kind=item)" required />
                <button class="btn" type="submit">Salvar</button>
              </div>
              <div style="height:10px;"></div>
              <div class="subtitle">Brother: http://IP/etc/mnt_info.html?kind=item | Samsung: http://IP/sws/index.html</div>
            </form>
            <div style="height:12px;"></div>
            <button class="btn" onclick="return scanCounters();">Atualizar contadores</button>
            <div style="height:12px;"></div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr><th>Origem</th><th>Impressora</th><th>IP</th><th>Marca/Host</th><th>Serial</th><th class="num">Impressões</th><th class="num">Cópias</th><th class="num">Scans</th><th>Atualizado</th></tr>
                </thead>
                <tbody id="counterBody">
                  {counter_rows}
                </tbody>
              </table>
            </div>
          </section>

          <section class="panel" style="margin-top:18px;">
            <h2>Jobs Recentes</h2>
            <div class="table-wrap">
              <table>
                <tr><th>Data/Hora</th><th>Usuário</th><th>Impressora</th><th class="num">Páginas</th><th class="num">Cópias</th><th>Documento</th></tr>
                {rows}
              </table>
            </div>
          </section>
        </div>
        <script>
          async function submitJson(ev, form) {{
            ev.preventDefault();
            const data = Object.fromEntries(new FormData(form).entries());
            const res = await fetch(form.action, {{
              method: "POST",
              headers: {{ "Content-Type": "application/json" }},
              body: JSON.stringify(data)
            }});
            const out = await res.json();
            if (!out.ok) {{
              alert(out.error || "Erro ao salvar");
              return false;
            }}
            form.reset();
            alert("Salvo com sucesso");
            return false;
          }}
          async function scanCounters() {{
            const res = await fetch("/api/printer-scan", {{ method: "POST" }});
            const out = await res.json();
            if (!out.ok) {{
              alert(out.error || "Erro ao atualizar contadores");
              return false;
            }}
            await refreshCounters();
            return false;
          }}
          async function refreshCounters() {{
            const fmtTs = (v) => {{
              if (!v) return "";
              const d = new Date(v);
              if (Number.isNaN(d.getTime())) return String(v);
              return d.toLocaleString("pt-BR", {{ hour12: false }});
            }};
            const [resCounters, resAgents] = await Promise.all([
              fetch("/api/printer-counters"),
              fetch("/api/agents")
            ]);
            const counters = await resCounters.json();
            const agents = await resAgents.json();

            const ipRows = counters.map(c => ({{
              origem: "IP",
              printer_name: c.printer_name || "",
              ip: c.ip || "",
              owner: c.brand || "",
              model: c.serial || c.model || "",
              total_print: c.total_print ?? 0,
              total_copy: c.total_copy ?? 0,
              total_scan: c.total_scan ?? 0,
              timestamp: c.timestamp || ""
            }}));

            const agentRows = agents.map(a => ({{
              origem: "Agent",
              printer_name: a.printer_name || "",
              ip: a.ip || "",
              owner: a.host || "",
              model: a.serial || a.printer_model || "",
              total_print: "-",
              total_copy: "-",
              total_scan: "-",
              timestamp: a.updated_at || ""
            }}));

            const data = [...ipRows, ...agentRows].sort((x, y) =>
              String(x.printer_name || "").localeCompare(String(y.printer_name || ""))
            );

            const body = document.getElementById("counterBody");
            if (!body) return;
            body.innerHTML = data.map(c => `
              <tr>
                <td>${{c.origem || ""}}</td>
                <td>${{c.printer_name || ""}}</td>
                <td>${{c.ip || ""}}</td>
                <td>${{c.owner || ""}}</td>
                <td>${{c.model || ""}}</td>
                <td class="num">${{c.total_print}}</td>
                <td class="num">${{c.total_copy}}</td>
                <td class="num">${{c.total_scan}}</td>
                <td>${{fmtTs(c.timestamp)}}</td>
              </tr>
            `).join("");
          }}
          refreshCounters();
          setInterval(refreshCounters, 5000);
        </script>
      </body>
    </html>
    """
    return HTMLResponse(html)


@app.get("/configuracoes", response_class=HTMLResponse)
def settings_page():
    html = """
    <html>
      <head>
        <title>Configurações - Print Server Dashboard</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <style>
          body { font-family: Segoe UI, Arial, sans-serif; margin: 0; background: #f4f6fb; color: #152238; }
          .wrap { max-width: 1200px; margin: 0 auto; padding: 20px; }
          .top { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
          .card { background: #fff; border: 1px solid #dbe3ef; border-radius: 12px; padding: 14px; margin-top: 14px; }
          .row { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; }
          input, select, button { padding: 8px; border-radius: 8px; border: 1px solid #cbd5e1; }
          button { cursor: pointer; background: #0ea5e9; color: white; border: none; font-weight: 600; }
          .btn-alt { background: #334155; }
          .btn-danger { background: #b91c1c; }
          table { width: 100%; border-collapse: collapse; font-size: 13px; }
          th, td { padding: 8px; border-bottom: 1px solid #e5e7eb; text-align: left; }
          .muted { color: #64748b; font-size: 12px; }
          .actions { display: flex; gap: 6px; flex-wrap: wrap; }
          @media (max-width: 900px) { .row { grid-template-columns: 1fr 1fr; } }
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="top">
            <h2>Configurações</h2>
            <a href="/">Voltar ao dashboard</a>
          </div>

          <div class="card">
            <h3>Setor por Usuário</h3>
            <form id="userDeptForm" onsubmit="return saveUserDept(event);">
              <div class="row">
                <input id="ud_user" placeholder="Usuário (ex: joao)" required />
                <input id="ud_department" placeholder="Setor (ex: Financeiro)" required />
                <select id="ud_source">
                  <option value="manual">manual</option>
                  <option value="ad">ad</option>
                </select>
                <button type="submit">Salvar</button>
              </div>
            </form>
            <table style="margin-top:10px;">
              <thead><tr><th>Usuário</th><th>Setor</th><th>Origem</th><th>Atualizado</th></tr></thead>
              <tbody id="userDeptRows"></tbody>
            </table>
          </div>

          <div class="card">
            <h3>Modelo da Impressora</h3>
            <form id="printerModelForm" onsubmit="return savePrinterModel(event);">
              <div class="row">
                <input id="pm_printer" placeholder="Impressora/Fila" required />
                <input id="pm_model" placeholder="Modelo" required />
                <select id="pm_source">
                  <option value="manual">manual</option>
                  <option value="windows">windows</option>
                  <option value="agent">agent</option>
                </select>
                <button type="submit">Salvar</button>
              </div>
            </form>
            <table style="margin-top:10px;">
              <thead><tr><th>Impressora</th><th>Modelo</th><th>Origem</th><th>Atualizado</th></tr></thead>
              <tbody id="printerModelRows"></tbody>
            </table>
          </div>

          <div class="card">
            <h3>Setores e Vinculo de Impressoras</h3>
            <form id="departmentForm" onsubmit="return saveDepartment(event);">
              <input type="hidden" id="dep_id" />
              <div class="row">
                <input id="dep_name" placeholder="Nome do setor" required />
                <button type="submit">Salvar setor</button>
                <button type="button" class="btn-alt" onclick="clearDepartment()">Cancelar</button>
                <span></span>
              </div>
            </form>
            <table style="margin-top:10px;">
              <thead><tr><th>ID</th><th>Setor</th><th>Atualizado</th><th>Ações</th></tr></thead>
              <tbody id="depRows"></tbody>
            </table>

            <form id="printerDeptForm" onsubmit="return savePrinterDept(event);" style="margin-top:12px;">
              <div class="row">
                <input id="pd_printer" list="printerKnownList" placeholder="Impressora/Fila" required />
                <datalist id="printerKnownList"></datalist>
                <select id="pd_department" required></select>
                <button type="submit">Vincular impressora ao setor</button>
                <span></span>
              </div>
            </form>
            <table style="margin-top:10px;">
              <thead><tr><th>Impressora</th><th>Setor</th><th>Atualizado</th><th>Ações</th></tr></thead>
              <tbody id="printerDeptRows"></tbody>
            </table>
          </div>

          <div class="card">
            <h3>Impressoras IP</h3>
            <form id="ipForm" onsubmit="return saveIp(event);">
              <input type="hidden" id="ip_id" />
              <div class="row">
                <input id="ip_name" placeholder="Nome" required />
                <input id="ip_ip" placeholder="IP" required />
                <input id="ip_brand" placeholder="Marca" />
                <input id="ip_model" placeholder="Modelo" />
              </div>
              <div class="row" style="margin-top:8px;">
                <input id="ip_serial" placeholder="Serial" />
                <input id="ip_location" placeholder="Local" />
                <input id="ip_url" placeholder="URL contadores" required />
                <button type="submit">Salvar</button>
                <button type="button" class="btn-alt" onclick="clearIp()">Cancelar</button>
                <button type="button" class="btn-alt" onclick="refreshAll()">Atualizar listas</button>
              </div>
            </form>
            <div class="muted" style="margin-top:6px;">Brother: http://IP/etc/mnt_info.html?kind=item | Samsung: http://IP/sws/index.html</div>
            <table style="margin-top:10px;">
              <thead><tr><th>Nome</th><th>IP</th><th>Marca</th><th>Modelo</th><th>Serial</th><th>Local</th><th>URL</th><th>Atualizado</th><th>Ações</th></tr></thead>
              <tbody id="ipRows"></tbody>
            </table>
          </div>

          <div class="card">
            <h3>Agents</h3>
            <form id="agentForm" onsubmit="return saveAgent(event);">
              <input type="hidden" id="ag_id" />
              <div class="row">
                <input id="ag_host" placeholder="Computador" required />
                <input id="ag_printer" placeholder="Impressora" required />
                <input id="ag_model" placeholder="Modelo impressora" />
                <input id="ag_serial" placeholder="Serial impressora" />
                <input id="ag_ip" placeholder="IP cliente" />
              </div>
              <div class="row" style="margin-top:8px;">
                <input id="ag_location" placeholder="Local" />
                <input id="ag_ver" placeholder="Versão do agent" />
                <button type="submit">Salvar</button>
                <button type="button" class="btn-alt" onclick="clearAgent()">Cancelar</button>
                <span></span>
              </div>
            </form>
            <table style="margin-top:10px;">
              <thead><tr><th>Agent ID</th><th>Host</th><th>Impressora</th><th>Modelo</th><th>Serial</th><th>Local</th><th>IP</th><th>Versão</th><th>Atualizado</th><th>Ações</th></tr></thead>
              <tbody id="agentRows"></tbody>
            </table>
          </div>

          <div class="card">
            <h3>Exclusões de Relatório</h3>
            <form id="exForm" onsubmit="return saveExclusion(event);">
              <div class="row">
                <select id="ex_kind">
                  <option value="printer">Impressora</option>
                  <option value="agent">Agent</option>
                </select>
                <input id="ex_value" placeholder="Valor (nome impressora ou agent_id)" required />
                <input id="ex_note" placeholder="Observação" />
                <button type="submit">Adicionar</button>
              </div>
            </form>
            <table style="margin-top:10px;">
              <thead><tr><th>Tipo</th><th>Valor</th><th>Obs</th><th>Atualizado</th><th>Ações</th></tr></thead>
              <tbody id="exRows"></tbody>
            </table>
            <div class="muted">Itens nesta lista não entram nos relatórios nem nos totais do dashboard.</div>
          </div>
        </div>
        <script>
          const esc = (v) => String(v || "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;");
          const fmtTs = (v) => {
            if (!v) return "";
            const d = new Date(v);
            if (Number.isNaN(d.getTime())) return String(v);
            return d.toLocaleString("pt-BR", { hour12: false });
          };

          async function j(url, opt) {
            const r = await fetch(url, opt);
            return await r.json();
          }

          let ipRowsData = [];
          let agentRowsData = [];
          let depRowsData = [];
          let printerDeptRowsData = [];
          let knownPrintersData = [];

          async function loadUserDepartments(){
            const rows = await j("/api/user-departments");
            document.getElementById("userDeptRows").innerHTML = rows.map(r => `
              <tr><td>${esc(r.user)}</td><td>${esc(r.department)}</td><td>${esc(r.source)}</td><td>${fmtTs(r.updated_at)}</td></tr>
            `).join("");
          }
          async function saveUserDept(ev){
            ev.preventDefault();
            const payload = {user:ud_user.value, department:ud_department.value, source:ud_source.value};
            const out = await j("/api/user-departments", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            userDeptForm.reset();
            await loadUserDepartments();
            return false;
          }

          async function loadPrinterModels(){
            const rows = await j("/api/printer-models");
            document.getElementById("printerModelRows").innerHTML = rows.map(r => `
              <tr><td>${esc(r.printer)}</td><td>${esc(r.model)}</td><td>${esc(r.source)}</td><td>${fmtTs(r.updated_at)}</td></tr>
            `).join("");
          }
          async function savePrinterModel(ev){
            ev.preventDefault();
            const payload = {printer:pm_printer.value, model:pm_model.value, source:pm_source.value};
            const out = await j("/api/printer-models", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            printerModelForm.reset();
            await loadPrinterModels();
            await loadKnownPrinters();
            return false;
          }

          async function loadDepartments(){
            depRowsData = await j("/api/departments");
            const options = ['<option value="">Selecione o setor</option>'].concat(
              depRowsData.map(d => `<option value="${d.id}">${esc(d.name)}</option>`)
            ).join("");
            pd_department.innerHTML = options;
            document.getElementById("depRows").innerHTML = depRowsData.map(r => `
              <tr>
                <td>${r.id}</td><td>${esc(r.name)}</td><td>${fmtTs(r.updated_at)}</td>
                <td class="actions">
                  <button type="button" onclick="editDepartment(${r.id})">Editar</button>
                  <button type="button" class="btn-danger" onclick="delDepartment(${r.id})">Excluir</button>
                </td>
              </tr>
            `).join("");
          }
          function editDepartment(id){
            const r = depRowsData.find(x => Number(x.id) === Number(id));
            if(!r) return;
            dep_id.value = r.id || "";
            dep_name.value = r.name || "";
          }
          function clearDepartment(){ dep_id.value = ""; dep_name.value = ""; }
          async function saveDepartment(ev){
            ev.preventDefault();
            const id = String(dep_id.value || "").trim();
            const payload = {name: dep_name.value};
            const out = id
              ? await j("/api/departments/"+id, {method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)})
              : await j("/api/departments", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            clearDepartment();
            await loadDepartments();
            await loadPrinterDepartments();
            return false;
          }
          async function delDepartment(id){
            if(!confirm("Excluir setor? Isso remove os vinculos de impressoras desse setor.")) return;
            await j("/api/departments/"+id, {method:"DELETE"});
            await loadDepartments();
            await loadPrinterDepartments();
          }

          async function loadKnownPrinters(){
            knownPrintersData = await j("/api/printers-known");
            document.getElementById("printerKnownList").innerHTML = knownPrintersData.map(p => `<option value="${esc(p)}"></option>`).join("");
          }

          async function loadPrinterDepartments(){
            printerDeptRowsData = await j("/api/printer-departments");
            document.getElementById("printerDeptRows").innerHTML = printerDeptRowsData.map(r => `
              <tr>
                <td>${esc(r.printer)}</td><td>${esc(r.department_name || "")}</td><td>${fmtTs(r.updated_at)}</td>
                <td class="actions"><button type="button" class="btn-danger" onclick="delPrinterDept('${encodeURIComponent(r.printer || "")}')">Desvincular</button></td>
              </tr>
            `).join("");
          }
          async function savePrinterDept(ev){
            ev.preventDefault();
            const payload = {printer: pd_printer.value, department_id: Number(pd_department.value || 0)};
            if(!payload.printer || !payload.department_id){ alert("Preencha impressora e setor."); return false; }
            const out = await j("/api/printer-departments", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            printerDeptForm.reset();
            await loadPrinterDepartments();
            return false;
          }
          async function delPrinterDept(encPrinter){
            const printer = decodeURIComponent(encPrinter || "");
            if(!printer) return;
            if(!confirm("Desvincular impressora do setor?")) return;
            await j("/api/printer-departments?printer="+encodeURIComponent(printer), {method:"DELETE"});
            await loadPrinterDepartments();
          }

          async function loadIps() {
            ipRowsData = await j("/api/printer-sources");
            document.getElementById("ipRows").innerHTML = ipRowsData.map(r => `
              <tr>
                <td>${esc(r.name)}</td><td>${esc(r.ip)}</td><td>${esc(r.brand)}</td><td>${esc(r.model)}</td><td>${esc(r.serial)}</td><td>${esc(r.location)}</td><td>${esc(r.counter_url)}</td><td>${fmtTs(r.updated_at)}</td>
                <td class="actions">
                  <button type="button" onclick="editIp(${r.id})">Editar</button>
                  <button type="button" onclick="testIp(${r.id})">Testar</button>
                  <button type="button" class="btn-danger" onclick="delIp(${r.id})">Excluir</button>
                </td>
              </tr>
            `).join("");
          }

          function editIp(id){
            const r = ipRowsData.find(x => Number(x.id) === Number(id));
            if(!r) return;
            ip_id.value=r.id || "";
            ip_name.value=r.name || "";
            ip_ip.value=r.ip || "";
            ip_brand.value=r.brand || "";
            ip_model.value=r.model || "";
            ip_serial.value=r.serial || "";
            ip_location.value=r.location || "";
            ip_url.value=r.counter_url || "";
          }

          async function saveIp(ev){
            ev.preventDefault();
            const payload = { id: ip_id.value || undefined, name: ip_name.value, ip: ip_ip.value, brand: ip_brand.value, model: ip_model.value, serial: ip_serial.value, location: ip_location.value, counter_url: ip_url.value };
            const out = await j("/api/printer-sources", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            clearIp(); await loadIps(); return false;
          }
          function clearIp(){ ipForm.reset(); ip_id.value=""; }
          async function delIp(id){ if(!confirm("Excluir impressora IP?")) return; await j("/api/printer-sources/"+id, {method:"DELETE"}); await loadIps(); }
          async function testIp(id){ const out = await j("/api/printer-sources/"+id+"/test", {method:"POST"}); alert(out.ok ? "Teste OK" : ("Falha: "+(out.error||"erro"))); }

          async function loadAgents(){
            agentRowsData = await j("/api/agents");
            document.getElementById("agentRows").innerHTML = agentRowsData.map(r => `
              <tr>
                <td>${esc(r.agent_id)}</td><td>${esc(r.host)}</td><td>${esc(r.printer_name)}</td><td>${esc(r.printer_model)}</td><td>${esc(r.serial)}</td><td>${esc(r.location)}</td><td>${esc(r.ip)}</td><td>${esc(r.version)}</td><td>${fmtTs(r.updated_at)}</td>
                <td class="actions">
                  <button type="button" onclick="editAgent('${encodeURIComponent(r.agent_id || "")}')">Editar</button>
                  <button type="button" class="btn-danger" onclick="delAgent('${encodeURIComponent(r.agent_id || "")}')">Excluir</button>
                </td>
              </tr>
            `).join("");
          }
          function editAgent(encId){
            const id = decodeURIComponent(encId || "");
            const r = agentRowsData.find(x => String(x.agent_id || "") === id);
            if(!r) return;
            ag_id.value = r.agent_id || "";
            ag_host.value = r.host || "";
            ag_printer.value = r.printer_name || "";
            ag_model.value = r.printer_model || "";
            ag_serial.value = r.serial || "";
            ag_location.value = r.location || "";
            ag_ip.value = r.ip || "";
            ag_ver.value = r.version || "";
          }
          function clearAgent(){ agentForm.reset(); ag_id.value=""; }
          async function saveAgent(ev){
            ev.preventDefault();
            if(!ag_id.value){ alert("Edição de agent exige selecionar um agent da lista."); return false; }
            const payload = {host:ag_host.value, printer_name:ag_printer.value, printer_model:ag_model.value, serial:ag_serial.value, location:ag_location.value, ip:ag_ip.value, version:ag_ver.value};
            const out = await j("/api/agents/"+encodeURIComponent(ag_id.value), {method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload)});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            clearAgent(); await loadAgents(); return false;
          }
          async function delAgent(encId){ const id=decodeURIComponent(encId||""); if(!confirm("Excluir agent?")) return; await j("/api/agents/"+encodeURIComponent(id), {method:"DELETE"}); await loadAgents(); }

          async function loadExclusions(){
            const rows = await j("/api/exclusions");
            document.getElementById("exRows").innerHTML = rows.map(r => `
              <tr>
                <td>${esc(r.kind)}</td><td>${esc(r.value)}</td><td>${esc(r.note)}</td><td>${fmtTs(r.updated_at)}</td>
                <td><button type="button" class="btn-danger" onclick="delEx('${esc(r.kind)}','${esc(r.value)}')">Remover</button></td>
              </tr>
            `).join("");
          }
          async function saveExclusion(ev){
            ev.preventDefault();
            const out = await j("/api/exclusions", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({kind:ex_kind.value, value:ex_value.value, note:ex_note.value})});
            if(!out.ok){ alert(out.error || "Erro"); return false; }
            exForm.reset(); await loadExclusions(); return false;
          }
          async function delEx(kind, value){ await j(`/api/exclusions?kind=${encodeURIComponent(kind)}&value=${encodeURIComponent(value)}`, {method:"DELETE"}); await loadExclusions(); }

          async function refreshAll(){
            await Promise.all([
              loadUserDepartments(),
              loadPrinterModels(),
              loadDepartments(),
              loadKnownPrinters(),
              loadPrinterDepartments(),
              loadIps(),
              loadAgents(),
              loadExclusions(),
            ]);
          }
          refreshAll();
        </script>
      </body>
    </html>
    """
    return HTMLResponse(html)


@app.get("/relatorios", response_class=HTMLResponse)
def reports_page():
    html = """
    <html>
      <head>
        <title>Relatórios - Print Server Dashboard</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <style>
          body { font-family: Segoe UI, Arial, sans-serif; margin: 0; background: #f4f6fb; color: #152238; }
          .wrap { max-width: 1100px; margin: 0 auto; padding: 20px; }
          .top { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 14px; }
          .links { display: flex; gap: 10px; }
          .pill { padding: 8px 12px; border-radius: 999px; border: 1px solid #dbe3ef; background: #fff; color: #1e293b; text-decoration: none; font-size: 13px; }
          .card { background: #fff; border: 1px solid #dbe3ef; border-radius: 12px; padding: 14px; margin-top: 14px; }
          .row { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
          input, select, button { padding: 8px; border-radius: 8px; border: 1px solid #cbd5e1; }
          button { cursor: pointer; background: #0ea5e9; color: white; border: none; font-weight: 600; }
          @media (max-width: 900px) { .row { grid-template-columns: 1fr; } }
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="top">
            <h2>Relatórios</h2>
            <div class="links">
              <a class="pill" href="/">Dashboard</a>
              <a class="pill" href="/configuracoes">Configurações</a>
            </div>
          </div>

          <div class="card">
            <h3>Relatório de Impressões</h3>
            <form method="get" action="/report">
              <div class="row">
                <select name="group_by">
                  <option value="user">Usuário</option>
                  <option value="department">Setor</option>
                  <option value="printer">Impressora</option>
                  <option value="model">Modelo</option>
                </select>
                <input type="date" name="since" />
                <input type="date" name="until" />
              </div>
              <div style="height:10px;"></div>
              <div class="row">
                <select name="format">
                  <option value="csv">CSV</option>
                  <option value="xlsx">Excel</option>
                  <option value="pdf">PDF</option>
                </select>
                <button type="submit">Gerar relatório</button>
                <div></div>
              </div>
            </form>
          </div>

          <div class="card">
            <h3>Relatório de Contadores (Cópias/Scans)</h3>
            <form method="get" action="/report-counters">
              <div class="row">
                <select name="group_by">
                  <option value="printer">Impressora</option>
                  <option value="brand">Marca</option>
                  <option value="model">Modelo</option>
                  <option value="serial">Serial</option>
                </select>
                <select name="metric">
                  <option value="print">Impressões</option>
                  <option value="copy">Cópias</option>
                  <option value="scan">Scans</option>
                </select>
                <input type="date" name="since" />
              </div>
              <div style="height:10px;"></div>
              <div class="row">
                <input type="date" name="until" />
                <select name="format">
                  <option value="csv">CSV</option>
                  <option value="xlsx">Excel</option>
                  <option value="pdf">PDF</option>
                </select>
                <button type="submit">Gerar relatório</button>
              </div>
            </form>
          </div>
        </div>
      </body>
    </html>
    """
    return HTMLResponse(html)
