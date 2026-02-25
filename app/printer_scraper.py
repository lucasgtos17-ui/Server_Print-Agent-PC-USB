import re
from typing import Dict, Optional
from urllib.parse import urlparse

import requests


def _to_int(value: Optional[str]) -> int:
    if not value:
        return 0
    try:
        return int(re.sub(r"[^0-9]", "", value))
    except Exception:
        return 0


def _find_label_value(text: str, label: str) -> int:
    pattern = re.compile(
        label + r"[\s\S]*?</td>\s*<td[^>]*>\s*([0-9.,]+)(?:\s*pages)?",
        re.IGNORECASE,
    )
    m = pattern.search(text)
    return _to_int(m.group(1)) if m else 0


def _find_brother_dd_value(text: str, label: str) -> int:
    pattern = re.compile(
        r"<DD>\s*" + label + r"\s*</DD>\s*</TD>\s*<TD[^>]*>\s*([0-9.,]+)",
        re.IGNORECASE,
    )
    m = pattern.search(text)
    return _to_int(m.group(1)) if m else 0


def parse_brother_counters(html: str) -> Dict[str, int]:
    total_copy = _find_brother_dd_value(html, r"Copy")
    total_print = _find_brother_dd_value(html, r"Print")

    scan_adf = _find_brother_dd_value(html, r"ADF\(SX\)")
    scan_adf_dx = _find_brother_dd_value(html, r"ADF\(DX\)")
    scan_flatbed = _find_brother_dd_value(html, r"Flatbed")

    if total_copy == 0:
        total_copy = _find_label_value(html, r"Copy")
    if total_print == 0:
        total_print = _find_label_value(html, r"Print")
    if scan_adf == 0 and scan_adf_dx == 0 and scan_flatbed == 0:
        scan_adf = _find_label_value(html, r"ADF\(SX\)")
        scan_adf_dx = _find_label_value(html, r"ADF\(DX\)")
        scan_flatbed = _find_label_value(html, r"Flatbed")

    return {
        "print": total_print,
        "copy": total_copy,
        "scan": scan_adf + scan_adf_dx + scan_flatbed,
    }


def parse_samsung_counters(html: str) -> Dict[str, int]:
    total_print = 0
    total_copy = 0
    total_scan = 0

    row_pattern = re.compile(r"Total\s+de\s+impress[oõ]es[\s\S]*?</tr>", re.IGNORECASE)
    row = row_pattern.search(html)
    if row:
        numbers = re.findall(r">\s*([0-9.,]+)\s*<", row.group(0))
        if len(numbers) >= 2:
            total_print = _to_int(numbers[0])
            total_copy = _to_int(numbers[1])

    if total_print == 0:
        total_print = _find_label_value(html, r"Imprimir")
    if total_copy == 0:
        total_copy = _find_label_value(html, r"Copiar")

    total_scan = _find_label_value(html, r"Digitalizar")
    if total_scan == 0:
        total_scan = _find_label_value(html, r"Scan")

    return {
        "print": total_print,
        "copy": total_copy,
        "scan": total_scan,
    }


def _find_js_key_int(text: str, key: str) -> int:
    m = re.search(rf"\b{re.escape(key)}\b\s*:\s*([0-9.,]+)", text, flags=re.IGNORECASE)
    return _to_int(m.group(1)) if m else 0


def parse_samsung_jsonlike_counters(text: str) -> Dict[str, int]:
    total_print = _find_js_key_int(text, "GXI_BILLING_PRINT_TOTAL_IMP_CNT")
    total_copy = _find_js_key_int(text, "GXI_BILLING_COPY_TOTAL_IMP_CNT")
    total_scan = _find_js_key_int(text, "GXI_BILLING_SEND_TO_TOTAL_CNT")
    if total_scan == 0:
        total_scan = _find_js_key_int(text, "GXI_BILLING_SEND_TOTAL_CNT")

    return {
        "print": total_print,
        "copy": total_copy,
        "scan": total_scan,
    }


def _build_samsung_candidate_urls(counter_url: str):
    parsed = urlparse(counter_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    return [
        f"{base}/sws/app/information/counters/counters.json",
        f"{base}/sws/app/information/counters/counters.html",
        f"{base}/sws/index.html",
    ]


def fetch_counters(counter_url: str, brand: str) -> Dict[str, int]:
    headers = {"User-Agent": "PrintDashboard/1.0"}

    html = ""
    last_err = None
    for _ in range(2):
        try:
            resp = requests.get(counter_url, timeout=(3, 15), headers=headers)
            resp.raise_for_status()
            html = resp.text
            break
        except Exception as e:
            last_err = e
            continue

    if not html and last_err:
        raise last_err

    brand_low = (brand or "").lower()
    if "brother" in brand_low or "dcp" in brand_low or "hl" in brand_low or "mfc" in brand_low:
        return parse_brother_counters(html)

    if "samsung" in brand_low or "syncthru" in html.lower():
        data = parse_samsung_jsonlike_counters(html)
        if any(data.values()):
            return data

        data = parse_samsung_counters(html)
        if any(data.values()):
            return data

        for candidate in _build_samsung_candidate_urls(counter_url):
            try:
                resp = requests.get(candidate, timeout=(3, 15), headers=headers)
                resp.raise_for_status()
                text = resp.text or ""
            except Exception:
                continue

            data = parse_samsung_jsonlike_counters(text)
            if any(data.values()):
                return data

            data = parse_samsung_counters(text)
            if any(data.values()):
                return data

        return {"print": 0, "copy": 0, "scan": 0}

    data = parse_brother_counters(html)
    if any(data.values()):
        return data
    return parse_samsung_counters(html)
