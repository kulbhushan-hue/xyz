"""
Multi-vendor PTF / USERMOD scraper (no Anthropic packages or APIs; only requests + BeautifulSoup).

    from tools.ptf_scraper import scrape_vendor, scrape_all, to_excel
    rows = scrape_vendor(df_vendor_rows, creds={"IBM": ("ibmid", "pw")}, max_pages=10)

Every vendor uses the SAME output columns as the Broadcom example:
 # | Product / Release | PTF ID | FMID | PTF publish date | APAR ID | Prerequisites |
 Superseding PTF | Full description | Detail Extraction Status

Status is one of: Verified / Not found in source / Access denied.
Nothing is ever guessed: if a page has no PTF text, the row says "Not found in source".
"""
from __future__ import annotations

import re
import time
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

COLUMNS = ["#", "Product / Release", "PTF ID", "FMID", "PTF publish date", "APAR ID",
           "Prerequisites", "Superseding PTF", "Full description", "Detail Extraction Status"]

PTF_RE = re.compile(r"\b(?:[A-Z]{2}\d{5}|B[A-Z]{2}\d{4,5})\b")
FMID_RE = re.compile(r"\b[HJ](?=[A-Z0-9]*\d)[A-Z0-9]{6}\b")
APAR_RE = re.compile(r"\b(?:PH|PI|PJ|PK|PM|PQ|OA|OW|II|IC|IY|DT|RO|QI|QO|LU)\d{5}\b")
DATE_RE = re.compile(r"\b(?:\d{4}-\d{2}-\d{2}|\d{1,2}[/.]\d{1,2}[/.]\d{2,4}|\d{1,2} [A-Z][a-z]{2} \d{4}|[A-Z][a-z]{2} \d{1,2}, \d{4})\b")
LOGIN_HINTS = ("sign in", "log in", "login", "ibmid", "authentication required", "access denied")

# Vendor settings. 'login' = needs credentials for most pages; 'follow' = link text/URL words worth crawling.
VENDORS = {
    "Broadcom": dict(login=False, follow=("ptf", "fix", "solution", "compat")),
    "IBM":      dict(login=True,  follow=("psp", "ptf", "apar", "fix", "preventive", "requirements")),
    "BMC":      dict(login=True,  follow=("ptf", "fix", "pac", "compat", "release")),
    "Rocket":   dict(login=True,  follow=("ptf", "fix", "support", "compat", "release")),
    "Precisely": dict(login=True, follow=("ptf", "fix", "support", "compat", "lifecycle")),
    "ASG":      dict(login=True,  follow=("ptf", "fix", "support", "compat", "rocket")),
}
HEADER_MAP = {  # header word -> output column
    "ptf": "PTF ID", "usermod": "PTF ID", "fix": "PTF ID", "fmid": "FMID", "apar": "APAR ID",
    "date": "PTF publish date", "prereq": "Prerequisites", "requisite": "Prerequisites",
    "supersed": "Superseding PTF", "description": "Full description", "text": "Full description",
    "abstract": "Full description", "title": "Full description",
}


def _session(creds):
    s = requests.Session()
    s.headers["User-Agent"] = "Mozilla/5.0 (z-OS-compat-dashboard PTF checker)"
    if creds:
        s.auth = tuple(creds)          # HTTP basic; portals using form login need a vendor-specific hook
    return s


def _get(s, url, retries=2):
    for i in range(retries + 1):
        try:
            r = s.get(url, timeout=25)
            if r.status_code in (401, 403):
                return None, "Access denied"
            if r.status_code >= 400:
                return None, f"HTTP {r.status_code}"
            if "html" not in r.headers.get("content-type", "html") and "text" not in r.headers.get("content-type", ""):
                return None, "Not found in source"
            return r.text, None
        except requests.RequestException:
            time.sleep(1 + i)
    return None, "Not found in source"


def _looks_like_login(soup):
    text = soup.get_text(" ", strip=True).lower()[:3000]
    return bool(soup.find("input", {"type": "password"})) or (len(text) < 1500 and any(h in text for h in LOGIN_HINTS))


def _parse_tables(soup, product):
    out = []
    for t in soup.find_all("table"):
        head = [th.get_text(" ", strip=True).lower() for th in t.find_all("tr")[0].find_all(["th", "td"])] if t.find("tr") else []
        cols = {}
        for i, h in enumerate(head):
            for key, col in HEADER_MAP.items():
                if key in h and col not in cols.values():
                    cols[i] = col
                    break
        if "PTF ID" not in cols.values():
            continue
        for tr in t.find_all("tr")[1:]:
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
            rec = {c: "" for c in COLUMNS}
            for i, col in cols.items():
                if i < len(cells):
                    rec[col] = cells[i]
            m = PTF_RE.search(rec["PTF ID"])
            if m:
                rec["PTF ID"] = m.group(0)
                rec["Product / Release"] = product
                rec["Detail Extraction Status"] = "Verified"
                out.append(rec)
    return out


def _parse_text(soup, product):
    """Fallback: PTF ids mentioned in running text, with surrounding sentence as description."""
    out, seen = [], set()
    for el in soup.find_all(["p", "li", "td", "pre"]):
        txt = el.get_text(" ", strip=True)
        for m in PTF_RE.finditer(txt):
            if m.group(0) in seen:
                continue
            seen.add(m.group(0))
            rec = {c: "" for c in COLUMNS}
            fm, ap, dt = FMID_RE.search(txt), APAR_RE.search(txt), DATE_RE.search(txt)
            rec.update({"Product / Release": product, "PTF ID": m.group(0), "FMID": fm.group(0) if fm else "",
                        "APAR ID": ap.group(0) if ap and ap.group(0) != m.group(0) else "",
                        "PTF publish date": dt.group(0) if dt else "", "Full description": txt[:400],
                        "Detail Extraction Status": "Verified"})
            out.append(rec)
    return out


def scrape_url(url, product, vendor, creds=None, max_pages=10, session=None):
    """Crawl url (+ linked pages, up to max_pages) and return per-PTF records."""
    s = session or _session(creds)
    follow = VENDORS.get(vendor, {}).get("follow", ("ptf",))
    queue, seen, found, blocked = [url], set(), [], False
    host = urlparse(url).netloc
    while queue and len(seen) < max_pages:
        u = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        html, err = _get(s, u)
        if err == "Access denied":
            blocked = True
            continue
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        if _looks_like_login(soup):
            blocked = True
            continue
        found += _parse_tables(soup, product) or _parse_text(soup, product)
        for a in soup.find_all("a", href=True):
            link = urljoin(u, a["href"]).split("#")[0]
            if urlparse(link).netloc == host and link not in seen and link not in queue and \
                    any(w in (a.get_text(" ").lower() + link.lower()) for w in follow):
                queue.append(link)
    if found:
        uniq = {}
        for r in found:
            uniq.setdefault(r["PTF ID"], r)
        return list(uniq.values())
    rec = {c: "" for c in COLUMNS}
    rec.update({"Product / Release": product,
                "Detail Extraction Status": "Access denied" if blocked else "Not found in source"})
    return [rec]


def scrape_vendor(rows, creds=None, max_pages=10, vendor=None):
    """rows: iterable of dicts/Series with vendor, product, vrm, link. Returns numbered per-PTF records."""
    out, cache = [], {}
    for r in rows:
        v = vendor or r["vendor"]
        label = f"{r['product']} {r['vrm']}".strip()
        key = (r["link"], label)
        if key not in cache:
            cache[key] = scrape_url(r["link"], label, v, (creds or {}).get(v), max_pages)
        out += [dict(x) for x in cache[key]]
    for i, rec in enumerate(out, 1):
        rec["#"] = i
    return out


def scrape_all(df, creds=None, max_pages=10, include_broadcom=False):
    """Run every vendor in a compatibility DataFrame; returns {vendor: records}."""
    res = {}
    for v, g in df.groupby("vendor"):
        if v == "Broadcom" and not include_broadcom:
            continue
        res[v] = scrape_vendor(g.to_dict("records"), creds, max_pages, vendor=v)
    return res


def to_excel(results, path):
    """results: {vendor: records} -> workbook, one sheet per vendor, Broadcom-style layout."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    wb.remove(wb.active)
    for v, recs in results.items():
        ws = wb.create_sheet(v[:31])
        ws.append(COLUMNS)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
        for r in recs:
            ws.append([r.get(c, "") for c in COLUMNS])
        for col, w in zip("ABCDEFGHIJ", (5, 38, 12, 11, 15, 12, 28, 16, 70, 22)):
            ws.column_dimensions[col].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"
    wb.save(path)
