"""
Structured data layer for the z/OS product compatibility matrix.

* Loads the vendor CSVs (*_compat.csv).
* Filters / matches the user's existing product versions for a chosen z/OS version.
* Builds the Excel workbook with the 8 required columns and CLICKABLE hyperlinks.

"Is product X at version Y compatible with z/OS Z" is exact tabular data, so it is
answered with pandas, never by an LLM guess.
"""

from __future__ import annotations

import glob
import io
import os
import re
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

VENDORS = ["IBM", "Broadcom", "BMC", "Rocket", "Precisely", "ASG"]
ZOS_VERSIONS = ["3.2", "3.1", "2.5"]  # versions the data covers (newest first)

CSV_COLUMNS = [
    "vendor", "product", "vrm", "zos31", "zos32", "zos25", "ptf_required", "ptf_required_31", "ptf_required_25",
    "eos_date", "latest_version", "comments", "link", "data_status", "last_checked",
]

# per z/OS version: (status column, PTF column, the previous release's status column, its label)
VER = {
    "3.2": ("zos32", "ptf_required", "zos31", "3.1"),
    "3.1": ("zos31", "ptf_required_31", "zos25", "2.5"),
    "2.5": ("zos25", "ptf_required_25", "", ""),   # oldest version covered
}

STATUS_FILL = {
    "Yes": "C6EFCE", "No": "FFC7CE", "In Progress": "FFEB9C",
    "Unverified": "D9D9D9", "N/A": "EDEDED", "Unknown": "D9D9D9",
}

# Yellow = "this product needs PTFs / usermods applied" (same colour as the reference workbook)
PTF_FILL = "FFD966"
PTF_CSS = "background-color: #FFD966; color: #000000; font-weight: 600"
_PTF_NOISE = re.compile(r"\b(?:PE PTF|RMID)\s+[A-Z]{2}\d{5}")        # PTFs mentioned only as "do not use / already in"
_PTF_RANGE = re.compile(r"\b([A-Z]{2}\d{3})(\d{2})((?:/\d{2})+)\b")   # LU05157/58/59 -> LU05157, LU05158, LU05159
_PTF_ID = re.compile(r"\b[A-Z]{2}\d{5}\b")                              # LU16693, RO60438, SO01425 ...

# Official entry points per vendor (all were reachable when this was built).
SOURCES = [
    ("IBM", "Software Product Compatibility Reports (SPCR)", "https://www.ibm.com/software/reports/compatibility/clarity/index.html", "Per-product Detailed System Requirements incl. z/OS 3.2"),
    ("IBM", "z/OS 3.2 announcement", "https://www.ibm.com/new/announcements/ibm-z-os-3-2-unlocks-the-value-of-ibm-z17", "z/OS 3.2 overview (GA 30 Sep 2025)"),
    ("IBM", "CICS TS detailed system requirements", "https://www.ibm.com/support/pages/detailed-system-requirements-cics-ts-and-related-products", "DSR reports for CICS TS 5.5 - 6.3"),
    ("IBM", "CICS product announcements", "https://www.ibm.com/support/pages/announcements-cics-products", "CICS announcement letters"),
    ("IBM", "IBM software lifecycle (EOS dates)", "https://www.ibm.com/support/pages/lifecycle/", "Search a product to get its end-of-service date"),
    ("Broadcom", "z/OS compatibility grid (all products)", "https://ftpdocs.broadcom.com/phpdocs/0/MSPSaccount/COMPAT/zos_compat.HTML", "Y/N/In-progress per release for z/OS 3.2 and 3.1, with PTF links"),
    ("Broadcom", "How to find compatibility matrices", "https://knowledge.broadcom.com/external/article/18439/how-to-find-endevor-compatibilities-matr.html", "Knowledge article explaining the compatibility pages"),
    ("BMC", "z/OS compatibility resources", "https://www.bmc.com/support/resources/os390-and-zos-compatibilty.html", "Points to the compatibility database / PAC utility"),
    ("BMC", "BMC AMI latest release", "https://www.bmc.com/it-solutions/bmc-ami-latest-release-oct-2025.html", "What's new across the BMC AMI portfolio"),
    ("BMC", "Control-M compatibility 9.0.20", "https://docs.bmc.com/docs/controlm/9020/control-m-compatibility-for-version-9-0-20-900847337.html", "Example of a per-version BMC compatibility page"),
    ("Rocket", "Rocket Software", "https://www.rocketsoftware.com/", "Vendor home; support portal links"),
    ("Rocket", "Rocket Open Source Solutions for z/OS", "https://www.rocketsoftware.com/node/56336", "Open-source tooling on z/OS"),
    ("Rocket", "Rocket Support for Zowe", "https://www.rocketsoftware.com/node/56111", "Zowe support"),
    ("Precisely", "Ironstream now compatible with z/OS 3.2", "https://community.precisely.com/discussion/ironstream-for-ibm-z-now-compatible-with-zos-32", "Announcement, 27 Jul 2026"),
    ("Precisely", "Syncsort MFX support lifecycle", "https://www.precisely.com/support/mfx-support-lifecycle-schedule", "Versions, support tiers, minimum z/OS"),
    ("Precisely", "Syncsort MFX product support", "https://support.precisely.com/products/syncsort-mfx/", "MFX downloads and docs"),
    ("ASG", "ASG products now under Rocket (TMON datasheet)", "https://www.rocketsoftware.com/sites/default/files/resource_files/tmon-zos.pdf", "ASG/TMON now sold as Rocket TMON"),
]


def out_columns(version: str = "3.2") -> list[str]:
    """The 8 requested columns; the compatibility header follows the chosen z/OS version."""
    return [
        "Vendor", "Product", "VRM", f"R{version} Compatible (Yes/No)", "If Yes, PTF Required?",
        "EOS Date", "Latest Version", "Comments / Links",
    ]


def needs_ptf(text) -> bool:
    """True when a PTF cell says PTFs / usermods are required (starts with 'Yes')."""
    return str(text).strip().lower().startswith("yes")


def extract_ptf_ids(text) -> list[str]:
    """PTF / usermod IDs mentioned in a PTF cell, without duplicates.

    Understands the shorthand 'LU05157/58/59' and ignores IDs that are only mentioned as a
    problem fix ('corrects PE PTF LU18249') or as an RMID condition ('RMID UJ98216').
    """
    t = _PTF_NOISE.sub("", str(text))
    found = []
    for m in _PTF_RANGE.finditer(t):
        found.append(m.group(1) + m.group(2))
        found += [m.group(1) + x for x in m.group(3).split("/")[1:]]
    found += _PTF_ID.findall(t)
    return list(dict.fromkeys(found))


def norm(s: str) -> str:
    s = re.sub(r"[^a-z0-9.]+", " ", str(s).lower()).strip()
    return re.sub(r"\bmgmt\b", "management", s)


def norm_vrm(s: str) -> str:
    return re.sub(r"^v\s*", "", norm(s))


@dataclass
class CompatibilityDB:
    df: pd.DataFrame

    # ------------------------------------------------------------------ load
    @classmethod
    def load(cls, data_dir: str = "data") -> "CompatibilityDB":
        """Find the vendor CSVs (*_compat.csv) wherever they were uploaded.

        Looks in data_dir first, then next to the app, then anywhere below it
        (so a flat GitHub upload without a data/ folder also works). Old
        '*_sample.csv' files from earlier versions are ignored.
        """
        base = os.path.dirname(os.path.abspath(data_dir))
        roots = [os.path.abspath(data_dir), base, os.getcwd()]
        files = []
        for root in roots:
            files = sorted(glob.glob(os.path.join(root, "*_compat.csv")))
            if files:
                break
        if not files:
            files = sorted(glob.glob(os.path.join(base, "**", "*_compat.csv"), recursive=True))
        files = [f for f in files if "sample" not in os.path.basename(f).lower()]
        if not files:
            raise FileNotFoundError(
                "No vendor data files found. Put broadcom_compat.csv, ibm_compat.csv, bmc_compat.csv, "
                "rocket_compat.csv, precisely_compat.csv and asg_compat.csv in a 'data' folder next to app.py "
                f"(searched: {', '.join(roots)})."
            )
        frames = [pd.read_csv(f, dtype=str, keep_default_na=False) for f in files]
        return cls(df=cls._clean(pd.concat(frames, ignore_index=True)))

    @staticmethod
    def _clean(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        for c in CSV_COLUMNS:
            if c not in df.columns:
                df[c] = ""
        df = df[CSV_COLUMNS].fillna("").astype(str)
        for c in df.columns:
            df[c] = df[c].str.strip()
        df["ptf_required_31"] = df["ptf_required_31"].replace("", "Verify")
        df["ptf_required_25"] = df["ptf_required_25"].replace("", "Verify")
        df["zos25"] = df["zos25"].replace("", "Unverified")   # no statement sourced = Unverified, never assumed
        df = df[(df["vendor"] != "") & (df["product"] != "")]
        df = df.drop_duplicates(subset=["vendor", "product", "vrm"], keep="last").reset_index(drop=True)
        df["label"] = df["vendor"] + " | " + df["product"] + " | " + df["vrm"]
        return df

    def merged_with(self, extra: pd.DataFrame | None) -> "CompatibilityDB":
        """New DB where rows from `extra` override matching vendor/product/VRM rows."""
        if extra is None or extra.empty:
            return self
        base = self.df.drop(columns=["label"], errors="ignore")
        return CompatibilityDB(df=self._clean(pd.concat([base, extra], ignore_index=True)))

    # --------------------------------------------------------------- filters
    def filter(self, vendors=None, status=None, ptf=None, breaking_only=False, search="",
               version: str = "3.2") -> pd.DataFrame:
        scol, pcol, prev_col, _ = VER[version]
        d = self.df
        if vendors:
            d = d[d["vendor"].isin(vendors)]
        if status:
            d = d[d[scol].isin(status)]
        if ptf == "PTF required":
            d = d[d[pcol].str.startswith("Yes")]
        elif ptf == "No PTF needed":
            d = d[d[pcol] == "No"]
        if breaking_only and prev_col:      # works on the previous release but not on this one
            d = d[(d[prev_col] == "Yes") & (d[scol] == "No")]
        if search:
            s = search.lower()
            d = d[d.apply(lambda r: s in " ".join(r.astype(str)).lower(), axis=1)]
        return d

    # ----------------------------------------------------- inventory matching
    def match_inventory(self, inv: pd.DataFrame):
        """inv needs columns Vendor, Product, VRM. Returns (matched_rows, unmatched_messages)."""
        cols = {c.lower().strip(): c for c in inv.columns}
        for need in ("vendor", "product", "vrm"):
            if need not in cols:
                raise ValueError("Inventory needs the columns: Vendor, Product, VRM")
        hits, misses = [], []
        for _, r in inv.iterrows():
            v, p, vr = str(r[cols["vendor"]]), str(r[cols["product"]]), str(r[cols["vrm"]])
            pool = self.df[self.df["vendor"].str.lower() == v.strip().lower()]
            pn = norm(p)
            exact = pool[pool["product"].map(norm) == pn]
            part = exact if not exact.empty else pool[pool["product"].map(norm).str.contains(re.escape(pn), na=False)]
            if part.empty:
                misses.append(f"{v} | {p} | {vr}: product not in the database - check the vendor link")
                continue
            m = part[part["vrm"].map(norm_vrm) == norm_vrm(vr)]
            if m.empty:
                known = ", ".join(sorted(part["vrm"].unique()))
                misses.append(f"{v} | {p} | {vr}: VRM not in the database (known: {known})")
            else:
                hits.append(m)
        matched = pd.concat(hits, ignore_index=True) if hits else self.df.iloc[0:0]
        return matched, misses

    # -------------------------------------------------------------- summaries
    @staticmethod
    def summarize(d: pd.DataFrame, version: str = "3.2") -> dict:
        scol, pcol, _, _ = VER[version]
        return {
            "total": len(d),
            "yes": int((d[scol] == "Yes").sum()),
            "no": int((d[scol] == "No").sum()),
            "in_progress": int((d[scol] == "In Progress").sum()),
            "unverified": int(d[scol].isin(["Unverified", "Unknown"]).sum()),
            "ptf": int(((d[scol] == "Yes") & d[pcol].str.startswith("Yes")).sum()),
        }


# ---------------------------------------------------------------------- PTF worklist
WORKLIST_COLUMNS = ["Vendor", "Product", "VRM", "PTFs / Usermods", "Details", "Link"]


def ptf_worklist(d: pd.DataFrame, version: str = "3.2") -> pd.DataFrame:
    """One row per product that needs PTFs / usermods, with the IDs pulled out of the text."""
    if d.empty:
        return pd.DataFrame(columns=WORKLIST_COLUMNS)
    pcol = VER[version][1]
    d = d[d[pcol].map(needs_ptf)]
    out = pd.DataFrame({
        "Vendor": d["vendor"], "Product": d["product"], "VRM": d["vrm"],
        "PTFs / Usermods": d[pcol].map(lambda t: ", ".join(extract_ptf_ids(t)) or "see Details / vendor link"),
        "Details": d[pcol], "Link": d["link"],
    })
    return out.reset_index(drop=True)


def unique_ptf_ids(d: pd.DataFrame, version: str = "3.2") -> list[str]:
    ids = []
    for t in d[VER[version][1]]:
        if needs_ptf(t):
            ids += extract_ptf_ids(t)
    return sorted(set(ids))


def smpe_template(ids: list[str]) -> str:
    """Starter SMP/E control statements. A template only: review zone names and options first."""
    body = "\n".join(f"    {i}" for i in ids)
    return (f"/* TEMPLATE - review zone names / options before running. {len(ids)} PTF(s) / usermod(s). */\n"
            f"SET BDY(GLOBAL) .\nRECEIVE SELECT(\n{body}\n) .\n\n"
            f"SET BDY(your-target-zone) .\nAPPLY SELECT(\n{body}\n) CHECK .")


# ---------------------------------------------------------------------- output
def comment_text(r: pd.Series, version: str = "3.2") -> str:
    others = [v for v in ZOS_VERSIONS if v != version]
    parts = [" ".join(f"z/OS {v}: {r[VER[v][0]]}." for v in others)]
    if r["comments"]:
        parts.append(r["comments"])
    parts.append(f"[{r['data_status']}; checked {r['last_checked']}]")
    return " ".join(parts)


def to_output(d: pd.DataFrame, with_link_column: bool = False, version: str = "3.2") -> pd.DataFrame:
    """Rows -> the 8 requested columns (optionally + a clickable 'Open link' column)."""
    cols = out_columns(version) + (["Open link"] if with_link_column else [])
    if d.empty:
        return pd.DataFrame(columns=cols)
    scol, pcol, _, _ = VER[version]
    out = pd.DataFrame({
        "Vendor": d["vendor"], "Product": d["product"], "VRM": d["vrm"],
        cols[3]: d[scol], "If Yes, PTF Required?": d[pcol],
        "EOS Date": d["eos_date"], "Latest Version": d["latest_version"],
        "Comments / Links": d.apply(
            lambda r: comment_text(r, version) + ("" if with_link_column else f" {r['link']}"), axis=1),
    })
    if with_link_column:
        out["Open link"] = d["link"]
    return out.reset_index(drop=True)


def _write_matrix_sheet(ws, d: pd.DataFrame, version: str):
    scol, pcol, _, _ = VER[version]
    ws.append(out_columns(version))
    head_fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for _, r in d.iterrows():
        text = comment_text(r, version)
        ws.append([r["vendor"], r["product"], r["vrm"], r[scol], r[pcol],
                   r["eos_date"], r["latest_version"], text + ("\n" + r["link"] if r["link"] else "")])
        row = ws.max_row
        cell = ws.cell(row=row, column=8)
        if r["link"].lower().startswith(("http://", "https://")):
            cell.hyperlink = r["link"]          # real, clickable Excel hyperlink
            cell.font = Font(color="0563C1", underline="single")
        fill = STATUS_FILL.get(r[scol])
        if fill:
            ws.cell(row=row, column=4).fill = PatternFill("solid", fgColor=fill)
        if needs_ptf(r[pcol]):                     # YELLOW = PTFs / usermods required
            ws.cell(row=row, column=5).fill = PatternFill("solid", fgColor=PTF_FILL)
        for col in range(1, 9):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([12, 38, 12, 18, 48, 26, 16, 90], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    _print_setup(ws)


def _print_setup(ws):
    """Landscape, one page wide, header row repeated: prints and PDF-exports cleanly."""
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"


def build_excel(sheets: dict[str, pd.DataFrame], notes: list[str] | None = None, version: str = "3.2") -> bytes:
    """sheets: {sheet name: DataFrame in CSV_COLUMNS layout}. Adds clickable Sources and Notes sheets."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, d in sheets.items():
        safe = re.sub(r"[\\/?*\[\]:]", "-", name)[:31]   # Excel forbids \ / ? * [ ] : in sheet names
        _write_matrix_sheet(wb.create_sheet(safe), d, version)

    first = next(iter(sheets.values()), None)
    if first is not None:
        wl = ptf_worklist(first, version)
        ws = wb.create_sheet("PTF Worklist")
        ws.append(WORKLIST_COLUMNS)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        for r in wl.itertuples(index=False):
            ws.append(list(r))
            row = ws.max_row
            ws.cell(row=row, column=4).fill = PatternFill("solid", fgColor=PTF_FILL)
            if str(r.Link).lower().startswith(("http://", "https://")):
                ws.cell(row=row, column=6).hyperlink = r.Link
                ws.cell(row=row, column=6).font = Font(color="0563C1", underline="single")
            for col in range(1, 7):
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        for i, w in enumerate([12, 38, 12, 40, 90, 60], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        _print_setup(ws)

    ws = wb.create_sheet("Sources")
    ws.append(["Vendor", "Source", "Link (click to open)", "What you get"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for v, name, url, what in SOURCES:
        ws.append([v, name, url, what])
        cell = ws.cell(row=ws.max_row, column=3)
        cell.hyperlink = url
        cell.font = Font(color="0563C1", underline="single")
    for i, w in enumerate([12, 48, 95, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("Notes")
    for line in [f"z/OS version in this workbook: {version}."] + (notes or []) + [
        f"R{version} Compatible: Yes / No / In Progress (vendor says support is coming) / Unverified (no vendor statement found yet).",
        "Broadcom rows come from Broadcom's public z/OS compatibility grid. 'PTF required = Yes' means the grid shows a PTF list for that release and z/OS version.",
        "YELLOW cell in 'If Yes, PTF Required?' = PTFs / usermods must be applied; the IDs are in the cell and on the 'PTF Worklist' sheet.",
        "'Latest Version' for Broadcom is the highest release shown on that grid, not necessarily the newest vendor GA.",
        "EOS dates are only filled when a vendor publishes them; otherwise check the vendor lifecycle page.",
        "Always confirm against the live vendor page before production changes.",
    ]:
        ws.append([line])
    ws.column_dimensions["A"].width = 150
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
