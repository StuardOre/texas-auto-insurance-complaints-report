#!/usr/bin/env python3
"""
Texas auto insurance complaints - automated Excel report
=========================================================

Pulls public complaint data from the Texas Department of Insurance (TDI)
open-data API (data.texas.gov), cleans it with pandas and writes a formatted
Excel report: KPIs, monthly trend, top complaint reasons, most-named insurers,
a size-adjusted carrier complaint index, the full file-level data and a data
quality log. It also saves the aggregates as JSON for slides or dashboards.

Usage
-----
    python tdi_auto_complaints_report.py --year 2025
    python tdi_auto_complaints_report.py --year 2025 --out my_report.xlsx

Sources (public, refreshed daily by TDI)
----------------------------------------
    Insurance complaints: All data ............ data.texas.gov/d/ubdr-4uff
    Complaint indexes and policy counts ....... data.texas.gov/d/pa9u-9s9w

Author: Stuard O. (Cesar Ore) - Excel & Python data automation
"""
import argparse
import datetime as dt
import json
import time
from pathlib import Path

import numpy as np
import pandas as pd
import requests

BASE = "https://data.texas.gov/resource"
COMPLAINTS_ID = "ubdr-4uff"
INDEX_ID = "pa9u-9s9w"
NAVY, BLUE, LIGHT, RED, GREEN = "#1F3864", "#2E75B6", "#DDEBF7", "#C00000", "#548235"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
ACRONYMS = {"USAA", "GEICO", "LLC", "INC", "INC.", "CO", "CO.", "II", "III", "AAA", "TX", "PC", "P&C"}


# --------------------------------------------------------------------------- #
# 1. Extract
# --------------------------------------------------------------------------- #
def fetch(dataset: str, where: str, page: int = 50_000) -> pd.DataFrame:
    """Download every row that matches a SoQL $where filter (paged, with retries)."""
    rows, offset = [], 0
    while True:
        params = {"$where": where, "$limit": page, "$offset": offset, "$order": ":id"}
        for attempt in range(3):
            try:
                r = requests.get(f"{BASE}/{dataset}.json", params=params, timeout=180)
                r.raise_for_status()
                break
            except requests.RequestException:
                if attempt == 2:
                    raise
                time.sleep(3 * (attempt + 1))
        chunk = r.json()
        rows.extend(chunk)
        print(f"  {dataset}: {len(rows):,} rows")
        if len(chunk) < page:
            return pd.DataFrame(rows)
        offset += page


# --------------------------------------------------------------------------- #
# 2. Transform
# --------------------------------------------------------------------------- #
def split_multi(value) -> list:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return []
    return [x.strip() for x in str(value).split(";") if x.strip()]


def join_unique(series: pd.Series) -> str:
    seen = []
    for v in series:
        for x in split_multi(v):
            if x not in seen:
                seen.append(x)
    return "; ".join(seen)


def nice_name(name: str) -> str:
    """'PROGRESSIVE COUNTY MUTUAL INSURANCE COMPANY' -> 'Progressive County Mutual Insurance Company'."""
    if not isinstance(name, str) or not name.isupper():
        return name
    return " ".join(w if w in ACRONYMS else w.capitalize() for w in name.split())


def to_files(rows: pd.DataFrame) -> pd.DataFrame:
    """The source has one row per complaint x respondent: collapse to one row per complaint file."""
    df = rows.copy()
    for c in ("received_date", "closed_date"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
    df["complaint_number"] = pd.to_numeric(df["complaint_number"], errors="coerce")
    files = (
        df.groupby("complaint_number")
        .agg(
            received=("received_date", "min"),
            closed=("closed_date", "max"),
            filed_by=("complainant_role", "first"),
            respondents=("respondent_name", lambda s: "; ".join(dict.fromkeys(nice_name(x) for x in s.dropna()))),
            coverage_level=("coverage_level", "first"),
            reasons=("reason", join_unique),
            confirmed=("complaint_confirmed_code", lambda s: "Yes" if (s == "Yes").any() else "No"),
            disposition=("disposition", join_unique),
            keyword=("keyword", "first"),
            source_rows=("complaint_number", "size"),
        )
        .reset_index()
        .sort_values("complaint_number")
    )
    files["days_to_close"] = (files["closed"] - files["received"]).dt.days
    files["status"] = np.where(files["closed"].notna(), "Closed", "Open")
    return files


def explode_count(files: pd.DataFrame, col: str, top: int) -> pd.DataFrame:
    ex = files[["complaint_number", col, "confirmed"]].copy()
    ex[col] = ex[col].map(split_multi)
    ex = ex.explode(col).dropna(subset=[col])
    out = (
        ex.groupby(col)
        .agg(files=("complaint_number", "nunique"),
             confirmed=("confirmed", lambda s: int((s == "Yes").sum())))
        .sort_values("files", ascending=False)
        .head(top)
        .reset_index()
    )
    out["share"] = out["files"] / len(files)
    out["confirmed_rate"] = out["confirmed"] / out["files"]
    return out


def build(year: int, asof: dt.date, lag_days: int = 120) -> dict:
    start = f"{year - 1}-01-01T00:00:00"
    print("Downloading complaint records ...")
    rows = fetch(COMPLAINTS_ID, f"coverage_type='Automobile' AND received_date >= '{start}'")
    print("Downloading carrier complaint index ...")
    idx = fetch(INDEX_ID, f"type_cd='Automobile' AND year='{year}'")

    files_all = to_files(rows)
    files = files_all[files_all["received"].dt.year == year].copy()
    closed = files[files["status"] == "Closed"]

    # Monthly trend: months x years. TDI publishes complaints once they are CLOSED, so the most
    # recent months are still filling in; only months that ended `lag_days` ago are shown.
    last_complete = (pd.Timestamp(asof) - pd.Timedelta(days=lag_days)).to_period("M") - 1
    m = files_all.assign(period=files_all["received"].dt.to_period("M"))
    m = m[m["period"] <= last_complete]
    monthly = (
        m.groupby([m["period"].dt.year.rename("year"), m["period"].dt.month.rename("month")])
        .size().unstack("year").reindex(range(1, 13))
    )

    # Carrier complaint index (size-adjusted): largest auto insurers by policies in force
    for c in ("col1", "col2", "col3"):
        idx[c] = pd.to_numeric(idx[c], errors="coerce")
    idx = idx.rename(columns={"col1": "confirmed_complaints", "col2": "policies", "col3": "complaint_index"})
    market = {"policies": int(idx["policies"].sum()), "confirmed": int(idx["confirmed_complaints"].sum())}
    market["per_100k"] = market["confirmed"] / market["policies"] * 100_000
    carriers = idx.sort_values("policies", ascending=False).head(20).copy()
    carriers["company"] = carriers["company_name"].map(nice_name)
    carriers["per_100k"] = carriers["confirmed_complaints"] / carriers["policies"] * 100_000

    buckets = pd.cut(closed["days_to_close"], [-1, 30, 60, 90, 180, 10_000],
                     labels=["0-30 days", "31-60 days", "61-90 days", "91-180 days", "Over 180 days"])
    kpi = {
        "files": int(len(files)),
        "confirmed": int((files["confirmed"] == "Yes").sum()),
        "open": int((files["status"] == "Open").sum()),
        "within_90": float((closed["days_to_close"] <= 90).mean()),
        "median_days": float(closed["days_to_close"].median()),
        "attorney_share": float((files["filed_by"] == "Attorney").mean()),
        "money_share": float(files["disposition"].str.contains("Additional Monies Received", na=False).mean()),
        "prev_year_files": int((files_all["received"].dt.year == year - 1).sum()),
    }
    kpi["confirmed_rate"] = kpi["confirmed"] / kpi["files"]
    kpi["yoy_change"] = kpi["files"] / kpi["prev_year_files"] - 1

    quality = [
        ("Source rows downloaded (since Jan %d)" % (year - 1), f"{len(rows):,}", "One row per complaint x respondent"),
        ("Unique complaint files (all years pulled)", f"{len(files_all):,}", "Collapsed on complaint_number"),
        ("Files received in %d" % year, f"{len(files):,}", "Scope of this report"),
        ("Files naming more than one respondent", f"{int((files['source_rows'] > 1).sum()):,}", "Respondents joined with ';'"),
        ("Records without a closed date", f"{kpi['open']:,}", "TDI publishes closed (processed) complaints only"),
        ("Files without a reason", f"{int((files['reasons'] == '').sum()):,}", ""),
        ("Files without a keyword", f"{int(files['keyword'].isna().sum()):,}", "Keyword is optional in the source"),
        ("Closed before received (date errors)", f"{int((files['days_to_close'] < 0).sum()):,}", "Should be 0"),
        ("Carriers in %d complaint index (auto)" % year, f"{len(idx):,}", "TDI policies in force as of Dec 31"),
    ]
    return {
        "year": year, "asof": asof.isoformat(), "kpi": kpi, "files": files, "monthly": monthly,
        "reasons": explode_count(files, "reasons", 12),
        "respondents": explode_count(files, "respondents", 12),
        "filed_by": files["filed_by"].fillna("Not stated").value_counts().rename_axis("filed_by").reset_index(name="files"),
        "buckets": buckets.value_counts().reindex(buckets.cat.categories).rename_axis("bucket").reset_index(name="files"),
        "carriers": carriers, "market": market, "quality": quality, "last_complete": str(last_complete),
    }


# --------------------------------------------------------------------------- #
# 3. Load: formatted Excel workbook
# --------------------------------------------------------------------------- #
def write_excel(d: dict, out: Path) -> None:
    year, k = d["year"], d["kpi"]
    wb = pd.ExcelWriter(out, engine="xlsxwriter")
    book = wb.book
    F = {
        "title": book.add_format({"bold": True, "font_size": 16, "font_color": NAVY}),
        "sub": book.add_format({"italic": True, "font_size": 9, "font_color": "#595959"}),
        "hdr": book.add_format({"bold": True, "font_color": "white", "bg_color": NAVY, "border": 1, "text_wrap": True, "valign": "vcenter"}),
        "txt": book.add_format({"border": 1, "valign": "top"}),
        "wrap": book.add_format({"border": 1, "text_wrap": True, "valign": "top"}),
        "int": book.add_format({"border": 1, "num_format": "#,##0"}),
        "dec": book.add_format({"border": 1, "num_format": "0.00"}),
        "one": book.add_format({"border": 1, "num_format": "0.0"}),
        "pct": book.add_format({"border": 1, "num_format": "0.0%"}),
        "date": book.add_format({"border": 1, "num_format": "yyyy-mm-dd"}),
        "kpil": book.add_format({"bold": True, "border": 1, "bg_color": LIGHT}),
        "kint": book.add_format({"bold": True, "border": 1, "font_size": 12, "num_format": "#,##0", "font_color": NAVY}),
        "kpct": book.add_format({"bold": True, "border": 1, "font_size": 12, "num_format": "0.0%", "font_color": NAVY}),
        "kone": book.add_format({"bold": True, "border": 1, "font_size": 12, "num_format": "0", "font_color": NAVY}),
        "note": book.add_format({"font_size": 9, "font_color": "#595959", "text_wrap": True, "valign": "top"}),
        "red": book.add_format({"font_color": RED, "bold": True}),
        "green": book.add_format({"font_color": GREEN, "bold": True}),
        "bold": book.add_format({"bold": True}),
    }
    src = (f"Source: Texas Department of Insurance open data (data.texas.gov, datasets ubdr-4uff and pa9u-9s9w), "
           f"retrieved {d['asof']}. Generated automatically by tdi_auto_complaints_report.py.")

    def header(ws, title, row=3, cols=()):
        ws.write(0, 0, title, F["title"])
        ws.write(1, 0, src, F["sub"])
        for j, c in enumerate(cols):
            ws.write(row, j, c, F["hdr"])

    ws_sum = book.add_worksheet("Summary")  # filled at the end so it stays the first tab

    # ---- README
    ws = book.add_worksheet("README")
    ws.set_column(0, 0, 120)
    lines = [
        (f"Texas Auto Insurance Complaints - {year} Report", F["title"]),
        ("Automated Excel report built with Python (requests + pandas + XlsxWriter) on real public regulator data.", None),
        ("", None),
        ("What it answers", F["bold"]),
        ("How many auto-insurance complaints Texans filed, how fast they were resolved, what they were about, which insurers were named most,", None),
        ("and how insurers compare once complaints are adjusted for size (TDI complaint index).", None),
        ("", None),
        ("Sheets", F["bold"]),
        ("Summary - KPIs (formulas over the data sheet) and charts", None),
        ("Monthly - complaint files received per month, three-year comparison", None),
        ("Reasons / Respondents / Breakdowns - what complaints were about, who was named, who filed them and how long they took", None),
        ("Carriers - 20 largest auto insurers by policies in force, confirmed complaints and TDI complaint index", None),
        (f"Complaints_{year} - one row per complaint file (filters on); Data_Quality - checks run on every refresh", None),
        ("", None),
        ("Method", F["bold"]),
        ("1. Download every automobile complaint received since Jan 1 of the previous year from the TDI API (paged SoQL queries).", None),
        ("2. Collapse the source (one row per complaint x respondent) to one row per complaint file; reasons and dispositions are de-duplicated.", None),
        ("3. A file counts as confirmed if TDI confirmed an error by at least one respondent. Days to close = closed date - received date.", None),
        ("4. Complaint index (from TDI): a company's share of confirmed complaints / its share of policies in force. 1.00 = average; above 1 = more complaints than its size suggests.", None),
        ("Note: TDI publishes complaints once they are closed, so the latest months are incomplete; the monthly trend stops at the last complete month.", None),
        ("", None),
        ("Refresh", F["bold"]),
        (f"python tdi_auto_complaints_report.py --year {year}   (runs in seconds; no manual steps)", None),
        ("", None),
        (src, F["sub"]),
    ]
    for i, (t, f) in enumerate(lines):
        ws.write(i, 0, t, f) if f else ws.write(i, 0, t)

    # ---- Complaints (file-level data)
    files = d["files"]
    sheet = f"Complaints_{year}"
    ws = book.add_worksheet(sheet)
    cols = ["Complaint #", "Received", "Closed", "Days to close", "Status", "Filed by", "Respondent(s)",
            "Coverage level", "Reason(s)", "Confirmed", "Disposition", "Keyword"]
    header(ws, f"Auto insurance complaint files received in {year} - one row per file", cols=cols)
    first, last = 4, 4 + len(files) - 1
    for i, r in enumerate(files.itertuples(index=False)):
        row = first + i
        ws.write_number(row, 0, int(r.complaint_number), F["txt"])
        ws.write_datetime(row, 1, r.received.to_pydatetime(), F["date"])
        if pd.notna(r.closed):
            ws.write_datetime(row, 2, r.closed.to_pydatetime(), F["date"])
            ws.write_number(row, 3, int(r.days_to_close), F["int"])
        else:
            ws.write_blank(row, 2, None, F["date"])
            ws.write_blank(row, 3, None, F["int"])
        for j, v in enumerate([r.status, r.filed_by, r.respondents, r.coverage_level, r.reasons,
                               r.confirmed, r.disposition, r.keyword], start=4):
            ws.write_string(row, j, "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v), F["txt"])
    ws.autofilter(3, 0, last, len(cols) - 1)
    ws.freeze_panes(4, 1)
    for j, w in enumerate([12, 11, 11, 9, 8, 13, 38, 20, 48, 10, 48, 22]):
        ws.set_column(j, j, w)

    def rng(col):
        return f"{sheet}!${col}${first + 1}:${col}${last + 1}"

    # ---- Monthly
    ws = book.add_worksheet("Monthly")
    mon = d["monthly"]
    years = list(mon.columns)
    lc_month = MONTHS[int(d["last_complete"][-2:]) - 1]
    header(ws, "Complaint files received per month (TDI publishes closed complaints, so recent months are left out until they fill in)",
           cols=["Month"] + [str(y) + (f" (Jan-{lc_month})" if y > year else "") for y in years])
    for i, mth in enumerate(range(1, 13)):
        ws.write(4 + i, 0, MONTHS[i], F["txt"])
        for j, y in enumerate(years, start=1):
            v = mon.loc[mth, y]
            if pd.notna(v):
                ws.write_number(4 + i, j, int(v), F["int"])
            else:
                ws.write_blank(4 + i, j, None, F["int"])
    ws.write(16, 0, "Total", F["kpil"])
    for j in range(1, len(years) + 1):
        c = chr(65 + j)
        ws.write_formula(16, j, f"=SUM({c}5:{c}16)", F["kint"], int(mon.iloc[:, j - 1].sum()))
    ws.set_column(0, 0, 10)
    ws.set_column(1, len(years), 14)

    def month_chart():
        ch = book.add_chart({"type": "line"})
        palette = ["#A6A6A6", BLUE, NAVY]
        for j, y in enumerate(years, start=1):
            ch.add_series({"name": ["Monthly", 3, j], "categories": ["Monthly", 4, 0, 15, 0],
                           "values": ["Monthly", 4, j, 15, j],
                           "line": {"color": palette[(j - 1) % 3], "width": 2.25}, "marker": {"type": "none"}})
        ch.set_title({"name": "Auto complaint files received per month", "name_font": {"size": 12}})
        ch.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
        ch.set_legend({"position": "bottom"})
        ch.set_size({"width": 620, "height": 300})
        return ch

    ws.insert_chart("F4", month_chart())

    # ---- Reasons / Respondents
    def ranked_chart(name, n, title):
        ch = book.add_chart({"type": "bar"})
        ch.add_series({"categories": [name, 4, 0, 3 + n, 0], "values": [name, 4, 1, 3 + n, 1],
                       "fill": {"color": BLUE}, "data_labels": {"value": True, "num_format": "#,##0"}})
        ch.set_title({"name": title, "name_font": {"size": 12}})
        ch.set_y_axis({"reverse": True})
        ch.set_x_axis({"num_format": "#,##0", "major_gridlines": {"visible": False}})
        ch.set_legend({"none": True})
        ch.set_size({"width": 620, "height": 380})
        return ch

    def ranked(name, df, label, title):
        ws = book.add_worksheet(name)
        header(ws, title, cols=[label, "Files", "% of files", "Confirmed files", "Confirmed rate"])
        for i, r in enumerate(df.itertuples(index=False)):
            ws.write(4 + i, 0, getattr(r, df.columns[0]), F["txt"])
            ws.write_number(4 + i, 1, r.files, F["int"])
            ws.write_number(4 + i, 2, r.share, F["pct"])
            ws.write_number(4 + i, 3, r.confirmed, F["int"])
            ws.write_number(4 + i, 4, r.confirmed_rate, F["pct"])
        ws.write(4 + len(df) + 1, 0, "A file can list several reasons/respondents, so shares add up to more than 100%.", F["note"])
        ws.set_column(0, 0, 46)
        ws.set_column(1, 4, 13)
        ws.insert_chart("G4", ranked_chart(name, len(df), title))

    reasons_title = f"Top complaint reasons, {year} (files)"
    ranked("Reasons", d["reasons"], "Reason", reasons_title)
    ranked("Respondents", d["respondents"], "Respondent (as named in the file)", f"Most-named respondents, {year} (files)")

    # ---- Breakdowns
    ws = book.add_worksheet("Breakdowns")
    header(ws, f"Who filed the complaint and how long TDI took to close it ({year} files)", cols=["Filed by", "Files", "% of files"])
    fb = d["filed_by"]
    for i, r in enumerate(fb.itertuples(index=False)):
        ws.write(4 + i, 0, r.filed_by, F["txt"])
        ws.write_number(4 + i, 1, r.files, F["int"])
        ws.write_formula(4 + i, 2, f"=B{5 + i}/SUM($B$5:$B${4 + len(fb)})", F["pct"], r.files / fb["files"].sum())
    b0 = 6 + len(fb)
    for j, c in enumerate(["Days to close (closed files)", "Files", "% of closed"]):
        ws.write(b0, j, c, F["hdr"])
    bk = d["buckets"]
    for i, r in enumerate(bk.itertuples(index=False)):
        ws.write(b0 + 1 + i, 0, str(r.bucket), F["txt"])
        ws.write_number(b0 + 1 + i, 1, int(r.files), F["int"])
        ws.write_formula(b0 + 1 + i, 2, f"=B{b0 + 2 + i}/SUM($B${b0 + 2}:$B${b0 + 1 + len(bk)})", F["pct"], r.files / bk["files"].sum())
    ws.set_column(0, 0, 30)
    ws.set_column(1, 2, 12)

    # ---- Carriers (size-adjusted)
    ws = book.add_worksheet("Carriers")
    cols = ["Company", "NAIC #", "Policies in force (Dec 31)", "Confirmed complaints", "TDI complaint index",
            "Confirmed complaints per 100k policies"]
    header(ws, f"20 largest Texas auto insurers by policies in force - complaint index {year}", cols=cols)
    car = d["carriers"]
    for i, r in enumerate(car.itertuples(index=False)):
        row = 4 + i
        ws.write(row, 0, r.company, F["txt"])
        ws.write(row, 1, str(r.naic_id), F["txt"])
        ws.write_number(row, 2, int(r.policies), F["int"])
        ws.write_number(row, 3, int(r.confirmed_complaints), F["int"])
        ws.write_number(row, 4, float(r.complaint_index), F["dec"])
        ws.write_formula(row, 5, f"=IFERROR(D{row + 1}/C{row + 1}*100000,0)", F["one"], float(r.per_100k))
    end = 4 + len(car)
    mk = d["market"]
    ws.write(end, 0, "All auto insurers in the TDI index", F["kpil"])
    ws.write(end, 1, "", F["kpil"])
    ws.write_number(end, 2, mk["policies"], F["kint"])
    ws.write_number(end, 3, mk["confirmed"], F["kint"])
    ws.write_number(end, 4, 1.0, F["dec"])
    ws.write_formula(end, 5, f"=D{end + 1}/C{end + 1}*100000", F["one"], mk["per_100k"])
    ws.conditional_format(4, 4, end - 1, 4, {"type": "cell", "criteria": ">", "value": 1.2, "format": F["red"]})
    ws.conditional_format(4, 4, end - 1, 4, {"type": "cell", "criteria": "<", "value": 0.8, "format": F["green"]})
    ws.write(end + 2, 0, "Complaint index = company share of confirmed complaints / company share of policies in force (TDI). "
                         "1.00 is the market average. Red: above 1.2, green: below 0.8.", F["note"])
    ws.set_column(0, 0, 50)
    ws.set_column(1, 1, 9)
    ws.set_column(2, 5, 16)

    def car_chart():
        ch = book.add_chart({"type": "bar"})
        ch.add_series({"name": "TDI complaint index", "categories": ["Carriers", 4, 0, end - 1, 0],
                       "values": ["Carriers", 4, 4, end - 1, 4], "fill": {"color": NAVY},
                       "data_labels": {"value": True, "num_format": "0.00"}})
        ch.set_title({"name": f"Complaint index, 20 largest auto insurers ({year}; 1.00 = average)", "name_font": {"size": 12}})
        ch.set_y_axis({"reverse": True})
        ch.set_legend({"none": True})
        ch.set_size({"width": 680, "height": 520})
        return ch

    ws.insert_chart("H4", car_chart())

    # ---- Data quality
    ws = book.add_worksheet("Data_Quality")
    header(ws, "Data quality checks (run on every refresh)", cols=["Check", "Result", "Note"])
    for i, (a, b, c) in enumerate(d["quality"]):
        ws.write(4 + i, 0, a, F["txt"])
        ws.write(4 + i, 1, b, F["txt"])
        ws.write(4 + i, 2, c, F["wrap"])
    ws.set_column(0, 0, 44)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 46)

    # ---- Summary (first tab)
    ws = ws_sum
    ws.activate()
    ws.write(0, 0, f"Texas Auto Insurance Complaints - {year}", F["title"])
    ws.write(1, 0, src, F["sub"])
    for j, c in enumerate(["KPI", "Value", "How it is calculated"]):
        ws.write(3, j, c, F["hdr"])
    kp = [
        ("Complaint files received", f"=COUNT({rng('A')})", k["files"], "kint", "Unique complaint numbers received in the year"),
        ("Change vs previous year", "=B5/Monthly!B17-1", k["yoy_change"], "kpct", "Files received vs the full previous year"),
        ("Confirmed by TDI", f"=COUNTIF({rng('J')},\"Yes\")", k["confirmed"], "kint", "TDI confirmed an error by the insurer or licensee"),
        ("Confirmed rate", "=B7/B5", k["confirmed_rate"], "kpct", "Confirmed files / files received"),
        ("Median days to close", f"=MEDIAN({rng('D')})", k["median_days"], "kone", "Closed files only"),
        ("Closed within 90 days", f"=COUNTIF({rng('D')},\"<=90\")/B5", k["within_90"], "kpct", "Share of files resolved by TDI in 90 days or less"),
        ("Filed by an attorney", f"=COUNTIF({rng('F')},\"Attorney\")/B5", k["attorney_share"], "kpct", "Share of files"),
        ("Consumer received additional money", f"=COUNTIF({rng('K')},\"*Additional Monies Received*\")/B5",
         k["money_share"], "kpct", "Disposition includes 'Additional Monies Received'"),
    ]
    for i, (label, formula, value, fmt, note) in enumerate(kp):
        ws.write(4 + i, 0, label, F["kpil"])
        ws.write_formula(4 + i, 1, formula, F[fmt], value)
        ws.write(4 + i, 2, note, F["txt"])
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 48)
    ws.insert_chart("A15", month_chart())
    ws.insert_chart("E4", ranked_chart("Reasons", len(d["reasons"]), reasons_title), {"x_offset": 20})
    ws.insert_chart("E25", car_chart(), {"x_offset": 20})
    ws.hide_gridlines(2)

    # ---- Print setup: every sheet prints cleanly (Summary on one landscape page)
    for sh in book.worksheets():
        sh.set_landscape()
        sh.set_paper(1)  # US Letter
        sh.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
        sh.fit_to_pages(1, 0)
    ws_sum.fit_to_pages(1, 1)
    ws_sum.print_area("A1:Q53")
    wb.close()


def save_json(d: dict, out: Path) -> None:
    payload = {
        "year": d["year"], "asof": d["asof"], "kpi": d["kpi"], "market": d["market"], "last_complete": d["last_complete"],
        "monthly": {str(c): [None if pd.isna(v) else int(v) for v in d["monthly"][c]] for c in d["monthly"].columns},
        "reasons": d["reasons"].to_dict("records"), "respondents": d["respondents"].to_dict("records"),
        "filed_by": d["filed_by"].to_dict("records"),
        "buckets": [{"bucket": str(r.bucket), "files": int(r.files)} for r in d["buckets"].itertuples()],
        "carriers": d["carriers"][["company", "naic_id", "policies", "confirmed_complaints", "complaint_index", "per_100k"]].to_dict("records"),
    }
    out.write_text(json.dumps(payload, indent=2, default=float))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--year", type=int, default=dt.date.today().year - 1, help="report year (default: last year)")
    ap.add_argument("--out", type=Path, default=None, help="output .xlsx path")
    ap.add_argument("--lag-days", type=int, default=120, help="months newer than this are still filling in (default 120)")
    a = ap.parse_args()
    out = a.out or Path(f"TX_Auto_Insurance_Complaints_Report_{a.year}.xlsx")
    t0 = time.time()
    d = build(a.year, dt.date.today(), a.lag_days)
    write_excel(d, out)
    save_json(d, out.with_suffix(".json"))
    print(f"Done in {time.time() - t0:.0f}s -> {out} ({d['kpi']['files']:,} complaint files)")


if __name__ == "__main__":
    main()
