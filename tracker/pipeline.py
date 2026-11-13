"""Build the Publicis Sapient Check-wise tracker from the live DB."""
from __future__ import annotations

import io
import os
import re
import threading
import warnings
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import numpy as np
import pandas as pd
import pymysql
from dotenv import load_dotenv

import queries as Q

warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(BASE, ".env"))

# ---- output layout (identical to "Publicis Sapient Check wise Tracker-<date>.xlsx") ----
COLUMNS = [
    "Case Check Id", "Candidate Name", "Check_Status", "Client Name", "ARS No", "Process Name", "Location",
    "Received date", "Initiation date", "Flexi Field 1", "Flexi Field 2", "Flexi Field 3", "Flexi Field 4",
    "Flexi Field 5", "Flexi Field 6", "Flexi Field 7", "Ageing", "case_created_date", "case_due_date",
    "Case_status", "Case Latest Severity", "Latest Case Reopen date", "Check Created date", "Check Due date",
    "Check Name", "Check Unique Name", "First Insuff date", "Latest Insuff date", "Insuff Fulfilled date",
    "Go Ahead date", "Check Status", "Closure Comments", "Disposition Name", "Ver_Overseas", "Ver_Summary",
    "Ver_Type", "Verification_source", "New Check Added date", "Check Severity", "Check Closure date",
    "Check Reopen date", "Insuff Remarks", "TAT_Status", "Insuff_Ageing", "Insuff Ageing Bucket", "Department",
    "Flat_TAT", "Max_date", "Re_Cal_Tat", "email_address", "mobile_no", "DataSource", "Case Source",
    "Latest Interim Report Sent Date", "Latest Interim Report Severity", "Letest Report Sent Date",
]
DATE_COLS = ["Received date", "case_created_date", "case_due_date", "Latest Case Reopen date",
             "Check Created date", "Check Due date", "First Insuff date", "Latest Insuff date",
             "Insuff Fulfilled date", "Go Ahead date", "Check Closure date", "Check Reopen date", "Max_date"]
DATETIME_COLS = ["Initiation date", "New Check Added date", "Latest Interim Report Sent Date", "Letest Report Sent Date"]

WIP_SOURCE = {"awaiting response", "new/unassigned", "work in progress", "escalated", "in research"}
SEVERITY_ORDER = ["Attention Required", "Clear", "Discrepant", "Green", "Insufficient", "No Response Received",
                  "Submission Done", "Not Applicable", "Work in Progress", "case closed by client", "On Hold"]

# "Notes" sheet of the template: DataSource -> Case Source
CASE_SOURCE_MAP = {
    "RuleEngine - iBridge Quick Flow": "API - Snow",
    "RuleEngine-No JAF": "Offline - Email",
    "RuleEngine-Digital JAF": "Offline - Email",
    "ExcelCase": "Offline - Email",
    "RuleEngine-Handwritten JAF": "Offline - Email",
    "RuleEngine": "Offline - Email",
    "Mail": "Offline - Email",
    "ClientPortal": "Offline - Email",
    "iBridge": "Ibridge",
}


def connect():
    return pymysql.connect(host=os.environ["DB_HOST"], user=os.environ["DB_USER"],
                           password=os.environ["DB_PASSWORD"], database=os.environ["DB_NAME"],
                           connect_timeout=20, read_timeout=900)


def _chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def _read(sql, params=None, conn=None):
    return pd.read_sql(sql, conn or _conn(), params=params)


_POOL = ThreadPoolExecutor(8)
_TL = threading.local()


def _conn():
    """One persistent connection per worker thread (avoids ~1 s SSL handshake per query)."""
    c = getattr(_TL, "c", None)
    try:
        if c is not None:
            c.ping(reconnect=True)
            return c
    except Exception:  # noqa: BLE001
        pass
    c = _TL.c = connect()
    return c


def _submit(template, ids, size=2000):
    """Queue one query per id-chunk on the pool; returns futures (collect with _gather)."""
    ids = [int(i) for i in ids]
    return [_POOL.submit(lambda p=p: pd.read_sql(template.format(ids=",".join(map(str, p))), _conn()))
            for p in _chunks(ids, size)]


def _gather(futs):
    res = [f.result() for f in futs]
    return pd.concat(res, ignore_index=True) if res else pd.DataFrame()


def _by_ids(template, ids, size=2000):
    return _gather(_submit(template, ids, size))


def _fn_map(fn, codes, conn):
    """Evaluate a DB scalar function once per distinct code."""
    codes = [int(c) for c in pd.unique(pd.Series(codes).dropna())]
    if not codes:
        return {}
    cur = conn.cursor()
    cur.execute("SELECT " + ",".join(f"{fn}({c})" for c in codes))
    return dict(zip(codes, cur.fetchone()))


def _networkdays(pairs, conn=None):
    """DB networkdays() (holiday aware) evaluated once per distinct (start,end) pair, in parallel."""
    pairs = list(dict.fromkeys(pairs))

    def run(part):
        cur = _conn().cursor()
        cur.execute("SELECT " + ",".join(f"checkpoint_live.networkdays('{a}','{b}')" for a, b in part))
        return dict(zip(part, cur.fetchone()))
    out = {}
    for r in [f.result() for f in [_POOL.submit(run, part) for part in _chunks(pairs, 300)]]:
        out.update(r)
    return out


def _d(s):
    return pd.to_datetime(s, errors="coerce")


def excel_networkdays(start: date, end: date) -> int:
    """Excel NETWORKDAYS (Mon-Fri, inclusive, negative when start > end)."""
    a, b = np.datetime64(start, "D"), np.datetime64(end, "D")
    return int(np.busday_count(a, b + 1)) if a <= b else -int(np.busday_count(b, a + 1))


def build_tracker(d_from: date, d_to: date, client: str = Q.CLIENT_NAME, progress=lambda m: None):
    """Returns (tracker_df, summary_df, warnings_list)."""
    warn = []
    today = date.today()
    conn = connect()
    try:
        progress("Fetching cases…")
        cid = _read(Q.COMPANY_IDS, {"client": client}, conn).company_id.astype(int).tolist()
        cases = _read(Q.CASES.format(ids=",".join(map(str, cid)), d_from=d_from.isoformat(), d_to=d_to.isoformat()), None, conn)
        cases = cases.drop_duplicates("case_id")
        case_ids = cases.case_id.tolist()
        progress(f"{len(case_ids)} cases found. Fetching everything in parallel…")
        uniq = lambda col: cases[col].dropna().unique()
        f_cand, f_proc = _submit(Q.CANDIDATES, uniq("candidate_id")), _submit(Q.PROCESSES, uniq("process_id"))
        f_loc, f_fields = _submit(Q.LOCATIONS, uniq("client_office_id")), _submit(Q.FIELDS, case_ids)
        f_hist = _submit(Q.HISTORY, case_ids, 300)           # slowest query -> start first
        f_checks = _submit(Q.CHECKS, case_ids, 1000)
        f_rep = _submit(Q.REPORTS, case_ids)

        cases = cases.merge(_gather(f_cand).drop_duplicates("candidate_id"), on="candidate_id", how="left")
        cases = cases[~cases.first_name.fillna("").str.lower().str.contains("dummy")]
        cases["candidate_name"] = (cases.first_name.fillna("") + " " + cases.middle_name.fillna("") + " " + cases.last_name.fillna(""))
        cases = cases.merge(_gather(f_proc).drop_duplicates("process_id"), on="process_id", how="left")
        cases = cases.merge(_gather(f_loc).drop_duplicates("office_id"), left_on="client_office_id",
                            right_on="office_id", how="left")
        cases = cases.merge(_gather(f_fields).drop_duplicates("case_id"), on="case_id", how="left")
        checks = _gather(f_checks)
        checks = checks[checks.case_id.isin(cases.case_id)]
        checks["go_ahead_date"] = _d(checks["go_ahead_date"])
        chk_ids = checks.case_check_id.tolist()
        progress(f"{len(checks)} checks. Fetching insufficiency / sources / reports / history…")
        f_ins, f_src = _submit(Q.INSUFF, chk_ids, 3000), _submit(Q.SOURCE, chk_ids, 3000)
        insuff, source = _gather(f_ins), _gather(f_src)
        reports, hist = _gather(f_rep), _gather(f_hist)

        # ---- insuff: first / latest per check (by insuff_id) ----
        for c_ in ("Insuff_raised_date", "Insuff_fulfil_date"):
            insuff[c_] = _d(insuff[c_])
        insuff = insuff.sort_values(["case_check_id", "insuff_id"])
        g = insuff.groupby("case_check_id")
        ins = pd.DataFrame({
            "First Insuff date": g["Insuff_raised_date"].first(),
            "Latest Insuff date": g["Insuff_raised_date"].last(),
            "Insuff Fulfilled date": g["Insuff_fulfil_date"].last(),
        })
        # ---- reports ----
        reports = reports.sort_values("case_report_id", ascending=False)
        latest_any = reports.drop_duplicates("case_id").set_index("case_id")["report_severity"]
        interim = reports[reports.report_type == 0].drop_duplicates("case_id").set_index("case_id")
        final = reports[reports.report_type == 1].drop_duplicates("case_id").set_index("case_id")
        # ---- history ----
        newchk = hist.sort_values("action_id").drop_duplicates("case_id", keep="last").set_index("case_id")["action_taken_on"]
        # ---- verification source name ----
        source = source.sort_values(["case_check_id", "org_id"]).drop_duplicates("case_check_id")
        org_ids = source.org_id.dropna().astype(int).unique().tolist()
        names = {k: _by_ids(t, org_ids).drop_duplicates("id").set_index("id")["name"]
                 for k, t in (("co", Q.ORG_COMPANY), ("inst", Q.ORG_INSTITUTE), ("city", Q.ORG_CITY))}
        src_org = source.set_index("case_check_id")["org_id"]

        # ---- assemble ----
        df = checks.merge(cases, on="case_id", how="inner")
        df = df.join(ins, on="case_check_id")
        goahead = checks.groupby("case_id")["go_ahead_date"].max()
        df["Go Ahead date"] = df.case_id.map(goahead)

        def vsrc(r):
            org = src_org.get(r.case_check_id)
            if pd.isna(org):
                return None
            fam = r.family_id
            key = "co" if fam == 5 else "inst" if fam == 4 else "city"
            return names[key].get(int(org))
        df["Verification_source"] = df.apply(vsrc, axis=1)

        progress("Resolving status codes…")
        cs_map = _fn_map("checkpoint_live.fn_check_status", df.check_status, conn)
        case_status_map = _fn_map("checkpoint_live.fn_case_status", df.case_status, conn)
        ver_map = _fn_map("checkpoint_live.fn_ver_procedure", df.VER_SUMMARY, conn)
        ds_map = _fn_map("fn_case_data_source", df.DATA_SOURCE, conn)

        out = pd.DataFrame(index=df.index)
        out["Case Check Id"] = df.case_check_id
        out["Candidate Name"] = df.candidate_name
        out["Check_Status"] = df.check_status.map(cs_map)                       # step 4: untouched
        out["Client Name"] = df.company_name
        out["ARS No"] = df.case_ars_no
        out["Process Name"] = df.process_name
        out["Location"] = df.office_name
        out["Received date"] = _d(df.received_date)
        out["Initiation date"] = _d(df.case_accepted_date)
        for i in range(1, 8):                                                   # step 6: flexi 8-12 removed
            out[f"Flexi Field {i}"] = df[f"case_flex_field{i}"]
        out["case_created_date"] = _d(df.case_created_date)
        out["case_due_date"] = _d(df.case_due_date)
        out["Case_status"] = df.case_status.map(case_status_map)
        out["Case Latest Severity"] = df.case_id.map(latest_any)
        out["Latest Case Reopen date"] = _d(df.case_reopen_date)
        out["Check Created date"] = _d(df.check_created_on)
        out["Check Due date"] = _d(df.CHECK_EXPECTED_CLOSURE)
        out["Check Name"] = df.check_name
        out["Check Unique Name"] = df.check_unique_name
        out["First Insuff date"] = _d(df["First Insuff date"])
        out["Latest Insuff date"] = _d(df["Latest Insuff date"])
        out["Insuff Fulfilled date"] = _d(df["Insuff Fulfilled date"])
        out["Go Ahead date"] = _d(df["Go Ahead date"])
        # step 3: Check Status -> "Work in Progress" for the open statuses
        raw = out["Check_Status"]
        out["Check Status"] = raw.where(~raw.str.strip().str.lower().isin(WIP_SOURCE), "Work in Progress")
        out["Closure Comments"] = df.closure_comments
        out["Disposition Name"] = df.disposition_name
        out["Ver_Overseas"] = df.VER_OVERSEAS.map({0: "No", 1: "Yes", "0": "No", "1": "Yes"})
        out["Ver_Summary"] = df.VER_SUMMARY.map(ver_map)
        out["Ver_Type"] = df.VER_TYPE.map({1: "Verbal", 2: "Written", 3: "VRWP", "1": "Verbal", "2": "Written", "3": "VRWP"})
        out["Verification_source"] = df["Verification_source"]
        out["New Check Added date"] = _d(df.case_id.map(newchk))
        # step 5: blank Check Severity <- Check Status
        sev = df.check_severity.astype("object")
        blank = sev.isna() | (sev.astype(str).str.strip() == "")
        out["Check Severity"] = sev.where(~blank, out["Check Status"])
        out["Check Closure date"] = _d(df.check_closure_date)
        out["Check Reopen date"] = _d(df.check_reopen_date)
        out["Insuff Remarks"] = df.insuff_remarks
        out["Department"] = df.department
        out["email_address"] = df.email_address
        out["mobile_no"] = df.mobile_no
        out["DataSource"] = df.DATA_SOURCE.map(ds_map)
        out["Case Source"] = out["DataSource"].map(CASE_SOURCE_MAP)
        out["Latest Interim Report Sent Date"] = _d(df.case_id.map(interim["report_sent_on"]))
        out["Latest Interim Report Severity"] = df.case_id.map(interim["report_severity"])
        out["Letest Report Sent Date"] = _d(df.case_id.map(final["report_sent_on"]))

        for c_ in DATE_COLS:          # original query used DATE_FORMAT(...,'%d-%b-%Y') -> date only
            if c_ != "Max_date":
                out[c_] = out[c_].dt.normalize()

        # ---- row filters (steps 1 & 2) ----
        before = len(out)
        out = out[out["Check_Status"].fillna("").str.strip().str.lower() != "disabled"]
        out = out[out["Check Unique Name"].fillna("").str.strip().str.lower() != "site visit"]
        progress(f"Removed {before - len(out)} Disabled / Site Visit rows.")

        # ---- TAT metrics ----
        progress("Calculating TAT (DB networkdays)…")
        today_ts = pd.Timestamp(today)
        created = out["Check Created date"].dt.normalize()
        closed = out["Check Closure date"].dt.normalize()
        end = closed.fillna(today_ts)
        out["Max_date"] = pd.concat([created, out["Check Reopen date"].dt.normalize(),
                                     out["Go Ahead date"].dt.normalize(),
                                     out["Insuff Fulfilled date"].dt.normalize()], axis=1).max(axis=1)
        fmt = lambda s: s.dt.strftime("%Y-%m-%d")
        flat_pairs = list(zip(fmt(created), fmt(end)))
        recal_pairs = list(zip(fmt(out["Max_date"]), fmt(end)))
        ins_pairs = [(d, today.isoformat()) for d in fmt(out["Latest Insuff date"].dt.normalize()).dropna().unique()]
        valid = [p for p in flat_pairs + recal_pairs if isinstance(p[0], str) and isinstance(p[1], str)]
        nd = _networkdays(valid + ins_pairs, conn)
        out["Flat_TAT"] = [nd.get(p) for p in flat_pairs]
        out["Re_Cal_Tat"] = [nd.get(p) for p in recal_pairs]
        out["Insuff_Ageing"] = [nd.get((d, today.isoformat())) if isinstance(d, str) else None
                                for d in fmt(out["Latest Insuff date"].dt.normalize())]
        a = pd.to_numeric(out["Insuff_Ageing"])
        out["Insuff Ageing Bucket"] = np.select(
            [a.between(0, 3), a.between(4, 5), a.between(6, 9), a.between(10, 14), a >= 15],
            ["0 to 3", "4 to 5", "6 to 9", "10 to 14", "15+"], default=None)
        out.loc[a.isna(), "Insuff Ageing Bucket"] = None

        # ---- Ageing (Daily Reports-Ageing, col Q) & TAT_Status (col AQ) ----
        tat = _read(Q.TAT, {"client": client}, conn)
        tat_map = tat.dropna(subset=["formula"]).drop_duplicates("formula").set_index("formula")["new_tat_day"]
    finally:
        conn.close()

    ageing = []
    missing = set()
    for ars, proc, due, clo in zip(out["ARS No"], out["Process Name"], out["Check Due date"], out["Check Closure date"]):
        days = tat_map.get(str(ars)[:4] + str(proc))
        if days is None or pd.isna(days):
            missing.add(str(proc))
            ageing.append(None)
            continue
        base = today if pd.isna(due) else due.date()
        e = today if pd.isna(clo) else clo.date()
        ageing.append(max(0, excel_networkdays(base - timedelta(days=int(days)), e)))
    out["Ageing"] = ageing
    if missing:
        warn.append(f"No TAT configuration found for process(es): {', '.join(sorted(missing))} – Ageing left blank.")

    def tat_status(due, clo):
        if pd.isna(due):
            return "IN TAT"
        e = today_ts if pd.isna(clo) else clo
        return "IN TAT" if (e - due).days <= 0 else "OUT TAT"
    out["TAT_Status"] = [tat_status(d, c) for d, c in zip(out["Check Due date"], out["Check Closure date"])]

    unmapped = out.loc[out["Case Source"].isna() & out["DataSource"].notna(), "DataSource"].unique()
    if len(unmapped):
        warn.append("DataSource with no Case Source mapping (left blank): " + ", ".join(map(str, unmapped)))

    out = out[COLUMNS].sort_values(["Received date", "Case Check Id"], ascending=[False, False]).reset_index(drop=True)

    summary = (out.pivot_table(index="Check Unique Name", columns="Check Severity", values="Case Check Id",
                               aggfunc="count", fill_value=0))
    cols = [c for c in SEVERITY_ORDER if c in summary.columns] + sorted(c for c in summary.columns if c not in SEVERITY_ORDER)
    summary = summary[cols]
    summary["Grand Total"] = summary.sum(axis=1)
    summary.loc["Grand Total"] = summary.sum()
    return out, summary, warn


NUMERIC_TEXT_COLS = ["Flexi Field 1", "mobile_no"]


def _num_or_text(v):
    """Clean NBSP/whitespace; digit-only values (no leading zero, <=15 digits) become numbers."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    t = str(v).replace("\xa0", " ").strip()
    if t == "":
        return None
    return int(t) if re.fullmatch(r"[1-9]\d{0,14}", t) else t


def _col_letter_(n):
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def to_excel_bytes(df: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter",
                        engine_kwargs={"options": {"strings_to_urls": False, "nan_inf_to_errors": True}}) as xw:
        wb = xw.book
        # --- Summary ---
        ws = wb.add_worksheet("Summary")
        xw.sheets["Summary"] = ws
        base = {"font_name": "Calibri", "font_size": 10, "border": 1, "border_color": "#BFBFBF", "valign": "vcenter"}
        f_corner = wb.add_format({**base, "bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF"})
        f_head = wb.add_format({**base, "bold": True, "bg_color": "#2F75B5", "font_color": "#FFFFFF",
                                "align": "center", "text_wrap": True})
        f_lab = [wb.add_format({**base}), wb.add_format({**base, "bg_color": "#DDEBF7"})]
        f_num = [wb.add_format({**base, "align": "center"}), wb.add_format({**base, "align": "center", "bg_color": "#DDEBF7"})]
        f_tot_l = wb.add_format({**base, "bold": True, "bg_color": "#BDD7EE", "top": 2, "top_color": "#1F4E78"})
        f_tot_n = wb.add_format({**base, "bold": True, "bg_color": "#BDD7EE", "align": "center", "top": 2, "top_color": "#1F4E78"})
        f_gt = [wb.add_format({**base, "bold": True, "align": "center", "bg_color": "#BDD7EE"}),
                wb.add_format({**base, "bold": True, "align": "center", "bg_color": "#BDD7EE"})]
        ws.hide_gridlines(2)
        ws.set_tab_color("#2F75B5")
        ws.set_row(0, 22)
        ws.set_row(1, 32)
        ws.write(0, 0, "Count of Check Severity", f_corner)
        ws.write(0, 1, "Column Labels", f_corner)      # no merged cells: they are not allowed inside a pivot
        for j in range(2, len(summary.columns) + 1):
            ws.write_blank(0, j, None, f_corner)
        ws.write(1, 0, "Check Unique Name", f_corner)
        for j, c in enumerate(summary.columns, start=1):
            ws.write(1, j, c, f_head)
        for i, (name, row) in enumerate(summary.iterrows(), start=2):
            total = name == "Grand Total"
            ws.write(i, 0, name, f_tot_l if total else f_lab[i % 2])
            for j, v in enumerate(row, start=1):
                last = j == len(row)
                f = f_tot_n if total else f_gt[i % 2] if last else f_num[i % 2]
                ws.write(i, j, int(v) if v else "", f)
        ws.set_column(0, 0, 58)
        ws.set_column(1, len(summary.columns), 15)
        ws.freeze_panes(2, 1)

        # --- data sheet ---
        ws = wb.add_worksheet("Publicis Sapient")
        xw.sheets["Publicis Sapient"] = ws
        head = wb.add_format({"font_name": "Calibri", "font_size": 9, "align": "center", "valign": "vcenter",
                              "border": 1, "bg_color": "#5B9BD5", "font_color": "#FFFFFF", "bold": True})
        f_date = wb.add_format({"num_format": "d-mmm-yy"})
        f_dt = wb.add_format({"num_format": "d-mmm-yy h:mm"})
        ws.write_row(0, 0, list(df.columns), head)
        # column formats are inherited by cells written without a format -> fast row-wise write_row
        for j, c in enumerate(df.columns):
            if c in DATE_COLS:
                ws.set_column(j, j, 20.7, f_date)
            elif c in DATETIME_COLS:
                ws.set_column(j, j, 20.7, f_dt)
            else:
                ws.set_column(j, j, 20.7)
        df = df.copy()
        for c_ in NUMERIC_TEXT_COLS:      # "123" stored as text -> real number (Excel green-triangle error)
            df[c_] = df[c_].map(_num_or_text)
        cols = [df[c].astype(object).where(df[c].notna(), None).tolist() for c in df.columns]
        for i, row in enumerate(zip(*cols), start=1):
            ws.write_row(i, 0, row)
        ws.set_column(3, 3, 45.7)
        ws.freeze_panes(1, 0)
        # any remaining text that merely looks like a number (e.g. leading zeros) -> no error indicator
        ws.ignore_errors({"number_stored_as_text": f"A2:{_col_letter_(len(df.columns))}{len(df) + 1}"})
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)
        ws.activate()
    from pivot import add_pivot
    return add_pivot(buf.getvalue(), COLUMNS, summary, len(df))
