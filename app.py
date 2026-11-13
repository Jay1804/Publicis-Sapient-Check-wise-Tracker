"""Streamlit app: Publicis Sapient Check wise Tracker.

Run:  streamlit run app.py
Credentials are read from .env (DB_HOST, DB_NAME, DB_USER, DB_PASSWORD).
"""
import os
import sys
from datetime import date, timedelta

import streamlit as st

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker"))
import pipeline as P  # noqa: E402
import queries as Q  # noqa: E402

st.set_page_config(page_title="Publicis Sapient Check wise Tracker", layout="wide")
st.title("Publicis Sapient – Check wise Tracker")
st.caption(f"Client: {Q.CLIENT_NAME} · live database `{os.environ.get('DB_NAME', '')}` (read-only queries)")

today = date.today()
c1, c2, c3 = st.columns([1, 1, 2])
d_from = c1.date_input("Received from", today.replace(year=today.year - 2) if not (today.month == 2 and today.day == 29)
                       else today - timedelta(days=730))
d_to = c2.date_input("Received to", today)
c3.info("Default window = last 2 years by Received date.")

if st.button("Generate tracker", type="primary"):
    status = st.status("Running…", expanded=True)
    try:
        df, summary, warns = P.build_tracker(d_from, d_to, progress=lambda m: status.write(m))
        status.update(label="Building Excel…")
        st.session_state["xlsx"] = P.to_excel_bytes(df, summary)
        st.session_state["df"], st.session_state["summary"], st.session_state["warns"] = df, summary, warns
        st.session_state["name"] = f"Publicis Sapient Check wise Tracker-{date.today().strftime('%d-%b-%y')}.xlsx"
        status.update(label=f"Done – {len(df):,} check rows", state="complete", expanded=False)
    except Exception as e:  # noqa: BLE001
        status.update(label="Failed", state="error")
        st.exception(e)

if "df" in st.session_state:
    for w in st.session_state["warns"]:
        st.warning(w)
    st.download_button("⬇ Download tracker (.xlsx)", st.session_state["xlsx"], st.session_state["name"],
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    m = st.columns(4)
    df = st.session_state["df"]
    m[0].metric("Check rows", f"{len(df):,}")
    m[1].metric("Cases", f"{df['ARS No'].nunique():,}")
    m[2].metric("OUT TAT", f"{(df['TAT_Status'] == 'OUT TAT').sum():,}")
    m[3].metric("Work in Progress", f"{(df['Check Status'] == 'Work in Progress').sum():,}")
    t1, t2 = st.tabs(["Summary (Check Unique Name × Check Severity)", "Tracker data"])
    t1.dataframe(st.session_state["summary"], use_container_width=True)
    t2.dataframe(df.head(2000), use_container_width=True)
    t2.caption("Preview of first 2,000 rows – the download has all rows.")
