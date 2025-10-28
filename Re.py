# streamlit_app.py
# Run this with: streamlit run streamlit_app.py

import streamlit as st
import pandas as pd
import re
import oracledb
from io import BytesIO
from datetime import datetime, timedelta, timezone
import pytz

st.set_page_config(page_title="Post-Test DB Query Executor", layout="wide")
st.title("Post-Test Automatic DB Query Execution")

# --- Database Connection Section ---
st.subheader("Enter Database Details")
db_host = st.text_input("Host", "XVU05-SCAN.SDI.CORP.BANKOFAMERICA.COM")
db_port = st.text_input("Port", "49125")
db_service = st.text_input("Service", "BRPIQ01_SVC01")
db_user = st.text_input("Username", "RPIST_WRITE")
db_pass = st.text_input("Password", type="password")

connect_btn = st.button("Connect to Database")

# --- Globals ---
conn = None
db_timezone_offset = None

# --- Connect to DB and Get DB Timezone ---
if connect_btn:
    try:
        conn = oracledb.connect(user=db_user, password=db_pass,
                                dsn=f"{db_host}:{db_port}/{db_service}")
        st.success("✅ Successfully connected to the database!")

        # Get DB timezone
        cur = conn.cursor()
        cur.execute("SELECT DBTIMEZONE FROM DUAL")
        db_timezone = cur.fetchone()[0]
        st.info(f"🕒 Database Timezone: {db_timezone}")

        # Convert Oracle timezone string (e.g., "-04:00") to a timezone object
        sign = 1 if db_timezone.startswith("+") else -1
        hours, mins = map(int, db_timezone[1:].split(":"))
        db_timezone_offset = timezone(timedelta(hours=sign * hours, minutes=sign * mins))

        # Store for later use
        st.session_state["db_connection"] = conn
        st.session_state["db_timezone"] = db_timezone_offset

    except Exception as e:
        st.error(f"❌ Failed to connect to the database: {e}")

# --- Query Upload Section ---
st.subheader("Upload Query File")
uploaded_file = st.file_uploader("Upload your query text file", type=["txt"])

# --- Time Input Section ---
st.subheader("Enter Time Range (IST)")
start_time_str = st.text_input("Start Time (YYYY-MM-DD HH:MM:SS)")
end_time_str = st.text_input("End Time (YYYY-MM-DD HH:MM:SS)")

run_button = st.button("Run Queries & Generate Excel")

# --- Helper Functions ---

def parse_queries_from_text(text: str):
    """
    Parse queries even if they start with '-- Query1' or 'Query1'
    Works regardless of case sensitivity.
    """
    pattern = r"(--\s*Query\d+|\bQuery\d+\b)"
    parts = re.split(pattern, text, flags=re.IGNORECASE)
    query_names = re.findall(pattern, text, flags=re.IGNORECASE)

    queries = {}
    for i, qname in enumerate(query_names):
        name = re.sub(r"[^a-zA-Z0-9]", "", qname)
        queries[name] = parts[i + 1].strip() if i + 1 < len(parts) else ""
    return queries


def convert_ist_to_db_time(ist_time_str, db_offset):
    """
    Convert IST timestamp to the database timezone timestamp string.
    """
    ist = pytz.timezone("Asia/Kolkata")
    db_tz = timezone(db_offset.utcoffset(None))
    local_time = ist.localize(datetime.strptime(ist_time_str, "%Y-%m-%d %H:%M:%S"))
    db_time = local_time.astimezone(db_tz)
    return db_time.strftime("%Y-%m-%d %H:%M:%S")


def replace_time_variables(sql, start_time, end_time):
    if start_time:
        sql = re.sub(r"&test_start_time", f"'{start_time}'", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"'{end_time}'", sql, flags=re.IGNORECASE)
    return sql


def safe_sheet_name(name):
    return re.sub(r"[\\/*?:\[\]]", "_", name)[:31]


def execute_queries(queries, conn, db_offset, start_time_str, end_time_str):
    cur = conn.cursor()
    excel_io = BytesIO()
    writer = pd.ExcelWriter(excel_io, engine="xlsxwriter")

    # Convert IST inputs to DB timezone equivalents
    db_start_time = convert_ist_to_db_time(start_time_str, db_offset) if start_time_str else None
    db_end_time = convert_ist_to_db_time(end_time_str, db_offset) if end_time_str else None

    st.info(f"🕓 Using DB Time Range: {db_start_time} → {db_end_time}")

    for name, query in queries.items():
        sql = replace_time_variables(query.strip(), db_start_time, db_end_time)
        try:
            cur.execute(sql)
            cols = [d[0] for d in cur.description] if cur.description else []
            rows = cur.fetchall()
            df = pd.DataFrame(rows, columns=cols) if cols else pd.DataFrame({"Result": ["No Data"]})
            df.to_excel(writer, sheet_name=safe_sheet_name(name), index=False)
        except Exception as e:
            err_df = pd.DataFrame({"Error": [str(e)]})
            err_df.to_excel(writer, sheet_name=safe_sheet_name(f"{name}_error"), index=False)

    writer.close()
    excel_io.seek(0)
    return excel_io


# --- Run Queries ---
if run_button:
    if "db_connection" not in st.session_state or "db_timezone" not in st.session_state:
        st.error("❌ Please connect to the database first.")
    elif not uploaded_file:
        st.error("❌ Please upload a query text file.")
    elif not start_time_str or not end_time_str:
        st.error("❌ Please enter both start and end times.")
    else:
        text_data = uploaded_file.read().decode("utf-8", errors="ignore")
        queries = parse_queries_from_text(text_data)
        if not queries:
            st.error("No queries found in the uploaded file.")
        else:
            st.info(f"Executing {len(queries)} queries...")
            try:
                excel_data = execute_queries(
                    queries,
                    st.session_state["db_connection"],
                    st.session_state["db_timezone"],
                    start_time_str,
                    end_time_str,
                )
                filename = f"posttest_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.success("✅ Queries executed successfully!")
                st.download_button("📂 Download Excel Results", data=excel_data, file_name=filename)
            except Exception as e:
                st.error(f"Error executing queries: {e}")
