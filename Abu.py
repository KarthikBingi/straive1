# streamlit_app.py
# Run this using: streamlit run streamlit_app.py

import streamlit as st
import pandas as pd
import re
import oracledb
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Post-Test DB Query Executor (IST)", layout="wide")
st.title("🇮🇳 Post-Test DB Query Execution with IST Timezone Adjustment")

# --- UI for DB connection ---
st.subheader("Database Connection Details")

db_host = st.text_input("Host", "XVU05-SCAN.SDI.CORP.BANKOFAMERICA.COM")
db_port = st.text_input("Port", "49125")
db_service = st.text_input("Service Name", "BRPIQ01_SVC01")
db_user = st.text_input("User", "RPIST_WRITE")
db_pass = st.text_input("Password", type="password")

connect_btn = st.button("Connect to Database")

conn = None
if connect_btn:
    try:
        conn_str = f"{db_user}/{db_pass}@{db_host}:{db_port}/{db_service}"
        conn = oracledb.connect(conn_str, encoding="UTF-8", nencoding="UTF-8")
        cur = conn.cursor()

        # Fetch DB timezone and current time
        cur.execute("SELECT DBTIMEZONE, SESSIONTIMEZONE, CURRENT_TIMESTAMP FROM DUAL")
        tz_info = cur.fetchone()

        db_tz, session_tz, db_time = tz_info
        st.success("✅ Successfully connected to the database!")
        st.write(f"**Database Timezone:** {db_tz}")
        st.write(f"**Session Timezone:** {session_tz}")
        st.write(f"**Current Database Time:** {db_time}")

        # Adjust session timezone to IST for this user
        cur.execute("ALTER SESSION SET TIME_ZONE = 'Asia/Kolkata'")
        cur.execute("SELECT SESSIONTIMEZONE FROM DUAL")
        new_tz = cur.fetchone()[0]
        st.info(f"Session timezone updated to: {new_tz} (IST)")

        conn.close()
    except Exception as e:
        st.error(f"❌ Failed to connect to database: {e}")

# --- Query Execution Section ---
st.subheader("Query Execution")
uploaded_file = st.file_uploader("Upload your SQL text file", type=["txt"])

st.text("Enter time range (optional, only used if queries contain &test_start_time or &test_end_time):")
start_time = st.text_input("Start Time (YYYY-MM-DD HH24:MI:SS)")
end_time = st.text_input("End Time (YYYY-MM-DD HH24:MI:SS)")

run_btn = st.button("Run Queries and Export to Excel")

# --- Helper Functions ---

def parse_queries_from_text(text):
    """
    Parses queries in formats like:
    -- Query1
    select * from table1;
    or
    Query1
    select * from table1;
    """
    queries = {}
    parts = re.split(r"(?i)(?:--\s*Query\d+|Query\d+)", text)
    query_names = re.findall(r"(?i)(?:--\s*(Query\d+)|Query\d+)", text)
    query_names = [q for q in query_names if q]
    for i, part in enumerate(parts[1:]):  # skip first empty split
        name = f"Query{i+1}" if i >= len(query_names) else query_names[i]
        queries[name.strip()] = part.strip()
    return queries


def replace_time_variables(sql, start_time, end_time):
    """
    Replaces &test_start_time and &test_end_time safely with TO_TIMESTAMP().
    Prevents :30 or similar bind variable errors.
    """
    if start_time:
        sql = re.sub(
            r"&test_start_time",
            f"TO_TIMESTAMP('{start_time}', 'YYYY-MM-DD HH24:MI:SS')",
            sql,
            flags=re.IGNORECASE,
        )
    if end_time:
        sql = re.sub(
            r"&test_end_time",
            f"TO_TIMESTAMP('{end_time}', 'YYYY-MM-DD HH24:MI:SS')",
            sql,
            flags=re.IGNORECASE,
        )
    return sql


def safe_sheet_name(name):
    return re.sub(r"[\\/*?:\[\]]", "_", name)[:31]


def execute_queries_with_ist(db_user, db_pass, db_host, db_port, db_service, queries, start_time=None, end_time=None):
    conn_str = f"{db_user}/{db_pass}@{db_host}:{db_port}/{db_service}"
    excel_io = BytesIO()
    conn = None
    try:
        conn = oracledb.connect(conn_str, encoding="UTF-8", nencoding="UTF-8")
        cur = conn.cursor()

        # Set session timezone to IST (Asia/Kolkata)
        cur.execute("ALTER SESSION SET TIME_ZONE = 'Asia/Kolkata'")

        writer = pd.ExcelWriter(excel_io, engine="xlsxwriter")

        for name, query in queries.items():
            sql = replace_time_variables(query.strip().rstrip(";"), start_time, end_time)
            try:
                cur.execute(sql)
                cols = [desc[0] for desc in cur.description] if cur.description else []
                rows = cur.fetchall()
                df = pd.DataFrame(rows, columns=cols) if cols else pd.DataFrame({"Result": ["No Data"]})
                df.to_excel(writer, sheet_name=safe_sheet_name(name), index=False)
            except Exception as e:
                df = pd.DataFrame({"Error": [str(e)]})
                df.to_excel(writer, sheet_name=safe_sheet_name(f"{name}_error"), index=False)

        writer.close()
        excel_io.seek(0)
        return excel_io

    finally:
        if conn:
            conn.close()


# --- Execution ---
if run_btn:
    if not uploaded_file:
        st.error("⚠️ Please upload your SQL text file first.")
    else:
        try:
            text_data = uploaded_file.read().decode("utf-8", errors="ignore")
            queries = parse_queries_from_text(text_data)
            if not queries:
                st.error("No queries found in file.")
            else:
                st.info(f"Executing {len(queries)} queries (in IST)...")
                excel_data = execute_queries_with_ist(
                    db_user, db_pass, db_host, db_port, db_service, queries, start_time, end_time
                )
                filename = f"db_results_IST_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.success("✅ Queries executed successfully with IST timezone!")
                st.download_button("Download Excel File", data=excel_data, file_name=filename)
        except Exception as e:
            st.error(f"❌ Error executing queries: {e}")
