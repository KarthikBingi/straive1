import os
import re
import pytz
import pandas as pd
import streamlit as st
import oracledb
from datetime import datetime
from io import BytesIO
from streamlit_datetime_picker import date_time_picker

# ---------- Optional: initialize Oracle client (update path if required) ----------
try:
    oracledb.init_oracle_client(lib_dir=r"C:\ORACLE19_X64\PRODUCT\19.3.0\client_1\bin")
except Exception:
    # If client not required or already configured, ignore
    pass

# ----------------- Helpers ----------------- #
def map_timezone(tz):
    tz_map = {
        "+05:30": "Indian Standard Time (IST)",
        "-04:00": "Eastern Time (ET)",
        "-05:00": "Central Time (CST)",
        "+00:00": "Greenwich Mean Time (GMT)",
        "+01:00": "Central European Time (CET)"
    }
    return tz_map.get(tz, tz or "Unknown Time Zone")


def test_connection_and_get_tz(host, port, service_name, username, password):
    dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
    conn = oracledb.connect(user=username, password=password, dsn=dsn)
    cur = conn.cursor()
    cur.execute("SELECT DBTIMEZONE FROM DUAL")
    db_tz = cur.fetchone()[0]
    cur.close()
    conn.close()
    return db_tz


def replace_time(sql, start_time, end_time):
    if start_time:
        sql = re.sub(r"&test_start_time", f"{start_time}", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"{end_time}", sql, flags=re.IGNORECASE)
    return sql


def execute_single_query(conn_details, sql_query):
    host, port, service_name, username, password = conn_details
    dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
    conn = oracledb.connect(user=username, password=password, dsn=dsn)
    cur = conn.cursor()
    cur.execute(sql_query)
    cols = [c[0] for c in cur.description] if cur.description else []
    rows = cur.fetchall()
    cur.close()
    conn.close()
    df = pd.DataFrame(rows, columns=cols)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Result")
    out.seek(0)
    return out


def execute_file_queries(conn_details, sql_file_path, start_time, end_time):
    host, port, service_name, username, password = conn_details
    dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
    conn = oracledb.connect(user=username, password=password, dsn=dsn)
    cur = conn.cursor()
    with open(sql_file_path, "r") as f:
        sql_statements = [line.strip().rstrip(';') for line in f if line.strip() and not line.strip().startswith("--")]

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, raw_sql in enumerate(sql_statements, start=1):
            sql = replace_time(raw_sql, start_time, end_time)
            cur.execute(sql)
            cols = [c[0] for c in cur.description] if cur.description else []
            rows = cur.fetchall()
            df = pd.DataFrame(rows, columns=cols)
            sheet_name = f"Query_{i}" if len(sql_statements) > 1 else "Result"
            # Excel sheet names max 31 chars
            writer.book.create_sheet(sheet_name) if sheet_name not in writer.book.sheetnames else None
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    cur.close()
    conn.close()
    out.seek(0)
    return out


def convert_and_display(user_time, user_tz):
    """Return localized strings for CST and ET and also return the CST string for SQL replacement."""
    try:
        user_tz_obj = pytz.timezone(user_tz)
        cst_tz = pytz.timezone("US/Central")
        et_tz = pytz.timezone("US/Eastern")

        if isinstance(user_time, datetime):
            user_dt = user_time
        else:
            user_dt = datetime.strptime(user_time, "%Y-%m-%d %H:%M:%S")

        if user_dt.tzinfo is None:
            user_dt = user_tz_obj.localize(user_dt)

        cst_dt = user_dt.astimezone(cst_tz)
        et_dt = user_dt.astimezone(et_tz)
        return cst_dt.strftime("%Y-%m-%d %H:%M:%S"), et_dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        return None, None

# ----------------- Streamlit UI ----------------- #
st.set_page_config(page_title="Oracle SQL Executor", layout="centered")
st.markdown("<h1 style='text-align:center;'>Oracle SQL Executor</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- Connection panel ---
st.markdown("### 🔌 Connection")
col1, col2 = st.columns(2)
with col1:
    host = st.text_input("Host", placeholder="hostname or IP")
    port = st.text_input("Port", placeholder="e.g. 1521 or 49125")
    service_name = st.text_input("Service Name", placeholder="SERVICE_NAME")
with col2:
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

if "conn_details" not in st.session_state:
    st.session_state.conn_details = None
if "db_tz" not in st.session_state:
    st.session_state.db_tz = None
if "db_tz_readable" not in st.session_state:
    st.session_state.db_tz_readable = None

if st.button("Connect"):
    # validate inputs
    if not (host and port and service_name and username and password):
        st.error("Please fill all connection fields.")
    else:
        try:
            db_tz = test_connection_and_get_tz(host, port, service_name, username, password)
            st.session_state.conn_details = (host, port, service_name, username, password)
            st.session_state.db_tz = db_tz
            st.session_state.db_tz_readable = map_timezone(db_tz)
            st.success(f"Connected. DBTIMEZONE: {db_tz} ({st.session_state.db_tz_readable})")
        except Exception as e:
            st.session_state.conn_details = None
            st.session_state.db_tz = None
            st.session_state.db_tz_readable = None
            st.error(f"Connection failed: {e}")

st.markdown("---")

# --- Execution Mode ---
st.markdown("### 🧾 Execution Mode")
mode = st.radio("", ["Single Query", "Browse File"], horizontal=True)

# --- SINGLE QUERY MODE ---
if mode == "Single Query":
    st.markdown("#### Execute a single SQL query")
    st.info("Only a single statement is expected here. Results will be exported to Excel for download.")
    sql_query = st.text_area("Enter SQL Query", height=180)
    run_single = st.button("Run Query")
    if run_single:
        if st.session_state.conn_details is None:
            st.error("Not connected to DB. Please connect first.")
        elif not sql_query.strip():
            st.error("Please enter a SQL query.")
        else:
            try:
                excel_io = execute_single_query(st.session_state.conn_details, sql_query)
                st.success("Query executed successfully.")
                st.download_button(
                    "⬇️ Download result.xlsx",
                    data=excel_io,
                    file_name="single_query_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Execution failed: {e}")

# --- BROWSE FILE MODE ---
else:
    st.markdown("#### Execute SQLs from uploaded .txt file")
    sql_file = st.file_uploader("Upload SQL file (.txt)", type=["txt"])
    available_timezones = pytz.all_timezones
    user_timezone = st.selectbox("Select your timezone", available_timezones, index=available_timezones.index("UTC"))
    start_time_input = date_time_picker(label="Select Start Time", key="start_picker")
    end_time_input = date_time_picker(label="Select End Time", key="end_picker")

    # Show converted times as soon as inputs provided
    start_cst, start_et = (None, None)
    end_cst, end_et = (None, None)
    if start_time_input:
        start_cst, start_et = convert_and_display(start_time_input, user_timezone)
        if start_cst and start_et:
            st.info(f"Start Time → CST: **{start_cst}** | ET: **{start_et}**")
    if end_time_input:
        end_cst, end_et = convert_and_display(end_time_input, user_timezone)
        if end_cst and end_et:
            st.info(f"End Time → CST: **{end_cst}** | ET: **{end_et}**")

    run_file = st.button("Execute SQLs from File")

    if run_file:
        if st.session_state.conn_details is None:
            st.error("Not connected to DB. Please connect first.")
        elif sql_file is None:
            st.error("Please upload a .txt file containing SQL statements.")
        else:
            try:
                # save uploaded file temporarily
                local_path = f"uploaded_{sql_file.name}"
                with open(local_path, "wb") as f:
                    f.write(sql_file.getbuffer())

                # Use CST values (if available) for placeholder replacement; else None
                start_for_sql = start_cst if start_cst else None
                end_for_sql = end_cst if end_cst else None

                excel_io = execute_file_queries(st.session_state.conn_details, local_path, start_for_sql, end_for_sql)
                st.success("File executed successfully.")
                st.download_button(
                    "⬇️ Download results.xlsx",
                    data=excel_io,
                    file_name="file_query_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # cleanup
                try:
                    os.remove(local_path)
                except Exception:
                    pass
            except Exception as e:
                st.error(f"Execution failed: {e}")

st.markdown("---")
if st.session_state.db_tz:
    st.markdown(f"**Connected DB Time Zone:** `{st.session_state.db_tz}` — *{st.session_state.db_tz_readable}*")
else:
    st.markdown("**Connected DB Time Zone:** Not connected")

st.markdown("<p style='text-align:center;color:gray;'>© 2025 SQL Executor</p>", unsafe_allow_html=True)
