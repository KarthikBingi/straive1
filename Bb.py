import os
import re
import pytz
import pandas as pd
import streamlit as st
import oracledb
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
from streamlit_datetime_picker import date_time_picker

# Initialize Oracle client
oracledb.init_oracle_client(lib_dir=r"C:\\ORACLE19_X64\\PRODUCT\\19.3.0\\client_1\\bin")

# ------------------ Utility Functions ------------------ #
def replace_time(sql, start_time, end_time):
    if start_time:
        sql = re.sub(r"&test_start_time", f"{start_time}", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"{end_time}", sql, flags=re.IGNORECASE)
    return sql


def test_connection(host, port, service_name, username, password):
    """Test DB connection and return timezone."""
    try:
        dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
        conn = oracledb.connect(user=username, password=password, dsn=dsn)
        cursor = conn.cursor()
        cursor.execute("SELECT DBTIMEZONE FROM DUAL")
        tz = cursor.fetchone()[0]
        conn.close()
        readable_tz = map_timezone(tz)
        return True, f"✅ Connected successfully! Database Time Zone: {tz} ({readable_tz})", tz
    except Exception as e:
        return False, f"❌ Connection failed: {e}", None


def map_timezone(tz):
    """Map Oracle timezone offset to readable name."""
    tz_map = {
        "+05:30": "Indian Standard Time (IST)",
        "-04:00": "Eastern Time (ET)",
        "-05:00": "Central Time (CST)",
        "+00:00": "Greenwich Mean Time (GMT)",
        "+01:00": "Central European Time (CET)"
    }
    return tz_map.get(tz, "Unknown Time Zone")


def execute_sql(conn, sql_query):
    """Execute single SQL query and return Excel bytes."""
    cursor = conn.cursor()
    cursor.execute(sql_query)
    columns = [col[0] for col in cursor.description]
    rows = cursor.fetchall()

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(writer, sheet_name="Result", index=False)
    writer.close()
    output.seek(0)
    cursor.close()
    return output


def execute_sqls_from_file(conn, sql_file_path, start_time, end_time):
    """Execute multiple SQLs from file and return Excel bytes."""
    cursor = conn.cursor()
    with open(sql_file_path, "r") as file:
        sql_statements = [
            line.strip().rstrip(';') for line in file if line.strip() and not line.strip().startswith("--")
        ]

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    for i, sql in enumerate(sql_statements, start=1):
        sql = replace_time(sql, start_time, end_time)
        cursor.execute(sql)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        df = pd.DataFrame(rows, columns=columns)
        df.to_excel(writer, sheet_name=f"Query_{i}", index=False)

    writer.close()
    output.seek(0)
    cursor.close()
    return output


def convert_to_cst_and_et(user_time, user_tz):
    """Convert user input time to CST and ET."""
    try:
        user_tz_obj = pytz.timezone(user_tz)
        cst_tz = pytz.timezone("US/Central")
        et_tz = pytz.timezone("US/Eastern")

        if isinstance(user_time, datetime):
            user_time_obj = user_time
        else:
            user_time_obj = datetime.strptime(user_time, "%Y-%m-%d %H:%M:%S")

        if user_time_obj.tzinfo is None:
            user_time_with_tz = user_tz_obj.localize(user_time_obj)
        else:
            user_time_with_tz = user_time_obj

        cst_time = user_time_with_tz.astimezone(cst_tz)
        et_time = user_time_with_tz.astimezone(et_tz)

        return (
            cst_time.strftime("%Y-%m-%d %H:%M:%S"),
            et_time.strftime("%Y-%m-%d %H:%M:%S")
        )
    except Exception as e:
        st.error(f"Error in time conversion: {e}")
        return None, None


# ------------------ Streamlit UI ------------------ #
st.set_page_config(page_title="Post-Test DB Query Executor", layout="centered")

st.markdown("<h1 style='text-align: center;'>🧠 Post-Test Oracle SQL Executor</h1>", unsafe_allow_html=True)
st.markdown("---")

# ---- Database Connection Section ----
st.markdown("<h4 style='text-align: center;'>🔗 Database Connection</h4>", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    host = st.text_input("Host", placeholder="e.g., XV005-SCAN.SDI.CORP.BANKOFAMERICA.COM")
    service_name = st.text_input("Service Name", placeholder="e.g., BRI")
    username = st.text_input("Username", placeholder="e.g., RPIST_WRITE")
with col2:
    port = st.text_input("Port", placeholder="e.g., 49125")
    password = st.text_input("Password", type="password", placeholder="Enter password")

connect_btn = st.button("🔌 Connect to Database")

connection_details = None
if connect_btn:
    success, message, db_tz = test_connection(host, port, service_name, username, password)
    if success:
        st.success(message)
        connection_details = (host, port, service_name, username, password)
    else:
        st.error(message)

st.markdown("---")

# ---- SQL Execution Section ----
st.markdown("<h4 style='text-align: center;'>📄 SQL Execution Panel</h4>", unsafe_allow_html=True)

execution_mode = st.radio("Select Execution Mode", ["Single Query", "Browse File"], horizontal=True)

available_timezones = pytz.all_timezones
user_timezone = st.selectbox("Select your timezone", available_timezones, index=available_timezones.index("UTC"))

start_time_input = date_time_picker(label="Select Start Time", key="start_time_picker")
end_time_input = date_time_picker(label="Select End Time", key="end_time_picker")

if start_time_input:
    start_cst, start_et = convert_to_cst_and_et(start_time_input, user_timezone)
    st.info(f"🕓 Start Time (CST): **{start_cst}** | (ET): **{start_et}**")

if end_time_input:
    end_cst, end_et = convert_to_cst_and_et(end_time_input, user_timezone)
    st.info(f"🕒 End Time (CST): **{end_cst}** | (ET): **{end_et}**")

if connection_details:
    host, port, service_name, username, password = connection_details
    dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"

    if execution_mode == "Single Query":
        st.subheader("🧾 Execute a Single Query")
        sql_query = st.text_area("Enter SQL Query", height=150)
        execute_btn = st.button("▶️ Execute Query")

        if execute_btn and sql_query.strip():
            try:
                conn = oracledb.connect(user=username, password=password, dsn=dsn)
                excel_data = execute_sql(conn, sql_query)
                conn.close()
                st.success("✅ Query executed successfully!")

                st.download_button(
                    label="⬇️ Download Excel File",
                    data=excel_data,
                    file_name="single_query_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error executing query: {e}")

    else:
        st.subheader("📁 Execute Queries from File")
        sql_file = st.file_uploader("Upload SQL File", type=["txt"])

        if sql_file:
            with open(f"uploaded_{sql_file.name}", "wb") as f:
                f.write(sql_file.getbuffer())

            with open(f"uploaded_{sql_file.name}", "r") as f:
                st.text_area("📜 File Contents", f.read(), height=200)

        execute_file_btn = st.button("▶️ Execute SQLs from File")

        if execute_file_btn and sql_file:
            try:
                conn = oracledb.connect(user=username, password=password, dsn=dsn)
                excel_data = execute_sqls_from_file(conn, f"uploaded_{sql_file.name}", start_cst, end_cst)
                conn.close()
                st.success("✅ SQLs executed successfully!")

                st.download_button(
                    label="⬇️ Download Excel File",
                    data=excel_data,
                    file_name="multi_query_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error executing SQLs: {e}")

st.markdown("---")
st.markdown("<p style='text-align:center;color:gray;'>© 2025 SQL Executor Tool</p>", unsafe_allow_html=True)
