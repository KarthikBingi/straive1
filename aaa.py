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

# Initialize Oracle client (update your Oracle client path)
oracledb.init_oracle_client(lib_dir=r"C:\ORACLE19_X64\PRODUCT\19.3.0\client_1\bin")

# ------------------ Utility Functions ------------------ #
def replace_time(sql, start_time, end_time):
    """Replace placeholders in SQL with actual start and end times."""
    if start_time:
        sql = re.sub(r"&test_start_time", f"{start_time}", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"{end_time}", sql, flags=re.IGNORECASE)
    return sql


def test_connection(host, port, service_name, username, password):
    """Test Oracle DB connection."""
    try:
        dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
        conn = oracledb.connect(user=username, password=password, dsn=dsn)
        conn.close()
        return True, "Successfully connected to the database!"
    except Exception as e:
        return False, str(e)


def execute_sqls(conn, sql_file_path, start_time, end_time):
    """Execute SQL queries and return Excel bytes."""
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
    """Convert user input time to CST timezone."""
    try:
        user_tz_obj = pytz.timezone(user_tz)
        cst_tz = pytz.timezone("US/Central")

        if isinstance(user_time, datetime):
            user_time_obj = user_time
        else:
            user_time_obj = datetime.strptime(user_time, "%Y-%m-%d %H:%M:%S")

        if user_time_obj.tzinfo is None:
            user_time_with_tz = user_tz_obj.localize(user_time_obj)
        else:
            user_time_with_tz = user_time_obj

        cst_time = user_time_with_tz.astimezone(cst_tz)
        return cst_time.strftime("%Y-%m-%d %H:%M:%S")

    except Exception as e:
        st.error(f"Error in time conversion: {e}")
        return None


# ------------------ Streamlit UI ------------------ #
st.set_page_config(page_title="Post-Test DB Query Executor", layout="wide")
st.title("🔍 Oracle SQL Executor Tool")

st.subheader("🗄️ Database Connection Details")

host = st.text_input("Host", placeholder="e.g., XV005-SCAN.SDI.CORP.BANKOFAMERICA.COM")
port = st.text_input("Port", placeholder="e.g., 49125")
service_name = st.text_input("Service Name", placeholder="e.g., BRI")
username = st.text_input("Username", placeholder="e.g., RPIST_WRITE")
password = st.text_input("Password", type="password", placeholder="Enter password")

connect_btn = st.button("🔗 Connect")

connection = None
if connect_btn:
    success, message = test_connection(host, port, service_name, username, password)
    if success:
        st.success(message)
        connection = True
    else:
        st.error(f"Connection failed: {message}")

st.divider()

st.subheader("📄 SQL File Execution")

sql_file = st.file_uploader("Upload SQL File", type=["txt"])

available_timezones = pytz.all_timezones
user_timezone = st.selectbox("Select your timezone", available_timezones, index=available_timezones.index("UTC"))

start_time_input = date_time_picker(label="Select Start Time", key="start_time_picker")
end_time_input = date_time_picker(label="Select End Time", key="end_time_picker")

start_time = convert_to_cst_and_et(start_time_input, user_timezone) if start_time_input else None
end_time = convert_to_cst_and_et(end_time_input, user_timezone) if end_time_input else None

execute_btn = st.button("▶️ Execute SQLs")

if execute_btn:
    if not sql_file:
        st.error("Please upload a .txt SQL file.")
    elif not (host and port and service_name and username and password):
        st.error("Please fill all database connection details.")
    else:
        try:
            dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
            conn = oracledb.connect(user=username, password=password, dsn=dsn)

            sql_file_path = f"uploaded_{sql_file.name}"
            with open(sql_file_path, "wb") as f:
                f.write(sql_file.getbuffer())

            with open(sql_file_path, "r") as f:
                file_contents = f.read()

            st.text_area("📜 File Contents", file_contents, height=200)

            excel_data = execute_sqls(conn, sql_file_path, start_time, end_time)
            conn.close()

            st.success("✅ SQLs executed successfully!")

            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error executing SQLs: {e}")
