import os
import re
import pytz
import smtplib
import sqlite3
import pandas as pd
import streamlit as st
import oracledb
from datetime import datetime
from openpyxl import Workbook
from sqlalchemy import create_engine
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from streamlit_datetime_picker import date_time_picker

# Initialize Oracle client (update path as required)
oracledb.init_oracle_client(lib_dir=r"C:\ORACLE19_X64\PRODUCT\19.3.0\client_1\bin")

# ------------------ Utility Functions ------------------ #
def replace_time(sql, start_time, end_time):
    """Replace placeholders in SQL with actual start and end times."""
    if start_time:
        sql = re.sub(r"&test_start_time", f"{start_time}", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"{end_time}", sql, flags=re.IGNORECASE)
    return sql


def send_email(recipient_email, file_path):
    """Send the Excel file as an email attachment."""
    try:
        sender_email = "your_email@example.com"
        sender_password = "your_password"
        subject = "SQL Execution Results"
        body = "Please find attached the SQL execution results."

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        with open(file_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(file_path)}")
        msg.attach(part)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)

        st.success("Email sent successfully!")

    except Exception as e:
        st.error(f"Error sending email: {e}")


def execute_sqls(sql_file_path, start_time, end_time, recipient_email):
    """Execute SQL queries and export results to Excel."""
    try:
        conn = oracledb.connect(
            user="RPIST_WRITE",
            password="U4EGbeTl#6FcxmNvZzV5ik9CA",
            dsn="(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=XV005-SCAN.SDI.CORP.BANKOFAMERICA.COM)(PORT=49125))(CONNECT_DATA=(SERVICE_NAME=BRI)))"
        )

        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM DUAL")
        st.success("Successfully connected to the database.")

        with open(sql_file_path, "r") as file:
            sql_statements = [line.strip().rstrip(';') for line in file if line.strip() and not line.strip().startswith("--")]

        writer = pd.ExcelWriter("output.xlsx", engine="openpyxl")

        for i, sql in enumerate(sql_statements, start=1):
            sql = replace_time(sql, start_time, end_time)
            cursor.execute(sql)
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

            df = pd.DataFrame(rows, columns=columns)
            df.to_excel(writer, sheet_name=f"Query_{i}", index=False)

        writer.close()
        conn.close()

        if recipient_email:
            send_email(recipient_email, "output.xlsx")

        st.success("SQLs executed successfully. Results saved to output.xlsx")

    except Exception as e:
        st.error(f"Error: {str(e)}")


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
st.title("SQL Executor")

sql_file = st.file_uploader("Upload SQL File", type=["txt"])

# Timezone selection
available_timezones = pytz.all_timezones
user_timezone = st.selectbox("Select your timezone", available_timezones, index=available_timezones.index("UTC"))

# Datetime pickers
start_time_input = date_time_picker(label="Select Start Time", key="start_time_picker")
end_time_input = date_time_picker(label="Select End Time", key="end_time_picker")

# Convert user input time
start_time = convert_to_cst_and_et(start_time_input, user_timezone) if start_time_input else None
end_time = convert_to_cst_and_et(end_time_input, user_timezone) if end_time_input else None

recipient_email = st.text_input("Recipient Email (optional)")

if st.button("Execute SQLs"):
    if sql_file is None:
        st.error("Please upload a .txt file.")
    else:
        sql_file_path = f"uploaded_{sql_file.name}"
        with open(sql_file_path, "wb") as f:
            f.write(sql_file.getbuffer())

        with open(sql_file_path, "r") as f:
            file_contents = f.read()

        st.text_area("File Contents", file_contents, height=200)

        execute_sqls(sql_file_path, start_time, end_time, recipient_email)
