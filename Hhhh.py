import streamlit as st
import pandas as pd
import re
import oracledb
import pytz
from datetime import datetime
from streamlit_datetime_picker import date_time_picker
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# --------------------------- Streamlit Page Config --------------------------- #
st.set_page_config(page_title="DB Query Executor", layout="centered")
st.title("🏦 Post-Test DB Query Executor")

# --------------------------- DB Connection Section --------------------------- #
st.header("🔗 Database Connection")

host = st.text_input("Host")
port = st.text_input("Port")
service_name = st.text_input("Service Name")
username = st.text_input("Username")
password = st.text_input("Password", type="password")

if st.button("Connect to Database"):
    try:
        dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
        conn = oracledb.connect(user=username, password=password, dsn=dsn)
        st.success("✅ Connected to the database successfully!")

        cursor = conn.cursor()
        cursor.execute("SELECT SESSIONTIMEZONE FROM DUAL")
        db_timezone = cursor.fetchone()[0]

        timezone_map = {
            '-04:00': 'Eastern Time (ET)',
            '-05:00': 'Central Time (CT)',
            '+05:30': 'India Standard Time (IST)',
            '00:00': 'Greenwich Mean Time (GMT)'
        }
        readable_tz = timezone_map.get(db_timezone, f"Unknown ({db_timezone})")
        st.info(f"🕐 Database Time Zone: **{readable_tz}**")

        st.session_state["db_connection"] = conn
        st.session_state["db_password"] = password
    except Exception as e:
        st.error(f"❌ Failed to connect: {e}")

# --------------------------- Helper Functions --------------------------- #
def replace_time(sql, start_time, end_time):
    if start_time:
        sql = re.sub(r"&test_start_time", f"{start_time}", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"{end_time}", sql, flags=re.IGNORECASE)
    return sql

def send_email(recipient, file_path, sender_password):
    sender_email = "bingi.karthik@bofa.com"
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient
    msg["Subject"] = "Automated SQL Report"

    body = "Please find the attached SQL execution report."
    msg.attach(MIMEText(body, "plain"))

    with open(file_path, "rb") as file:
        part = MIMEApplication(file.read(), Name="output.xlsx")
    part["Content-Disposition"] = 'attachment; filename="output.xlsx"'
    msg.attach(part)

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        st.success(f"📧 Report successfully sent to {recipient}")
    except Exception as e:
        st.error(f"❌ Email send failed: {e}")

# --------------------------- SQL Execution Section --------------------------- #
if "db_connection" in st.session_state:
    conn = st.session_state["db_connection"]
    db_password = st.session_state["db_password"]

    st.header("⚙️ SQL Execution Panel")

    mode = st.radio("Select Mode", ["Single Query", "Browse File"])
    recipient_email = st.text_input("Recipient Email (optional)")
    share_email = st.checkbox("📤 Share report via email")

    # ---------------------- SINGLE QUERY MODE ---------------------- #
    if mode == "Single Query":
        query = st.text_area("Enter your SQL Query here")

        if st.button("Run Query"):
            try:
                cursor = conn.cursor()
                query = query.strip().rstrip(';')  # Remove trailing semicolon
                cursor.execute(query)
                columns = [desc[0] for desc in cursor.description]
                rows = cursor.fetchall()

                df = pd.DataFrame(rows, columns=columns)
                file_path = "single_query_output.xlsx"
                df.to_excel(file_path, index=False)

                st.success("✅ Query executed successfully.")
                st.download_button("⬇️ Download Excel", data=open(file_path, "rb"),
                                   file_name="single_query_output.xlsx")

                if share_email and recipient_email:
                    send_email(recipient_email, file_path, db_password)

            except Exception as e:
                st.error(f"Execution failed: {e}")

    # ---------------------- BROWSE FILE MODE ---------------------- #
    elif mode == "Browse File":
        sql_file = st.file_uploader("Upload SQL File", type=["txt"])
        available_timezones = pytz.all_timezones
        user_timezone = st.selectbox("Select your Timezone", available_timezones,
                                     index=available_timezones.index("UTC"))

        start_time_input = date_time_picker(label="Select Start Time", key="start_time_picker")
        end_time_input = date_time_picker(label="Select End Time", key="end_time_picker")

        def convert_to_et(user_time, user_tz):
            try:
                user_tz_obj = pytz.timezone(user_tz)
                et_tz = pytz.timezone("US/Eastern")
                if isinstance(user_time, datetime):
                    user_time_obj = user_time
                else:
                    user_time_obj = datetime.strptime(user_time, "%Y-%m-%d %H:%M:%S")

                if user_time_obj.tzinfo is None:
                    user_time_with_tz = user_tz_obj.localize(user_time_obj)
                else:
                    user_time_with_tz = user_time_obj

                et_time = user_time_with_tz.astimezone(et_tz)
                st.write(f"🕓 Converted Time (ET): {et_time.strftime('%Y-%m-%d %H:%M:%S')}")
                return et_time.strftime("%Y-%m-%d %H:%M:%S")
            except Exception as e:
                st.error(f"Error converting time: {e}")
                return None

        start_time = convert_to_et(start_time_input, user_timezone) if start_time_input else None
        end_time = convert_to_et(end_time_input, user_timezone) if end_time_input else None

        if st.button("Execute SQLs"):
            if not sql_file:
                st.error("Please upload a .txt file.")
            else:
                try:
                    lines = sql_file.read().decode("utf-8").splitlines()
                    sqls = [line.strip().rstrip(';') for line in lines if line.strip() and not line.strip().startswith("--")]

                    writer = pd.ExcelWriter("output.xlsx", engine="openpyxl")

                    for i, sql in enumerate(sqls):
                        sql = replace_time(sql, start_time, end_time)
                        sql = sql.strip().rstrip(';')  # Ensure no semicolon
                        cursor = conn.cursor()
                        cursor.execute(sql)
                        columns = [col[0] for col in cursor.description]
                        rows = cursor.fetchall()
                        df = pd.DataFrame(rows, columns=columns)
                        df.to_excel(writer, sheet_name=f"Query_{i+1}", index=False)

                    writer.close()
                    st.success("✅ SQLs executed successfully.")
                    st.download_button("⬇️ Download Excel", data=open("output.xlsx", "rb"),
                                       file_name="output.xlsx")

                    if share_email and recipient_email:
                        send_email(recipient_email, "output.xlsx", db_password)

                except Exception as e:
                    st.error(f"Execution failed: {e}")
