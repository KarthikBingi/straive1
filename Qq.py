# streamlit_app.py
# Run using: streamlit run streamlit_app.py

import streamlit as st
import pandas as pd
import re
import oracledb
from io import BytesIO
from datetime import datetime, timedelta, timezone
import pytz
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

st.set_page_config(page_title="Post-Test DB Query Executor", layout="wide")
st.title("🏦 Post-Test Automatic DB Query Execution")

# --- UI for DB Details ---
st.subheader("🔑 Enter Database Details")
db_host = st.text_input("Host", "XVU05-SCAN.SDI.CORP.BANKOFAMERICA.COM")
db_port = st.text_input("Port", "49125")
db_service = st.text_input("Service", "BRPIQ01_SVC01")
db_user = st.text_input("Username", "RPIST_WRITE")
db_pass = st.text_input("Password", type="password")

connect_btn = st.button("🔌 Connect to Database")

# Global variables
conn = None
db_timezone_offset = None

# --- Connect to DB ---
if connect_btn:
    try:
        conn = oracledb.connect(
            user=db_user,
            password=db_pass,
            dsn=f"{db_host}:{db_port}/{db_service}"
        )
        st.success("✅ Successfully connected to the database!")

        # Fetch DB timezone
        cur = conn.cursor()
        cur.execute("SELECT DBTIMEZONE FROM DUAL")
        db_timezone = cur.fetchone()[0]
        st.info(f"🕒 Database Timezone: {db_timezone}")

        # Convert DB timezone (e.g., -04:00) to timezone object
        sign = 1 if db_timezone.startswith("+") else -1
        hours, mins = map(int, db_timezone[1:].split(":"))
        db_timezone_offset = timezone(timedelta(hours=sign * hours, minutes=sign * mins))

        # Save connection info
        st.session_state["db_connection"] = conn
        st.session_state["db_timezone"] = db_timezone_offset
        st.session_state["db_info"] = {
            "host": db_host, "port": db_port, "service": db_service,
            "user": db_user, "pass": db_pass
        }

    except Exception as e:
        st.error(f"❌ Failed to connect: {e}")

# --- Query File Upload ---
st.subheader("📂 Upload Query File")
uploaded_file = st.file_uploader("Upload your query text file", type=["txt"])

# --- Time Input Section ---
st.subheader("⏱ Enter Time Range (IST)")
start_time = st.text_input("Start Time (YYYY-MM-DD HH:MM:SS)")
end_time = st.text_input("End Time (YYYY-MM-DD HH:MM:SS)")

# --- Email Section ---
st.subheader("📧 Send Results via Outlook (Internal Only)")
send_email = st.checkbox("Send results through internal Outlook mail")
email_sender = None
email_recipient = None
if send_email:
    email_sender = st.text_input("Sender Email (e.g. you@corp.bankofamerica.com)")
    email_recipient = st.text_input("Recipient Email(s) (comma separated, internal only)")

run_button = st.button("🚀 Run Queries & Generate Excel")

# --- Helper Functions ---
def parse_queries_from_text(text: str):
    queries = {}
    parts = re.split(r"(?i)--\s*(Query\d+)", text)
    current_name = None
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if re.match(r"(?i)^Query\d+$", part):
            current_name = part.strip()
            queries[current_name] = ""
        elif current_name:
            queries[current_name] += " " + part
    return queries

def replace_time_variables(sql, start_time, end_time):
    if start_time:
        sql = re.sub(r"&test_start_time", f"'{start_time}'", sql, flags=re.IGNORECASE)
    if end_time:
        sql = re.sub(r"&test_end_time", f"'{end_time}'", sql, flags=re.IGNORECASE)
    return sql

def convert_ist_to_db_time(ist_time_str, db_offset):
    """Convert IST time entered by user to DB timezone."""
    ist = pytz.timezone("Asia/Kolkata")
    db_tz = timezone(db_offset.utcoffset(None))
    local_time = ist.localize(datetime.strptime(ist_time_str, "%Y-%m-%d %H:%M:%S"))
    db_time = local_time.astimezone(db_tz)
    return db_time.strftime("%Y-%m-%d %H:%M:%S")

def safe_sheet_name(name):
    return re.sub(r"[\\/*?:\[\]]", "_", name)[:31]

def execute_queries(queries, db_info, db_offset, start_time=None, end_time=None):
    conn = None
    excel_io = BytesIO()
    try:
        conn = oracledb.connect(
            user=db_info["user"],
            password=db_info["pass"],
            dsn=f"{db_info['host']}:{db_info['port']}/{db_info['service']}"
        )
        cur = conn.cursor()
        writer = pd.ExcelWriter(excel_io, engine="xlsxwriter")

        db_start_time = convert_ist_to_db_time(start_time, db_offset) if start_time else None
        db_end_time = convert_ist_to_db_time(end_time, db_offset) if end_time else None

        st.info(f"📆 DB Time Range: {db_start_time} → {db_end_time}")

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
    finally:
        if conn:
            conn.close()

def send_email_internal_outlook(sender_email, recipient_emails, attachment_bytes, filename):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipient_emails)
    msg["Subject"] = "Post-Test DB Query Results"
    body = "Hi,\n\nPlease find attached the results of post-test database queries.\n\nRegards,\nAutomated System"
    msg.attach(MIMEText(body, "plain"))

    part = MIMEApplication(attachment_bytes.getvalue(), _subtype="xlsx")
    part.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(part)

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"❌ Failed to send email: {e}")
        return False

# --- Main Run ---
if run_button:
    if "db_connection" not in st.session_state:
        st.error("❌ Please connect to the database first.")
    elif not uploaded_file:
        st.error("Please upload a query text file.")
    else:
        text_data = uploaded_file.read().decode("utf-8", errors="ignore")
        queries = parse_queries_from_text(text_data)
        if not queries:
            st.error("No queries found in the uploaded file.")
        else:
            try:
                excel_data = execute_queries(
                    queries,
                    st.session_state["db_info"],
                    st.session_state["db_timezone"],
                    start_time,
                    end_time
                )
                filename = f"posttest_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.success("✅ Queries executed successfully!")
                st.download_button("📥 Download Excel", data=excel_data, file_name=filename)

                if send_email:
                    if not email_sender or not email_recipient:
                        st.error("Please provide sender and recipient emails.")
                    else:
                        recipients = [r.strip() for r in email_recipient.split(",") if r.strip()]
                        if send_email_internal_outlook(email_sender, recipients, excel_data, filename):
                            st.success("📧 Email sent successfully (internal Outlook)!")

            except Exception as e:
                st.error(f"Error executing queries: {e}")
