# Run with: streamlit run streamlit_app.py

import streamlit as st
import pandas as pd
import re
import oracledb
from io import BytesIO
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

st.set_page_config(page_title="Post-Test DB Query Executor", layout="wide")
st.title("🏗️ Post-Test Automatic DB Query Execution")

# --- Section 1: Database Connection UI ---
st.header("🔌 Database Connection")

db_host = st.text_input("Host", placeholder="e.g., XVU05-SCAN.SDI.CORP.BANKOFAMERICA.COM")
db_port = st.text_input("Port", placeholder="e.g., 49125")
db_service = st.text_input("Service Name", placeholder="e.g., BRPIQ01_SVC01")
db_user = st.text_input("Username", placeholder="e.g., RPIST_WRITE")
db_pass = st.text_input("Password", type="password")

connect_button = st.button("Connect to Database")

# Global connection object
if "conn" not in st.session_state:
    st.session_state.conn = None

if connect_button:
    try:
        conn_str = f"{db_user}/{db_pass}@{db_host}:{db_port}/{db_service}"
        conn = oracledb.connect(conn_str, encoding="UTF-8", nencoding="UTF-8")
        st.session_state.conn = conn
        st.success("✅ Database connection established successfully!")
    except Exception as e:
        st.error(f"❌ Failed to connect: {e}")

# --- Proceed only if connected ---
if st.session_state.conn:
    st.header("📂 Upload and Execute SQL Queries")

    uploaded_file = st.file_uploader("Upload your query text file", type=["txt"])

    st.subheader("⏱️ Enter Time Range (if required)")
    start_time = st.text_input("Start Time (YYYY-MM-DD HH24:MI:SS)")
    end_time = st.text_input("End Time (YYYY-MM-DD HH24:MI:SS)")

    st.subheader("📧 Email Report (Optional)")
    send_email = st.checkbox("Send results via internal Outlook email")
    sender_email = st.text_input("Your Outlook Email ID (e.g., your.name@corp.bankofamerica.com)")
    recipient_email = st.text_input("Recipient Email (comma separated)")

    run_button = st.button("▶️ Run Queries & Generate Excel")

    # --- Helper Functions ---
    def parse_queries_from_text(text: str):
        """Handles queries starting with 'Query1' or '-- Query1'"""
        queries = {}
        parts = re.split(r"(?i)--\s*Query\d+|(?i)Query\d+", text)
        headers = re.findall(r"(?i)--\s*Query\d+|(?i)Query\d+", text)

        for i, query_text in enumerate(parts[1:], start=1):
            name = headers[i - 1].strip().replace("--", "").strip()
            queries[name] = query_text.strip()
        return queries

    def replace_time_variables(sql, start_time, end_time):
        if start_time:
            sql = re.sub(r"&test_start_time", f"'{start_time}'", sql, flags=re.IGNORECASE)
        if end_time:
            sql = re.sub(r"&test_end_time", f"'{end_time}'", sql, flags=re.IGNORECASE)
        return sql

    def safe_sheet_name(name):
        return re.sub(r"[\\/*?:\[\]]", "_", name)[:31]

    def execute_queries(queries, start_time=None, end_time=None):
        conn = st.session_state.conn
        cur = conn.cursor()
        excel_io = BytesIO()
        writer = pd.ExcelWriter(excel_io, engine="xlsxwriter")

        for name, query in queries.items():
            sql = replace_time_variables(query.strip().rstrip(";"), start_time, end_time)
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

    def send_email_internal_outlook(sender_email, recipient_emails, attachment_bytes, filename):
        """Send email using internal Outlook relay (no authentication)."""
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = ", ".join(recipient_emails)
        msg["Subject"] = "Post-Test DB Query Results"

        body = "Hi,\n\nPlease find attached the results of post-test database queries.\n\nRegards,\nAutomated System"
        msg.attach(MIMEText(body, "plain"))

        part = MIMEApplication(attachment_bytes.getvalue(), _subtype="xlsx")
        part.add_header("Content-Disposition", "attachment", filename=filename)
        msg.attach(part)

        internal_relays = [
            "mail.corp.bankofamerica.com",
            "smtp.corp.bankofamerica.com",
            "relay.corp.bankofamerica.com",
            "localhost",
        ]

        for relay in internal_relays:
            try:
                with smtplib.SMTP(relay, 25, timeout=10) as server:
                    server.send_message(msg)
                st.success(f"📧 Email sent successfully using relay: {relay}")
                return True
            except Exception as e:
                st.warning(f"Tried {relay} — failed: {e}")

        st.error("❌ Could not connect to any internal mail relay. Contact IT for the correct SMTP host.")
        return False

    # --- Execution ---
    if run_button:
        if not uploaded_file:
            st.error("Please upload a query text file.")
        else:
            text_data = uploaded_file.read().decode("utf-8", errors="ignore")
            queries = parse_queries_from_text(text_data)
            if not queries:
                st.error("No queries found in the uploaded file.")
            else:
                st.info(f"Executing {len(queries)} queries...")
                try:
                    excel_data = execute_queries(queries, start_time, end_time)
                    filename = f"posttest_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    st.success("✅ Queries executed successfully!")
                    st.download_button("📥 Download Excel", data=excel_data, file_name=filename)

                    if send_email and sender_email and recipient_email:
                        recipients = [r.strip() for r in recipient_email.split(",") if r.strip()]
                        send_email_internal_outlook(sender_email, recipients, excel_data, filename)

                except Exception as e:
                    st.error(f"Error executing queries: {e}")
