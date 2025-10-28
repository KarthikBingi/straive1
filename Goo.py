import streamlit as st
import oracledb
import pandas as pd
from io import BytesIO
from datetime import datetime
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Post-Test DB Execution", layout="wide")
st.title("📊 Post-Test Automatic DB Query Execution")

# -----------------------------
# SECTION 1: DATABASE CONNECTION
# -----------------------------
st.header("🔌 Database Connection")

db_host = st.text_input("Host", placeholder="e.g., XVU05-SCAN.SDI.CORP.BANKOFAMERICA.COM")
db_port = st.text_input("Port", placeholder="e.g., 49125")
db_service = st.text_input("Service Name", placeholder="e.g., BRPIQ01_SVC01")
db_user = st.text_input("Username", placeholder="e.g., RPIST_WRITE")
db_pass = st.text_input("Password", type="password")

connect_button = st.button("Connect to Database")

# Session variable for DB connection
if "conn" not in st.session_state:
    st.session_state.conn = None

if connect_button:
    try:
        conn_str = f"{db_user}/{db_pass}@{db_host}:{db_port}/{db_service}"
        conn = oracledb.connect(conn_str, encoding="UTF-8", nencoding="UTF-8")
        st.session_state.conn = conn
        st.success("✅ Database connection established successfully!")

        # Fetch timezone info after connection
        cur = conn.cursor()
        cur.execute("SELECT dbtimezone, sessiontimezone, CURRENT_TIMESTAMP FROM dual")
        row = cur.fetchone()
        if row:
            db_tz, session_tz, current_ts = row
            st.info(f"**Database Timezone:** {db_tz}")
            st.info(f"**Session Timezone:** {session_tz}")
            st.info(f"**Current DB Time:** {current_ts}")
        else:
            st.warning("Unable to fetch timezone information.")
        cur.close()

    except Exception as e:
        st.error(f"❌ Failed to connect: {e}")

# -----------------------------
# SECTION 2: SQL QUERY EXECUTION
# -----------------------------
st.header("📜 Execute SQL Queries")

uploaded_file = st.file_uploader("Upload a text file containing SQL queries", type=["txt"])
execute_button = st.button("Run Queries and Export to Excel")

if execute_button:
    if st.session_state.conn is None:
        st.error("⚠️ Please connect to the database first.")
    elif uploaded_file is None:
        st.error("⚠️ Please upload a .txt file containing SQL queries.")
    else:
        try:
            queries_text = uploaded_file.read().decode("utf-8")
            queries = [q.strip() for q in queries_text.split(";") if q.strip()]
            conn = st.session_state.conn
            cursor = conn.cursor()

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for i, query in enumerate(queries, start=1):
                    try:
                        df = pd.read_sql(query, conn)
                        sheet_name = f"Query{i}"
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        st.success(f"✅ Query{i} executed successfully — added to sheet '{sheet_name}'")
                    except Exception as qe:
                        st.warning(f"⚠️ Query{i} failed: {qe}")

            output.seek(0)
            st.download_button(
                label="⬇️ Download Excel File",
                data=output,
                file_name=f"DB_Query_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error executing queries: {e}")

# -----------------------------
# SECTION 3: EMAIL RESULTS
# -----------------------------
st.header("📧 Send Excel Results via Outlook Mail (Internal Only)")

receiver_email = st.text_input("Enter recipient email (within organisation)", placeholder="e.g., colleague@yourcompany.com")
subject = st.text_input("Email Subject", "Database Query Results")
body = st.text_area("Email Body", "Please find the attached query results.")

send_email_button = st.button("Send Email")

if send_email_button:
    if not receiver_email:
        st.error("⚠️ Please enter the recipient email.")
    else:
        try:
            # Simulate sending via internal mail relay
            smtp_server = "mail.yourcompany.com"  # replace with internal mail relay host
            smtp_port = 25  # standard internal port

            msg = EmailMessage()
            msg["From"] = f"{db_user}@yourcompany.com"
            msg["To"] = receiver_email
            msg["Subject"] = subject
            msg.set_content(body)

            # Attach last generated Excel if available
            if 'output' in locals():
                msg.add_attachment(output.getvalue(),
                                   maintype="application",
                                   subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   filename="Query_Results.xlsx")

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.send_message(msg)

            st.success(f"✅ Email successfully sent to {receiver_email}")
        except Exception as e:
            st.error(f"❌ Failed to send email: {e}")
