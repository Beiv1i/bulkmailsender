import streamlit as st
import pandas as pd
import os
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from datetime import datetime
from io import BytesIO
import config
import re

# --- Page Configuration ---
st.set_page_config(
    page_title="Mail Drop",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Swiss Design System (CSS) ---
st.markdown("""
<style>
    /* Font Import - Inter */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

    /* Global Reset & Typography */
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #1a1a1a;
        font-weight: 400;
    }
    
    /* Backgrounds */
    .stApp {
        background-color: #ffffff;
    }
    
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
        border-right: 1px solid #eaeaea;
    }

    /* Headings */
    h1, h2, h3 {
        font-weight: 600 !important;
        letter-spacing: -0.02em !important;
        color: #000000 !important;
    }
    h1 { font-size: 2.2rem !important; margin-bottom: 1.5rem !important; }
    h2 { font-size: 1.2rem !important; margin-top: 2rem !important; margin-bottom: 1rem !important; }
    h3 { font-size: 1.0rem !important; font-weight: 500 !important; opacity: 0.8; }

    /* Inputs & Textareas */
    .stTextInput input, .stTextArea textarea, .stNumberInput input {
        background-color: #ffffff !important;
        border: 1px solid #e0e0e0 !important;
        border-radius: 6px !important;
        color: #1a1a1a !important;
        padding: 0.8rem !important;
        font-size: 0.95rem !important;
        transition: border-color 0.2s ease;
    }
    .stTextInput input:focus, .stTextArea textarea:focus, .stNumberInput input:focus {
        border-color: #000000 !important;
        box-shadow: none !important;
    }

    /* File Uploader */
    [data-testid="stFileUploader"] {
        border: 1px dashed #e0e0e0;
        border-radius: 8px;
        padding: 2rem;
        text-align: center;
        background-color: #fafafa;
        transition: background-color 0.2s;
    }
    [data-testid="stFileUploader"]:hover {
        background-color: #f0f0f0;
    }

    /* Buttons */
    .stButton button {
        border-radius: 6px !important;
        font-weight: 500 !important;
        padding: 0.6rem 1.2rem !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    
    /* Primary Action (Send) */
    button[kind="primary"] {
        background-color: #000000 !important;
        color: #ffffff !important;
    }
    button[kind="primary"]:hover {
        background-color: #333333 !important;
        transform: translateY(-1px);
    }
    
    /* Secondary Action (Save/Download) */
    button[kind="secondary"] {
        background-color: #f0f0f0 !important;
        color: #000000 !important;
        border: 1px solid #e0e0e0 !important;
    }
    button[kind="secondary"]:hover {
        border-color: #000000 !important;
        background-color: #ffffff !important;
    }

    /* Dividers */
    hr {
        margin: 2rem 0 !important;
        border-color: #eaeaea !important;
    }

    /* Metrics & Cards */
    .css-1r6slb0 {
        border: 1px solid #eaeaea;
        padding: 1.5rem;
        border-radius: 8px;
    }
    
    /* Hide Default Streamlit Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Custom Padding Fixes */
    .block-container {
        padding-top: 3rem !important;
        padding-bottom: 5rem !important;
        max-width: 1200px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---
def smart_str(val):
    if pd.isna(val): return ""
    if isinstance(val, float):
        if val.is_integer(): return str(int(val))
    return str(val).strip()

def send_one_email(row, template_content, placeholders, subject, s_name, s_email):
    try:
        msg_body = template_content
        for key in placeholders:
            val = row.get(key)
            msg_body = msg_body.replace(f"{{{key}}}", smart_str(val))
            
        msg = MIMEMultipart()
        msg['From'] = formataddr((s_name, s_email))
        
        recipient = row.get('邮箱') or row.get('Email') or row.get('email')
        if not recipient or pd.isna(recipient):
            return False, "Missing email address", None
            
        msg['To'] = str(recipient).strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(msg_body, 'plain', 'utf-8'))
        
        return True, "Ready", msg
    except Exception as e:
        return False, str(e), None

# --- Sidebar ---
with st.sidebar:
    st.markdown("### Configuration")
    
    # Credentials
    st.markdown("#### Credentials")
    sender_name = st.text_input("Sender Name", value=config.SENDER_NAME, placeholder="e.g. John Doe")
    sender_email = st.text_input("Sender Email", value=config.SENDER_EMAIL, placeholder="name@company.com")
    sender_password = st.text_input("App Password", value=config.APP_PASSWORD, type="password")
    
    st.markdown("---")
    
    # Settings
    st.markdown("#### Delivery Settings")
    default_limit = getattr(config, 'BATCH_LIMIT', 0)
    batch_limit = st.number_input("Batch Limit (0 for infinite)", min_value=0, value=default_limit)
    
    st.markdown("#### Humanization")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        sleep_min = st.number_input("Min Delay (s)", 1.0, 60.0, 2.0)
    with col_s2:
        sleep_max = st.number_input("Max Delay (s)", sleep_min, 60.0, 5.0)

# --- Main Interface ---
st.title("Mail Drop")
st.markdown("<p style='font-size: 1.1rem; color: #666; margin-bottom: 2rem;'>Secure bulk email dispatch system.</p>", unsafe_allow_html=True)

col1, col_spacer, col2 = st.columns([1, 0.1, 1])

with col1:
    st.markdown("## 01. Audience")
    uploaded_file = st.file_uploader("Drop your Excel recipient list here", type=["xlsx"])
    
    df = None
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            if df.empty:
                st.toast("File is empty", icon="⚠️")
            else:
                st.dataframe(df.head(5), height=200, use_container_width=True)
                st.markdown(f"<p style='font-size: 0.9rem; color: #666; margin-top: 0.5rem;'>✓ Loaded {len(df)} recipients</p>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error reading file: {e}")

with col2:
    st.markdown("## 02. Content")
    
    # Template Loader
    try:
        with open("template.txt", "r") as f: default_template = f.read()
    except:
        default_template = "Hello {Name},\n\n..."

    default_subject = getattr(config, 'EMAIL_SUBJECT', "Update")
    email_subject = st.text_input("Subject Line", value=default_subject)
    template_content = st.text_area("Message Body", value=default_template, height=250)
    
    if st.button("Save Template", type="secondary"):
        with open("template.txt", "w") as f:
            f.write(template_content)
        st.toast("Template saved successfully")

# --- Action Area ---
if df is not None and not df.empty:
    st.markdown("---")
    st.markdown("## 03. Review & Dispatch")
    
    # Variable Extraction
    placeholders = set(re.findall(r'\{(.*?)\}', template_content))
    missing_cols = [p for p in placeholders if p not in df.columns]
    
    if missing_cols:
        st.error(f"Missing columns in Excel: {', '.join(missing_cols)}")
    else:
        # Preview Card
        with st.container():
            st.markdown("#### Preview")
            preview_row = df.iloc[0]
            preview_body = template_content
            for key in placeholders:
                preview_body = preview_body.replace(f"{{{key}}}", smart_str(preview_row.get(key)))
            
            preview_html = f"""
            <div style="background-color: #fafafa; border: 1px solid #eaeaea; padding: 1.5rem; border-radius: 8px; font-family: 'Inter', sans-serif;">
                <div style="margin-bottom: 0.5rem; font-size: 0.9rem; color: #666;">
                    <strong>To:</strong> {preview_row.get('邮箱') or preview_row.get('Email')}<br>
                    <strong>Subject:</strong> {email_subject}
                </div>
                <hr style="margin: 1rem 0; border-color: #eaeaea;">
                <div style="white-space: pre-wrap; font-size: 0.95rem; line-height: 1.5;">{preview_body}</div>
            </div>
            """
            st.markdown(preview_html, unsafe_allow_html=True)

        st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)
        
        # Send Button
        if st.button("Initialize Dispatch Sequence", type="primary", use_container_width=True):
            if not sender_email or not sender_password:
                st.error("Please configure sender credentials in the sidebar.")
                st.stop()
                
            # Batch Logic
            if batch_limit > 0 and len(df) > batch_limit:
                task_df = df.iloc[:batch_limit].copy()
                remaining_df = df.iloc[batch_limit:].copy()
            else:
                task_df = df.copy()
                remaining_df = pd.DataFrame()
                
            # Progress UI
            progress_bar = st.progress(0)
            status_text = st.empty()
            processed_records = []
            
            # SMTP Connection
            try:
                with st.spinner(f"Authenticating as {sender_email}..."):
                    if config.SMTP_PORT == 465:
                        server = smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT)
                    else:
                        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
                        server.starttls()
                    server.login(sender_email, sender_password)
            except Exception as e:
                st.error(f"Connection Failed: {e}")
                st.stop()
                
            # Sending Loop
            total = len(task_df)
            success_count = 0
            
            for i, (index, row) in enumerate(task_df.iterrows()):
                name = row.get('账号', row.get('姓名', 'Recipient'))
                status_text.markdown(f"<span style='color: #666; font-size: 0.9rem;'>Dispatching {i+1}/{total}: <strong>{name}</strong></span>", unsafe_allow_html=True)
                
                is_ready, msg_str, msg_obj = send_one_email(row, template_content, placeholders, email_subject, sender_name, sender_email)
                
                if is_ready:
                    try:
                        server.sendmail(sender_email, msg_obj['To'], msg_obj.as_string())
                        status, detail = "Success", "OK"
                        success_count += 1
                    except Exception as e:
                        status, detail = "Failed", str(e)
                else:
                    status, detail = "Failed", msg_str
                    
                record = row.to_dict()
                record.update({'Status': status, 'Details': detail, 'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                processed_records.append(record)
                
                progress_bar.progress((i + 1) / total)
                
                if i < total - 1:
                    import random
                    time.sleep(random.uniform(sleep_min, sleep_max))
            
            server.quit()
            status_text.empty()
            
            # Post-Process
            if processed_records:
                # History File
                history_file = "sent_history.xlsx"
                new_recs = pd.DataFrame(processed_records)
                try:
                    if os.path.exists(history_file):
                        pd.concat([pd.read_excel(history_file), new_recs]).to_excel(history_file, index=False)
                    else:
                        new_recs.to_excel(history_file, index=False)
                except: pass
                
                # Success UI
                st.success(f"Operation Complete. {success_count} sent, {total-success_count} failed.")
                
                # Download Report
                output_log = BytesIO()
                new_recs.to_excel(output_log, index=False)
                
                col_d1, col_d2 = st.columns([1, 1])
                with col_d1:
                    st.download_button(
                        label="Download Report",
                        data=output_log.getvalue(),
                        file_name=f"Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                if not remaining_df.empty:
                    with col_d2:
                        rem_out = BytesIO()
                        remaining_df.to_excel(rem_out, index=False)
                        st.download_button(
                            label=f"Download Remaining ({len(remaining_df)})",
                            data=rem_out.getvalue(),
                            file_name=f"Remaining_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
