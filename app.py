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

# --- 页面配置 ---
st.set_page_config(
    page_title="智能投递",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 瑞士风格设计系统 (CSS) ---
st.markdown("""
<style>
    /* 字体导入 - Inter (虽然主要显示中文，但数字和英文仍需好字体) */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

    /* 全局重置与排版 */
    html, body, [class*="css"] {
        font-family: -apple-system, "PingFang SC", "Microsoft YaHei", 'Inter', sans-serif !important;
        color: #1a1a1a;
        font-weight: 400;
    }
    
    /* 背景色 */
    .stApp {
        background-color: #ffffff;
    }
    
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
        border-right: 1px solid #eaeaea;
    }

    /* 标题 */
    h1, h2, h3 {
        font-weight: 600 !important;
        letter-spacing: -0.01em !important;
        color: #000000 !important;
    }
    h1 { font-size: 2.2rem !important; margin-bottom: 1.5rem !important; }
    h2 { font-size: 1.2rem !important; margin-top: 2rem !important; margin-bottom: 1rem !important; }
    h3 { font-size: 1.0rem !important; font-weight: 500 !important; opacity: 0.8; }

    /* 输入框与文本域 */
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

    /* 文件上传区 */
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

    /* 按钮样式 */
    .stButton button {
        border-radius: 6px !important;
        font-weight: 500 !important;
        padding: 0.6rem 1.2rem !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    
    /* 主操作按钮 (发送) */
    button[kind="primary"] {
        background-color: #000000 !important;
        color: #ffffff !important;
    }
    button[kind="primary"]:hover {
        background-color: #333333 !important;
        transform: translateY(-1px);
    }
    
    /* 次级操作按钮 (保存/下载) */
    button[kind="secondary"] {
        background-color: #f0f0f0 !important;
        color: #000000 !important;
        border: 1px solid #e0e0e0 !important;
    }
    button[kind="secondary"]:hover {
        border-color: #000000 !important;
        background-color: #ffffff !important;
    }

    /* 分割线 */
    hr {
        margin: 2rem 0 !important;
        border-color: #eaeaea !important;
    }

    /* 隐藏 Streamlit 默认元素 (保留 Header 以显示侧边栏按钮) */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 仅让 header 背景透明，不隐藏内容，否则侧边栏收起后无法展开 */
    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }
    
    /* 如果想隐藏右上角的汉堡菜单，可以使用这个 (可选) */
    /* .stApp > header > div:first-child { visibility: hidden; } */
    
    /* 布局微调 */
    .block-container {
        padding-top: 3rem !important;
        padding-bottom: 5rem !important;
        max-width: 1200px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 辅助函数 ---
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
            return False, "缺少邮箱地址", None
            
        msg['To'] = str(recipient).strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(msg_body, 'plain', 'utf-8'))
        
        return True, "就绪", msg
    except Exception as e:
        return False, str(e), None

# --- 侧边栏 ---
with st.sidebar:
    st.markdown("### 系统配置")
    
    # 凭证
    st.markdown("#### 发件人信息")
    sender_name = st.text_input("发件人名称", value=config.SENDER_NAME, placeholder="例如: 客服团队")
    sender_email = st.text_input("发件人邮箱", value=config.SENDER_EMAIL, placeholder="name@company.com")
    sender_password = st.text_input("应用专用密码", value=config.APP_PASSWORD, type="password")
    
    st.markdown("---")
    
    # 设置
    st.markdown("#### 投递参数")
    default_limit = getattr(config, 'BATCH_LIMIT', 0)
    batch_limit = st.number_input("单次发送上限 (0为无限)", min_value=0, value=default_limit)
    
    st.markdown("#### 拟人化设置")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        sleep_min = st.number_input("最小间隔 (秒)", 1.0, 60.0, 2.0)
    with col_s2:
        sleep_max = st.number_input("最大间隔 (秒)", sleep_min, 60.0, 5.0)

# --- 主界面 ---
st.title("智能投递")
st.markdown("<p style='font-size: 1.1rem; color: #666; margin-bottom: 2rem;'>安全、高效的极简邮件分发系统。</p>", unsafe_allow_html=True)

col1, col_spacer, col2 = st.columns([1, 0.1, 1])

with col1:
    st.markdown("## 01. 导入名单")
    uploaded_file = st.file_uploader("将 Excel 名单拖拽至此", type=["xlsx"])
    
    df = None
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            if df.empty:
                st.toast("⚠️ 文件是空的", icon="⚠️")
            else:
                # 序号从 1 开始
                df.index = df.index + 1
                
                # 智能展示：如果数据量大，固定高度以支持滑动；如果数据少，自动适应
                # height 参数控制容器高度，数据过多时会自动出现滚动条
                display_height = min(len(df) * 35 + 38, 300) 
                
                st.dataframe(df, height=display_height, use_container_width=True)
                st.markdown(f"<p style='font-size: 0.9rem; color: #666; margin-top: 0.5rem;'>✓ 已加载 {len(df)} 位收件人 (支持滑动查看)</p>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"文件读取错误: {e}")

with col2:
    st.markdown("## 02. 编辑内容")
    
    # 模板加载
    try:
        with open("template.txt", "r") as f: default_template = f.read()
    except:
        default_template = "你好 {姓名},\n\n..."

    default_subject = getattr(config, 'EMAIL_SUBJECT', "通知")
    email_subject = st.text_input("邮件标题", value=default_subject)
    template_content = st.text_area("邮件正文模板", value=default_template, height=250)
    
    if st.button("保存模板", type="secondary"):
        with open("template.txt", "w") as f:
            f.write(template_content)
        st.toast("模板已保存")

# --- 操作区域 ---
if df is not None and not df.empty:
    st.markdown("---")
    st.markdown("## 03. 预览与投递")
    
    # 变量提取
    placeholders = set(re.findall(r'\{(.*?)\}', template_content))
    missing_cols = [p for p in placeholders if p not in df.columns]
    
    if missing_cols:
        st.error(f"Excel 缺少对应列: {', '.join(missing_cols)}")
    else:
        # 预览卡片
        with st.container():
            st.markdown("#### 效果预览")
            preview_row = df.iloc[0]
            preview_body = template_content
            for key in placeholders:
                preview_body = preview_body.replace(f"{{{key}}}", smart_str(preview_row.get(key)))
            
            preview_html = f"""
            <div style="background-color: #fafafa; border: 1px solid #eaeaea; padding: 1.5rem; border-radius: 8px; font-family: -apple-system, sans-serif;">
                <div style="margin-bottom: 0.5rem; font-size: 0.9rem; color: #666;">
                    <strong>收件人:</strong> {preview_row.get('邮箱') or preview_row.get('Email')}<br>
                    <strong>标题:</strong> {email_subject}
                </div>
                <hr style="margin: 1rem 0; border-color: #eaeaea;">
                <div style="white-space: pre-wrap; font-size: 0.95rem; line-height: 1.5;">{preview_body}</div>
            </div>
            """
            st.markdown(preview_html, unsafe_allow_html=True)

        st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)
        
        # 发送按钮
        if st.button("启动投递任务", type="primary", use_container_width=True):
            if not sender_email or not sender_password:
                st.error("请先在左侧侧边栏配置发件人信息。")
                st.stop()
                
            # 分批逻辑
            if batch_limit > 0 and len(df) > batch_limit:
                task_df = df.iloc[:batch_limit].copy()
                remaining_df = df.iloc[batch_limit:].copy()
            else:
                task_df = df.copy()
                remaining_df = pd.DataFrame()
                
            # 进度界面
            progress_bar = st.progress(0)
            status_text = st.empty()
            processed_records = []
            
            # SMTP 连接
            try:
                with st.spinner(f"正在验证账号 {sender_email}..."):
                    if config.SMTP_PORT == 465:
                        server = smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT)
                    else:
                        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
                        server.starttls()
                    server.login(sender_email, sender_password)
            except Exception as e:
                st.error(f"连接失败: {e}")
                st.stop()
                
            # 发送循环
            total = len(task_df)
            success_count = 0
            
            for i, (index, row) in enumerate(task_df.iterrows()):
                name = row.get('账号', row.get('姓名', '未知'))
                status_text.markdown(f"<span style='color: #666; font-size: 0.9rem;'>正在投递 {i+1}/{total}: <strong>{name}</strong></span>", unsafe_allow_html=True)
                
                is_ready, msg_str, msg_obj = send_one_email(row, template_content, placeholders, email_subject, sender_name, sender_email)
                
                if is_ready:
                    try:
                        server.sendmail(sender_email, msg_obj['To'], msg_obj.as_string())
                        status, detail = "成功", "OK"
                        success_count += 1
                    except Exception as e:
                        status, detail = "失败", str(e)
                else:
                    status, detail = "失败", msg_str
                    
                record = row.to_dict()
                record.update({'发送状态': status, '详情': detail, '发送时间': datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                processed_records.append(record)
                
                progress_bar.progress((i + 1) / total)
                
                if i < total - 1:
                    import random
                    time.sleep(random.uniform(sleep_min, sleep_max))
            
            server.quit()
            status_text.empty()
            
            # 结果处理
            if processed_records:
                # 记录归档
                history_file = "sent_history.xlsx"
                new_recs = pd.DataFrame(processed_records)
                try:
                    if os.path.exists(history_file):
                        pd.concat([pd.read_excel(history_file), new_recs]).to_excel(history_file, index=False)
                    else:
                        new_recs.to_excel(history_file, index=False)
                except: pass
                
                # 成功提示
                st.success(f"任务完成。成功: {success_count}, 失败: {total-success_count}。")
                
                # 下载报告
                output_log = BytesIO()
                new_recs.to_excel(output_log, index=False)
                
                col_d1, col_d2 = st.columns([1, 1])
                with col_d1:
                    st.download_button(
                        label="下载本次发送报告",
                        data=output_log.getvalue(),
                        file_name=f"发送报告_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                if not remaining_df.empty:
                    with col_d2:
                        rem_out = BytesIO()
                        remaining_df.to_excel(rem_out, index=False)
                        st.download_button(
                            label=f"下载剩余名单 ({len(remaining_df)}人)",
                            data=rem_out.getvalue(),
                            file_name=f"剩余名单_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
