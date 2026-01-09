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

# --- 设置页面配置 ---
st.set_page_config(page_title="Smart Mail Drop", page_icon="📨", layout="wide")

# --- 辅助函数 ---
def smart_str(val):
    if pd.isna(val): return ""
    if isinstance(val, float):
        if val.is_integer(): return str(int(val))
    return str(val).strip()

def send_one_email(row, template_content, placeholders, subject, s_name, s_email):
    """发送逻辑核心 - 接收动态账号参数"""
    try:
        msg_body = template_content
        for key in placeholders:
            val = row.get(key)
            msg_body = msg_body.replace(f"{{{key}}}", smart_str(val))
            
        msg = MIMEMultipart()
        msg['From'] = formataddr((s_name, s_email))
        
        recipient = row.get('邮箱') or row.get('Email') or row.get('email')
        if not recipient or pd.isna(recipient):
            return False, "无有效邮箱地址", None
            
        msg['To'] = str(recipient).strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(msg_body, 'plain', 'utf-8'))
        
        return True, "准备发送", msg
    except Exception as e:
        return False, str(e), None

# --- 侧边栏 ---
with st.sidebar:
    st.title("⚙️ 发送配置")
    
    # 动态配置区
    with st.expander("👤 账号设置", expanded=True):
        sender_name = st.text_input("发件人名称", value=config.SENDER_NAME)
        sender_email = st.text_input("发件人邮箱", value=config.SENDER_EMAIL)
        sender_password = st.text_input("应用专用密码", value=config.APP_PASSWORD, type="password", help="请使用Google两步验证生成的16位应用专用密码")
    
    # 默认值保护
    default_limit = getattr(config, 'BATCH_LIMIT', 0)
    batch_limit = st.number_input("单次发送数量 (0=无限)", min_value=0, value=default_limit)
    
    st.divider()
    st.write("🤖 **人类模拟设置**")
    sleep_min = st.slider("最小间隔 (秒)", 1.0, 10.0, 2.0)
    sleep_max = st.slider("最大间隔 (秒)", sleep_min, 20.0, 5.0)

# --- 主界面 ---
st.title("📨 Smart Mail Drop")

# 1. 数据加载区 (支持上传)
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 导入名单")
    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            if df.empty:
                st.warning("⚠️ 文件是空的")
            else:
                st.success(f"✅ 已加载 (共 {len(df)} 人)")
                st.dataframe(df.head(5), height=200)
        except Exception as e:
            st.error(f"❌ 读取失败: {e}")
            df = None
    else:
        st.info("👋 请先上传包含收件人的 Excel 文件")
        df = None

# 2. 模板编辑区
with col2:
    st.subheader("2. 邮件内容")
    try:
        with open("template.txt", "r") as f:
            default_template = f.read()
    except:
        default_template = "你好 {UID}..."
        
    default_subject = getattr(config, 'EMAIL_SUBJECT', "通知")
    email_subject = st.text_input("邮件标题", value=default_subject)
    template_content = st.text_area("正文模板", value=default_template, height=200)
    
    if st.button("💾 保存模板变更"):
        with open("template.txt", "w") as f:
            f.write(template_content)
        st.toast("模板已保存!", icon="✅")

# 3. 预览与操作
if df is not None and not df.empty:
    st.divider()
    
    # 提取变量
    import re
    placeholders = set(re.findall(r'\{(.*?)\}', template_content))
    missing_cols = [p for p in placeholders if p not in df.columns]
    
    if missing_cols:
        st.error(f"❌ Excel 缺少列: {missing_cols}")
    else:
        # 预览
        with st.expander("👁️ 预览第一封邮件"):
            preview_row = df.iloc[0]
            preview_body = template_content
            for key in placeholders:
                preview_body = preview_body.replace(f"{{{key}}}", smart_str(preview_row.get(key)))
            st.markdown(f"**From**: `{sender_name} <{sender_email}>`")
            st.markdown(f"**To**: `{preview_row.get('邮箱')}`")
            st.markdown(f"**Subject**: `{email_subject}`")
            st.text(preview_body)

        # 启动按钮
        st.write("") # Spacer
        if st.button("🚀 开始发送", type="primary", use_container_width=True):
            # 校验
            if not sender_email or not sender_password:
                st.error("❌ 请先在左侧侧边栏填入发件人邮箱和密码！")
                st.stop()
                
            # 确定发送列表
            if batch_limit > 0 and len(df) > batch_limit:
                task_df = df.iloc[:batch_limit].copy()
                remaining_df = df.iloc[batch_limit:].copy()
            else:
                task_df = df.copy()
                remaining_df = pd.DataFrame()
                
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            processed_records = []
            
            # 连接服务器
            try:
                with st.spinner(f"正在以 {sender_email} 连接服务器..."):
                    if config.SMTP_PORT == 465:
                        server = smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT)
                    else:
                        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
                        server.starttls()
                    
                    server.login(sender_email, sender_password)
            except Exception as e:
                st.error(f"无法连接服务器: {e}")
                st.stop()
                
            # 循环发送
            total = len(task_df)
            success_count = 0
            
            for i, (index, row) in enumerate(task_df.iterrows()):
                name = row.get('账号', row.get('姓名', 'Unknown'))
                status_text.markdown(f"📨 正在发送 ({i+1}/{total}): **{name}**")
                
                # 构造并发送
                is_ready, msg_str, msg_obj = send_one_email(row, template_content, placeholders, email_subject, sender_name, sender_email)
                
                if is_ready:
                    try:
                        server.sendmail(sender_email, msg_obj['To'], msg_obj.as_string())
                        status = "成功"
                        detail = "OK"
                        success_count += 1
                    except Exception as e:
                        status = "失败"
                        detail = str(e)
                else:
                    status = "失败"
                    detail = msg_str
                    
                # 记录
                record = row.to_dict()
                record['发送状态'] = status
                record['详情'] = detail
                record['发送时间'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                processed_records.append(record)
                
                # 进度条
                progress_bar.progress((i + 1) / total)
                
                # 延时
                if i < total - 1:
                    import random
                    sleep_time = random.uniform(sleep_min, sleep_max)
                    time.sleep(sleep_time)
            
            server.quit()
            
            # 结果处理
            if processed_records:
                # 1. 归档日志
                history_file = "sent_history.xlsx"
                new_recs = pd.DataFrame(processed_records)
                try:
                    if os.path.exists(history_file):
                        pd.concat([pd.read_excel(history_file), new_recs]).to_excel(history_file, index=False)
                    else:
                        new_recs.to_excel(history_file, index=False)
                except Exception as e:
                    st.error(f"服务器日志归档失败: {e}")
                    
                st.success(f"🎉 任务完成! 成功: {success_count}, 失败: {total-success_count}")
                st.balloons()
                
                # 2. 生成下载按钮 (核心变更)
                if not remaining_df.empty:
                    st.warning(f"👉 还有 {len(remaining_df)} 人未发送。")
                    
                    output = BytesIO()
                    # 显式保留表头
                    remaining_df.to_excel(output, index=False, header=True)
                    data = output.getvalue()
                    
                    st.download_button(
                        label="📥 点击下载剩余名单.xlsx",
                        data=data,
                        file_name=f"剩余名单_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.success("✨ 所有名单已全部处理完毕！")
