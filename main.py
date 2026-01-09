import os
import glob
import time
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from tqdm import tqdm
import config
from datetime import datetime

def find_excel_file():
    """锁定查找 main.xlsx 文件"""
    target_file = "main.xlsx"
    
    if not os.path.exists(target_file):
        print(f"❌ 错误: 未找到数据源文件 '{target_file}'。")
        print("👉 请确保Excel文件名为 main.xlsx 并放入此文件夹。")
        return None
    
    print(f"✅ 锁定数据源: {target_file}")
    return target_file

def load_template():
    """读取模板并分析需要的列名"""
    try:
        with open("template.txt", "r", encoding="utf-8") as f:
            content = f.read()
        
        import re
        placeholders = set(re.findall(r'\{(.*?)\}', content))
        print(f"✅ 读取模板成功，检测到变量: {placeholders}")
        return content, placeholders
    except FileNotFoundError:
        print("❌ 错误: 未找到 template.txt 邮件模板。")
        return None, None

def smart_str(val):
    """智能转换字符串，处理 123.0 这种情况"""
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        # 如果是整数浮点数 (如 123.0)，转为整数
        if val.is_integer():
            return str(int(val))
    return str(val).strip()

def send_email(server, row, template_content, placeholders):
    """发送单封邮件"""
    try:
        msg_body = template_content
        for key in placeholders:
            val = row.get(key)
            # 使用智能转换
            msg_body = msg_body.replace(f"{{{key}}}", smart_str(val))
            
        msg = MIMEMultipart()
        msg['From'] = formataddr((config.SENDER_NAME, config.SENDER_EMAIL))
        
        recipient = row.get('邮箱') or row.get('Email') or row.get('email')
        if not recipient or pd.isna(recipient):
            return False, "无有效邮箱地址"
            
        msg['To'] = str(recipient).strip()
        # 从配置读取主题
        msg['Subject'] = getattr(config, 'EMAIL_SUBJECT', "账户通知")
        
        msg.attach(MIMEText(msg_body, 'plain', 'utf-8'))
        
        server.sendmail(config.SENDER_EMAIL, msg['To'], msg.as_string())
        return True, "发送成功"
        
    except Exception as e:
        return False, str(e)

def update_history_and_source(source_path, processed_records, remaining_df):
    """关键功能：将处理过的记录移入历史文件，并更新源文件"""
    print("\n💾 正在保存数据...")
    
    # 1. 追加到历史文件 (带重试)
    history_file = "sent_history.xlsx"
    new_records_df = pd.DataFrame(processed_records)
    
    while True:
        try:
            if os.path.exists(history_file):
                old_history = pd.read_excel(history_file)
                # 确保列一致
                combined = pd.concat([old_history, new_records_df], ignore_index=True)
                combined.to_excel(history_file, index=False)
            else:
                new_records_df.to_excel(history_file, index=False)
            print(f"✅ 已归档 {len(processed_records)} 条记录至 '{history_file}'")
            break # 成功则跳出循环
        except PermissionError:
            print(f"\n⚠️ 无法写入 '{history_file}'。文件可能被打开了。")
            input("👉 请关闭 Excel 文件，然后按回车键重试...")
        except Exception as e:
            print(f"❌ 归档失败 (数据未丢失，仍在内存中): {e}")
            return # 其他错误直接放弃，不敢动源文件

    # 2. 更新源文件 (带重试)
    while True:
        try:
            # 显式保留表头 header=True
            remaining_df.to_excel(source_path, index=False, header=True)
            
            if remaining_df.empty:
                print(f"✅ 源文件 '{source_path}' 已清空 (任务完成，仅保留表头)")
            else:
                print(f"✅ 源文件 '{source_path}' 已更新，剩余 {len(remaining_df)} 待发 (表头已保留)")
            break
        except PermissionError:
            print(f"\n⚠️ 无法写入源文件 '{source_path}'。文件可能被打开了。")
            input("👉 请关闭 Excel 文件，然后按回车键重试...")
        except Exception as e:
            print(f"❌ 更新源文件失败: {e}")
            break

def main():
    print("--- 🚀 Smart Mail Drop (自动归档版) ---")
    
    # 1. 资源准备
    excel_path = find_excel_file()
    if not excel_path: return
    
    template_content, placeholders = load_template()
    if not template_content: return
    
    try:
        df = pd.read_excel(excel_path)
        if df.empty:
            print("🎉 列表为空，所有任务已完成！")
            return
            
        # 检查列
        missing_cols = [p for p in placeholders if p not in df.columns]
        if missing_cols:
            print(f"❌ Excel 缺少模板中对应的列: {missing_cols}")
            return
            
    except Exception as e:
        print(f"❌ 读取Excel失败: {e}")
        return

    # 2. 分批逻辑
    limit = getattr(config, 'BATCH_LIMIT', 0)
    if limit > 0 and len(df) > limit:
        task_df = df.iloc[:limit].copy()
        remaining_df = df.iloc[limit:].copy()
        print(f"📋 分批模式: 本次发送前 {len(task_df)} 封 (剩余 {len(remaining_df)} 封)")
    else:
        task_df = df.copy()
        remaining_df = pd.DataFrame()
        print(f"📋 全量模式: 发送所有 {len(task_df)} 封")

    # 3. 连接服务器
    print("🔌 连接 Gmail...", end="")
    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.SENDER_EMAIL, config.APP_PASSWORD)
        print(" 成功!")
    except Exception as e:
        print(f"\n❌ 登录失败: {e}")
        return

    # 4. 执行发送
    processed_records = []
    print("\n📨 开始投递...")
    pbar = tqdm(total=len(task_df), unit="封")
    
    try:
        for index, row in task_df.iterrows():
            success, msg = send_email(server, row, template_content, placeholders)
            
            # 构造归档记录 (复制原行数据 + 状态)
            record = row.to_dict()
            record['发送状态'] = "成功" if success else "失败"
            record['详情'] = msg
            record['发送时间'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            processed_records.append(record)
            
            pbar.update(1)
            
            # 模拟人类延时
            if index < len(task_df) - 1:
                if (index + 1) % 50 == 0:
                    time.sleep(30)
                else:
                    import random
                    time.sleep(random.uniform(2, 5))
                    
    except KeyboardInterrupt:
        print("\n⚠️ 用户中断! 正在保存已处理的数据...")
        # 即使中断，也要把已经发了的那些归档
        remaining_in_task = task_df.iloc[len(processed_records):]
        if not remaining_in_task.empty:
             remaining_df = pd.concat([remaining_in_task, remaining_df])
    finally:
        pbar.close()
        server.quit()

    # 5. 归档与清理
    if processed_records:
        update_history_and_source(excel_path, processed_records, remaining_df)
    else:
        print("无数据处理")

if __name__ == "__main__":
    main()
