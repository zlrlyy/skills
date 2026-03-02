import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import json
import os
import random
import string
import re

# --- 配置信息 (从环境变量读取) ---
SMTP_HOST = os.getenv("EMAIL_SMTP_HOST", "")
SMTP_PORT = int(os.getenv("EMAIL_SMTP_PORT", 465)) if os.getenv("EMAIL_SMTP_PORT") else 465
SMTP_USER = os.getenv("EMAIL_SMTP_USER", "")
SMTP_PASS = os.getenv("EMAIL_SMTP_PASS", "")
TEST_EMAIL = os.getenv("EMAIL_TEST_TARGET", "")

EXCEL_PATH = os.path.expanduser(os.getenv("EMAIL_EXCEL_PATH", "~/Desktop/邮箱.xlsx"))
HTML_PATH = os.path.expanduser(os.getenv("EMAIL_HTML_PATH", "~/Desktop/邮件内容.html"))
TITLE_TXT_PATH = os.path.expanduser(os.getenv("EMAIL_TITLE_PATH", "~/Desktop/邮件标题.txt"))
LOG_FILE = "email_status.json"

BATCH_SIZE = 135  # 每个批次的收件人数量

def get_title_from_txt(path):
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return f.read().strip()
        return "你好"
    except Exception as e:
        print(f"读取 TXT 标题失败: {e}")
        return "你好"

def generate_random_tag():
    """生成随机隐藏标识符，干扰反垃圾扫描"""
    chars = string.ascii_letters + string.digits
    tag = ''.join(random.choice(chars) for _ in range(8))
    return f'<div style="display:none !important; color:transparent; visibility:hidden; opacity:0; font-size:0px;">ID:{tag}</div>'

def get_email_data():
    """读取 Excel 并识别标题行和数据列"""
    try:
        # 先读取前几行看是否有标题
        peek = pd.read_excel(EXCEL_PATH, nrows=5)
        # 简单判断首行是否包含邮箱格式特征
        has_header = not any("@" in str(col) for col in peek.columns)
        
        if has_header:
            df = pd.read_excel(EXCEL_PATH)
        else:
            df = pd.read_excel(EXCEL_PATH, header=None)
            # 给第一列起个默认名
            df.columns = ["email"] + [f"col_{i}" for i in range(1, len(df.columns))]
            
        # 寻找包含邮箱的那一列
        email_col = None
        for col in df.columns:
            if df[col].astype(str).str.contains("@").any():
                email_col = col
                break
        
        if email_col is None:
            raise ValueError("Excel 中未找到有效的邮箱地址列")
            
        return df, email_col
    except Exception as e:
        print(f"读取 Excel 数据失败: {e}")
        return None, None

def replace_placeholders(text, row_data):
    """将文本中的 【变量名】 替换为行数据中的对应值"""
    if not isinstance(text, str): return text
    
    # 查找所有 【...】 格式的占位符
    placeholders = re.findall(r"【(.*?)】", text)
    for p in placeholders:
        # 在列名中寻找匹配（不区分大小写，去掉空格）
        match_col = None
        clean_p = p.strip().lower()
        for col in row_data.index:
            if str(col).strip().lower() == clean_p:
                match_col = col
                break
        
        if match_col is not None:
            val = str(row_data[match_col])
            if val.lower() == "nan": val = ""
            text = text.replace(f"【{p}】", val)
    return text

def send_bulk_emails(test_mode=True, test_email=None):
    # 加载数据
    df, email_col = get_email_data()
    if df is None: return

    # 获取基础标题
    subject_template = get_title_from_txt(TITLE_TXT_PATH)
    
    # 读取 HTML 模板
    try:
        with open(HTML_PATH, 'r', encoding='utf-8') as f:
            html_template = f.read()
    except Exception as e:
        print(f"读取 HTML 失败: {e}")
        return

    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    results = {"total": 1 if test_mode else len(df), "success": 0, "failed": [], "start_time": time.strftime("%Y-%m-%d %H:%M:%S")}

    # 准备任务列表
    if test_mode:
        # 测试模式下，取第一行数据（如果有）来模拟替换效果，但发送给 test_email
        row = df.iloc[0].copy()
        row[email_col] = test_email
        task_list = [row]
    else:
        task_list = [row for _, row in df.iterrows()]

    # 分批发送
    for i in range(0, len(task_list), BATCH_SIZE):
        batch = task_list[i:i + BATCH_SIZE]
        print(f"正在发送批次 {i//BATCH_SIZE + 1}，包含 {len(batch)} 个收件人...")
        
        try:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
                server.login(SMTP_USER, SMTP_PASS)
                
                for row in batch:
                    addr = str(row[email_col])
                    try:
                        # 执行变量替换
                        final_subject = replace_placeholders(subject_template, row)
                        final_html = replace_placeholders(html_template, row)
                        
                        # 插入干扰码
                        final_html += generate_random_tag()
                        
                        msg = MIMEMultipart()
                        msg['From'] = f"Youdao Ads <{SMTP_USER}>"
                        msg['To'] = addr
                        msg['Subject'] = final_subject
                        msg.attach(MIMEText(final_html, 'html', 'utf-8'))

                        server.send_message(msg)
                        results["success"] += 1
                        print(f"成功发送至: {addr}")
                        
                        # 降低发信频率，增加随机性
                        if not test_mode:
                            # 基础间隔 + 随机波动
                            time.sleep(random.uniform(3.0, 8.0))
                            # 每发送 10 封进行一次长休息
                            if results["success"] % 10 == 0:
                                time.sleep(random.uniform(10, 20))
                        
                    except Exception as e:
                        print(f"发送至 {addr} 失败: {e}")
                        results["failed"].append({"email": addr, "reason": str(e)})
                
                if not test_mode: time.sleep(2)
                
        except Exception as e:
            print(f"连接失败: {e}")

    results["end_time"] = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=4)
    return results

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "run":
        send_bulk_emails(test_mode=False)
    else:
        test_addr = TEST_EMAIL
        if not test_addr:
            print("错误: 未配置 EMAIL_TEST_TARGET 环境变量")
            sys.exit(1)
        print(f"准备发送测试邮件至: {test_addr}")
        send_bulk_emails(test_mode=True, test_email=test_addr)
