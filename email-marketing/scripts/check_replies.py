import imaplib
import email
import json
import os
from datetime import datetime

# --- 配置信息 (优先从环境变量读取) ---
IMAP_HOST = os.getenv("EMAIL_IMAP_HOST","")
IMAP_PORT = int(os.getenv("EMAIL_IMAP_PORT",465)) if os.getenv("EMAIL_IMAP_PORT") else 465
IMAP_USER = os.getenv("EMAIL_SMTP_USER", "")
IMAP_PASS = os.getenv("EMAIL_SMTP_PASS", "")
REPLY_LOG = "reply_stats.json"

def analyze_bounce_reason(body):
    body_lower = body.lower()
    if "user not found" in body_lower or "invalid user" in body_lower or "不存在" in body:
        return "账号不存在 (Invalid User)"
    elif "spam" in body_lower or "rejected" in body_lower or "垃圾邮件" in body:
        return "触发垃圾邮件风控 (Spam/Rejected)"
    elif "full" in body_lower or "quota" in body_lower or "满" in body:
        return "对方邮箱已满 (Mailbox Full)"
    return "其他原因 (详见摘要)"

def check_replies():
    # 修正已发人数统计：从日志文件中尝试汇总今日数据
    sent_total = 0
    if os.path.exists("email_status.json"):
        try:
            with open("email_status.json", "r") as f:
                data = json.load(f)
                # 如果是单次发信逻辑，这里显示最后一次成功数
                sent_total = data.get("success", 0)
        except: pass

    replied_list = []
    bounce_list = []
    
    # 严格过滤名单：系统号/订阅号不计入回信
    SYSTEM_KEYWORDS = ["noreply", "notification", "service", "support", "no-reply", "report"]

    today_imap = datetime.now().strftime("%d-%b-%Y")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_USER, IMAP_PASS)
        mail.select("INBOX") 

        status, response = mail.search(None, f'(SINCE "{today_imap}")')
        
        if status == 'OK' and response[0]:
            ids = response[0].split()
            for r_id in ids:
                status, data = mail.fetch(r_id, '(RFC822)')
                if status != 'OK': continue
                
                msg = email.message_from_bytes(data[0][1])
                from_addr = msg.get('From', '').lower()
                subject = msg.get('Subject', '')
                
                # 解码标题
                try:
                    subject_parts = email.header.decode_header(subject)
                    decoded_subject = ""
                    for part, encoding in subject_parts:
                        if isinstance(part, bytes):
                            decoded_subject += part.decode(encoding or 'utf-8', errors='ignore')
                        else:
                            decoded_subject += part
                    subject = decoded_subject
                except: pass

                # 提取正文摘要用于退信分析
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            payload = part.get_payload(decode=True)
                            if payload: body = payload.decode('utf-8', errors='ignore')
                            break
                else:
                    payload = msg.get_payload(decode=True)
                    if payload: body = payload.decode('utf-8', errors='ignore')

                # 1. 检查退信
                if "系统退信" in subject or "systems bounce" in subject.lower() or "mailer-daemon" in from_addr:
                    reason = analyze_bounce_reason(body)
                    bounce_list.append({"from": from_addr, "reason": reason, "summary": body[:50]})
                
                # 2. 检查真实回信 (排除自己和系统号)
                else:
                    is_system = any(kw in from_addr for kw in SYSTEM_KEYWORDS)
                    is_self = IMAP_USER.lower() in from_addr
                    
                    if not is_system and not is_self:
                        replied_list.append({"from": from_addr, "subject": subject})

        mail.logout()

        stats = {
            "sent_total": sent_total,
            "replied_total": len(replied_list),
            "bounce_total": len(bounce_list),
            "replied_details": replied_list,
            "bounce_details": bounce_list,
            "check_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        with open(REPLY_LOG, 'w', encoding='utf-8') as f:
            json.dump(stats, f, ensure_ascii=False, indent=4)
        
        return stats

    except Exception as e:
        return None

if __name__ == "__main__":
    results = check_replies()
    if results:
        print("\n--- 邮件营销综合效果报告 ---")
        print(f"统计日期: {datetime.now().strftime('%Y-%m-%d')}")
        print(f"最近一次群发人数: {results['sent_total']}")
        print(f"真实回信人数: {results['replied_total']}")
        print(f"今日退信数量: {results['bounce_total']}")
        
        if results['replied_total'] > 0:
            print("\n[回信详情]:")
            for r in results['replied_details']:
                print(f"- 来自: {r['from']}")
                print(f"  标题: {r['subject']}")
        
        if results['bounce_total'] > 0:
            print("\n[退信分析]:")
            for b in results['bounce_details']:
                print(f"- 原因: {b['reason']}")
