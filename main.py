import os, random, string, requests, time, sys
from datetime import datetime

# ================= 🛡️ 核心配置 (通用变量名) =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "Data"
LOCK_FOLDER = "Data/Lock"

def get_access_token():
    print(">>> [Step 1] 正在刷新 Access Token...")
    data = {
        'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN, 'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data, timeout=20)
        if r.status_code != 200:
            print(f"!! [Auth Error] HTTP {r.status_code}: {r.text}")
            sys.exit(1)
        return r.json()['access_token']
    except Exception as e:
        print(f"!! [Auth Exception] {e}")
        sys.exit(1)

def get_headers(token):
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# ================= 🚧 抢占锁逻辑 (解决并发/多账号重复) =================
def try_lock(token, today_str, current_period):
    headers = get_headers(token)
    lock_file = f"lock_{today_str}_{current_period}.json"
    
    # 构造请求 URL，直接拼接 conflictBehavior 参数
    lock_url = f"{GRAPH_URL}/me/drive/root:/{LOCK_FOLDER}/{lock_file}:/content?@microsoft.graph.conflictBehavior=fail"
    
    lock_data = {
        "locked_by": GITHUB_REPO, 
        "utctime": datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    print(f">>> [Step 2] 尝试抢占时段锁: {current_period} 点档...")
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data, timeout=30)
        if r.status_code == 201:
            print(f"✅ [Lock] 抢占成功！本次任务由 {GITHUB_REPO} 执行。")
            return True
        elif r.status_code == 409:
            print(f"ℹ️ [Lock] 抢占失败：此时间段已有其他账号在执行。")
            return False
        elif r.status_code == 404:
            print(f"❌ [Lock] 路径不存在！请在网盘根目录手动创建 /Data/Lock 文件夹。")
            sys.exit(1)
        else:
            print(f"!! [Lock Error] HTTP {r.status_code}: {r.text}")
            sys.exit(1)
    except Exception as e:
        print(f"!! [Lock Exception] {e}")
        sys.exit(1)

# ================= 🚀 业务逻辑 =================
def task_execute(token, today_str):
    headers = get_headers(token)
    print(">>> [Step 3] 执行保活与通知任务...")
    
    # 1. 保活读操作 (日历)
    requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers, timeout=20)
    
    # 2. 更新日志 (ActivityLog.csv)
    log_path = f"{DATA_FOLDER}/ActivityLog.csv"
    log_url = f"{GRAPH_URL}/me/drive/root:/{log_path}:/content"
    
    print("    -> 正在更新网盘日志...")
    r_get_log = requests.get(log_url, headers=headers, timeout=20)
    log_text = r_get_log.text if r_get_log.status_code == 200 else "Time,Repo,Event"
    
    new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
    updated_log = log_text + new_row
    requests.put(log_url, headers=headers, data=updated_log.encode('utf-8'), timeout=30)
    
    # 3. 每日邮件通知 (每天仅限一次)
    if f"{today_str},MAIL_SENT" not in log_text:
        print("    -> 今日尚未发信，准备发送日报邮件...")
        try:
            r_me = requests.get(f'{GRAPH_URL}/me', headers=headers, timeout=20)
            my_email = r_me.json().get('userPrincipalName')
            
            mail_body = {
                "message": {
                    "subject": f"KeepAlive Daily: {today_str}",
                    "body": {"contentType": "Text", "content": f"Performed by {GITHUB_REPO}"},
                    "toRecipients": [{"emailAddress": {"address": my_email}}]
                },
                "saveToSentItems": False
            }
            m_res = requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_body, timeout=30)
            if m_res.status_code in [200, 202]:
                print(f"    ✅ 邮件已发送至 {my_email}")
                # 记录发信标记
                mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
                r_latest = requests.get(log_url, headers=headers, timeout=20)
                requests.put(log_url, headers=headers, data=(r_latest.text + mark).encode('utf-8'))
        except Exception as e:
            print(f"    ⚠️ 邮件发送失败: {e}")
    else:
        print("    ℹ️ 今日邮件已由之前任务发送，跳过。")

def main():
    # 环境变量预检
    if not all([CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN, TENANT_ID]):
        print("!! [Config Error] 缺少环境变量。请检查 GitHub Secrets 名称。")
        sys.exit(1)

    token = get_access_token()
    today_str = datetime.now().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H') # 使用 UTC 小时作为时段键值

    if try_lock(token, today_str, current_period):
        task_execute(token, today_str)
        print(f"\n>>> [Success] 账号 {GITHUB_REPO} 任务圆满完成。")
    else:
        print(f"\n>>> [Skip] 账号 {GITHUB_REPO} 放弃执行权（该时段已完成）。")
        sys.exit(0)

if __name__ == '__main__':
    main()
