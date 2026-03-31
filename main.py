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
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock"

def get_access_token():
    data = {
        'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN, 'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data)
        if r.status_code != 200: sys.exit(1)
        return r.json()['access_token']
    except: sys.exit(1)

def get_headers(token):
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# ================= 🚧 抢占锁 (解决多账号并发) =================
def try_lock(token, today_str, current_period):
    headers = get_headers(token)
    lock_file = f"lock_{today_str}_{current_period}.json"
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content"
    
    lock_data = {"locked_by": GITHUB_REPO, "utctime": datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}
    # 核心：如果文件存在则直接报错 (409 Conflict)
    params = {"@microsoft.graph.conflictBehavior": "fail"}
    
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data, params=params)
        return r.status_code == 201 # 只有新建成功才返回 True
    except: return False

# ================= 🚀 业务任务 =================
def task_execute(token, today_str):
    headers = get_headers(token)
    # 1. 读操作 (保活)
    requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers)
    
    # 2. 写日志 (保活核心)
    log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    r_log = requests.get(log_url, headers=headers)
    old_content = r_log.text if r_log.status_code == 200 else "Time,Repo,Event"
    new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
    requests.put(log_url, headers=headers, data=(old_content + new_row).encode('utf-8'))
    
    # 3. 智能邮件 (每日限一封)
    if f"{today_str},MAIL_SENT" not in (r_log.text if r_log.status_code == 200 else ""):
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers)
        my_email = r_me.json().get('userPrincipalName')
        mail_data = {
            "message": {
                "subject": f"KeepAlive: {today_str}",
                "body": {"contentType": "Text", "content": f"Success by {GITHUB_REPO}"},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        if requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data).status_code in [200, 202]:
            mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
            # 重新获取最新内容防止覆盖
            r_latest = requests.get(log_url, headers=headers)
            requests.put(log_url, headers=headers, data=(r_latest.text + mark).encode('utf-8'))

def main():
    token = get_access_token()
    today_str = datetime.now().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H')

    # 抢位逻辑：谁先抢到锁谁执行
    if try_lock(token, today_str, current_period):
        task_execute(token, today_str)
        print(f">>> 任务由 {GITHUB_REPO} 成功完成")
    else:
        print(">>> 该时段任务已被抢占，安全跳过")

if __name__ == '__main__':
    main()
