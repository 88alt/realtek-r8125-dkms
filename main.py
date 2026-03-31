import os, random, string, requests, time, sys
from datetime import datetime

# ================= 🛡️ 核心配置区域 =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
# 标识当前运行的仓库（从 YAML 传入）
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

# API 端点
TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock"

# ================= 🔐 鉴权模块 =================
def get_access_token():
    data = {
        'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN, 'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    r = requests.post(TOKEN_URL, data=data)
    if r.status_code != 200:
        print(f"!! 令牌刷新失败: {r.text}")
        sys.exit(1)
    return r.json()['access_token']

def get_headers(token):
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# ================= 🚧 抢占锁逻辑 (解决多账号冲突) =================
def try_lock(token, today_str, current_period):
    headers = get_headers(token)
    lock_file = f"lock_{today_str}_{current_period}.json"
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content"
    
    lock_data = {
        "locked_by": GITHUB_REPO,
        "utctime": datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'),
        "status": "active"
    }

    # 【核心参数】fail 表示如果文件已存在则报错，不覆盖
    params = {"@microsoft.graph.conflictBehavior": "fail"}
    
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data, params=params)
        if r.status_code == 201:
            print(f">>> [Lock] 抢占成功！本时段由 {GITHUB_REPO} 执行。")
            return True
        elif r.status_code == 409:
            print(f">>> [Lock] 抢占失败：此时间段已有其他账号完成任务。")
            return False
        return False
    except:
        return False

# ================= 🚀 业务任务模块 =================
def task_read_calendar(token):
    print(">>> 执行任务: 读取日历")
    requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=get_headers(token))

def task_update_log(token, today_str):
    print(">>> 执行任务: 更新日志")
    url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    headers = get_headers(token)
    r = requests.get(url, headers=headers)
    old_content = r.text if r.status_code == 200 else "Time,Repo,Event"
    new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
    requests.put(url, headers=headers, data=(old_content + new_row).encode('utf-8'))

def task_send_mail_smart(token, today_str):
    print(">>> 执行任务: 智能邮件检查")
    headers = get_headers(token)
    log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    
    # 检查今天是否发过邮件
    r = requests.get(log_url, headers=headers)
    if r.status_code == 200 and f"{today_str},MAIL_SENT" in r.text:
        print("    ℹ️ 今日邮件已发送，跳过。")
        return

    # 发送邮件逻辑
    r_me = requests.get(f'{GRAPH_URL}/me', headers=headers)
    my_email = r_me.json().get('userPrincipalName')
    mail_data = {
        "message": {
            "subject": f"KeepAlive Report: {today_str}",
            "body": {"contentType": "Text", "content": f"Success by {GITHUB_REPO}"},
            "toRecipients": [{"emailAddress": {"address": my_email}}]
        },
        "saveToSentItems": False
    }
    if requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data).status_code in [200, 202]:
        print(f"    ✅ 邮件已发至 {my_email}")
        # 记录发信标记
        mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
        requests.put(log_url, headers=headers, data=(r.text + mark).encode('utf-8'))

# ================= 🏁 主入口 =================
def main():
    token = get_access_token()
    today_str = datetime.now().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H') # 以小时作为时段锁

    # 1. 抢位：如果抢不到说明本时段已有账号执行过
    if not try_lock(token, today_str, current_period):
        sys.exit(0)

    # 2. 抢到位的账号执行保活任务
    task_read_calendar(token)
    task_update_log(token, today_str)
    
    # 3. 每天只在第一次成功时发邮件
    task_send_mail_smart(token, today_str)

    print(f"\n>>> [Success] 任务圆满完成")

if __name__ == '__main__':
    main()
