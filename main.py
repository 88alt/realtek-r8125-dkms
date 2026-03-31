import os, random, string, requests, time, sys
from datetime import datetime

# ================= 🛡️ 核心配置 =================
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
        r = requests.post(TOKEN_URL, data=data, timeout=20)
        r.raise_for_status()
        return r.json()['access_token']
    except Exception as e:
        print(f"!! [Auth] 令牌刷新失败: {e}")
        sys.exit(1)

def get_headers(token):
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# ================= 🚧 抢占锁逻辑 =================
def try_lock(token, today_str, current_period):
    headers = get_headers(token)
    lock_file = f"lock_{today_str}_{current_period}.json"
    # 注意：这里直接将 conflictBehavior 拼接到 URL 中，防止某些环境下的 Params 失效
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content?@microsoft.graph.conflictBehavior=fail"
    
    lock_data = {
        "locked_by": GITHUB_REPO, 
        "utctime": datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data, timeout=30)
        if r.status_code == 201:
            print(f">>> [Lock] 抢占成功！本时段执行者: {GITHUB_REPO}")
            return True
        elif r.status_code == 409:
            print(f">>> [Lock] 抢占失败：此时间段已有其他任务在运行。")
            return False
        elif r.status_code == 404:
            print(f"!! [Lock] 错误：网盘路径 {LOCK_FOLDER} 不存在，请手动创建文件夹。")
            sys.exit(1)
        else:
            print(f"!! [Lock] 异常状态码 {r.status_code}: {r.text}")
            sys.exit(1)
    except Exception as e:
        print(f"!! [Lock] 请求异常: {e}")
        sys.exit(1)

# ================= 🚀 业务逻辑 =================
def task_execute(token, today_str):
    headers = get_headers(token)
    
    # 1. 保活读操作
    requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers, timeout=20)
    
    # 2. 检查日志与邮件状态
    log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    r_log = requests.get(log_url, headers=headers, timeout=20)
    
    log_text = ""
    if r_log.status_code == 200:
        log_text = r_log.text
    
    # 写保活日志
    new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
    final_content = (log_text if log_text else "Time,Repo,Event") + new_row
    requests.put(log_url, headers=headers, data=final_content.encode('utf-8'), timeout=30)
    
    # 3. 每日一封邮件
    if f"{today_str},MAIL_SENT" not in log_text:
        try:
            r_me = requests.get(f'{GRAPH_URL}/me', headers=headers, timeout=20)
            my_email = r_me.json().get('userPrincipalName')
            
            mail_data = {
                "message": {
                    "subject": f"KeepAlive Daily: {today_str}",
                    "body": {"contentType": "Text", "content": f"Performed by {GITHUB_REPO}"},
                    "toRecipients": [{"emailAddress": {"address": my_email}}]
                },
                "saveToSentItems": False
            }
            res = requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data, timeout=30)
            if res.status_code in [200, 202]:
                print(f"✅ 邮件已发送至 {my_email}")
                # 记录发信成功标记
                mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
                # 再次获取最新内容追加
                r_upd = requests.get(log_url, headers=headers, timeout=20)
                requests.put(log_url, headers=headers, data=(r_upd.text + mark).encode('utf-8'))
        except Exception as e:
            print(f"!! [Mail] 发送失败: {e}")

def main():
    # 检查环境变量是否成功注入
    if not all([CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN, TENANT_ID]):
        print("!! [Config] 错误：GitHub Secrets 变量读取为空，请检查配置。")
        sys.exit(1)

    token = get_access_token()
    today_str = datetime.now().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H')

    if try_lock(token, today_str, current_period):
        task_execute(token, today_str)
        print(f"\n>>> [Done] 任务圆满完成")
    else:
        # 抢锁失败不属于脚本错误，以 code 0 正常退出
        sys.exit(0)

if __name__ == '__main__':
    main()
