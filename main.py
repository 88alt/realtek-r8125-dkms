import os
import random
import string
import requests
import time
import sys
from datetime import datetime

# ================= 🛡️ 核心配置区域 (保持原有) =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

if not TENANT_ID:
    print("!! 致命错误: 缺少 TENANT_ID 环境变量")
    sys.exit(1)

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock"

# ================= 🔐 鉴权模块 (保持原有) =================
def get_access_token():
    print(">>> [Auth] 正在刷新令牌...")
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data)
        if r.status_code != 200:
            print(f"!! 令牌刷新失败 [HTTP {r.status_code}]")
            sys.exit(1)
        return r.json()['access_token']
    except Exception as e:
        print(f"!! 网络请求异常: {e}")
        sys.exit(1)

# ================= 🚧 抢占逻辑 (不打架的核心) =================
def try_lock(token, today_str, current_period):
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    lock_file = f"lock_{today_str}_{current_period}.json"
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content?@microsoft.graph.conflictBehavior=fail"
    
    lock_data = {"locked_by": GITHUB_REPO, "time": datetime.utcnow().strftime('%H:%M:%S')}
    
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data)
        if r.status_code == 201:
            print(f"✅ [Lock] 抢占成功！本时段执行者: {GITHUB_REPO}")
            return True
        elif r.status_code == 409:
            print(f"ℹ️ [Lock] 抢占失败：此时间段已有其他账号在运行。")
            return False
        elif r.status_code == 404:
            print(f"❌ [Lock] 路径不存在，请先手动建立 {LOCK_FOLDER} 文件夹。")
            sys.exit(1)
        return False
    except:
        return False

# ================= 🚀 业务模块 (Task 1-4 保持原有) =================

def task_read_calendar(token):
    print("\n>>> [Task 1] 读取日历")
    headers = {'Authorization': f'Bearer {token}', 'Prefer': 'outlook.timezone="China Standard Time"'}
    try:
        requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers)
        print("    ✅ 日历读取成功")
    except: pass

def task_update_log(token):
    print("\n>>> [Task 2] 更新日志 (CSV)")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    filename = "ActivityLog.csv"
    content_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{filename}:/content'
    try:
        old_content = "Time,Status,ID"
        r = requests.get(content_url, headers=headers)
        if r.status_code == 200:
            old_content = r.text
            lines = old_content.splitlines()
            if len(lines) > 100: old_content = "\n".join(lines[:1] + lines[-99:])
        
        new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},AutoRun,OK,{random.randint(1000,9999)}"
        requests.put(content_url, headers=headers, data=(old_content + new_row).encode('utf-8'))
        print("    ✅ 日志更新成功")
        return old_content
    except: return ""

def task_send_mail(token, old_log, today_str):
    if f"{today_str},MAIL_SENT" in old_log:
        print("\n>>> [Task 3] 邮件跳过：今日已发。")
        return
    print("\n>>> [Task 3] 发送邮件")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    try:
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers)
        my_email = r_me.json().get('userPrincipalName')
        data = {
            "message": {
                "subject": f"KeepAlive: {today_str}",
                "body": {"contentType": "Text", "content": f"Success by {GITHUB_REPO}"},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        if requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=data).status_code in [200, 202]:
            mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
            r_latest = requests.get(f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content', headers=headers)
            requests.put(f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content', headers=headers, data=(r_latest.text + mark).encode('utf-8'))
            print(f"    ✅ 邮件已发送")
    except: pass

def task_upload_large_file(token):
    print("\n>>> [Task 4] 大文件上传")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    file_size = random.randint(1, 10) * 1024 * 1024 # 1-10MB
    file_name = f"Auto_{int(time.time())}_{''.join(random.choices(string.ascii_letters, k=4))}.bin"
    try:
        session_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/createUploadSession'
        r_session = requests.post(session_url, headers=headers, json={"item": {"@microsoft.graph.conflictBehavior": "rename"}})
        upload_url = r_session.json()['uploadUrl']
        requests.put(upload_url, data=b'\0' * file_size, headers={'Content-Length': str(file_size), 'Content-Range': f'bytes 0-{file_size-1}/{file_size}'})
        print(f"    ✅ 上传完成: {file_name}")
    except: pass

    # ================= 🏁 激进清理逻辑 (Task 5) =================
    print("\n>>> [Task 5] 激进清理 (满35留3)")
    
    def cleanup_folder(folder_path, suffix, limit=35, keep=3):
        try:
            list_url = f'{GRAPH_URL}/me/drive/root:{folder_path}:/children?$select=id,name,createdDateTime'
            r = requests.get(list_url, headers=headers)
            if r.status_code == 200:
                items = [x for x in r.json().get('value', []) if x['name'].endswith(suffix)]
                if len(items) >= limit:
                    # 按时间排序，旧的在前
                    items.sort(key=lambda x: x['createdDateTime'])
                    # 需要删除的数量 = 总数 - 保留数
                    to_delete = items[:(len(items) - keep)]
                    for item in to_delete:
                        requests.delete(f'{GRAPH_URL}/me/drive/items/{item["id"]}', headers=headers)
                        print(f"        -> 清理旧文件: {item['name']}")
                else:
                    print(f"        -> {folder_path} 文件数 ({len(items)}) 未达标，跳过。")
        except: pass

    # 清理 Data 文件夹下的 .bin
    cleanup_folder(DATA_FOLDER, ".bin")
    # 清理 Lock 文件夹下的 .json
    cleanup_folder(LOCK_FOLDER, ".json")

# ================= 🏁 主入口 =================
def main():
    token = get_access_token()
    today_str = datetime.now().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H')

    if try_lock(token, today_str, current_period):
        task_read_calendar(token)
        old_log = task_update_log(token)
        task_send_mail(token, old_log, today_str)
        task_upload_large_file(token) # 内部包含 Task 5 清理
        print("\n>>> 所有操作成功完成")
    else:
        sys.exit(0)

if __name__ == '__main__':
    main()
