import os
import random
import string
import requests
import time
import sys
from datetime import datetime

# ================= 🛡️ 核心配置区域 =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
# 获取仓库名，用于在日志中区分是哪个账号在干活
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

if not TENANT_ID:
    print("!! 致命错误: 缺少 TENANT_ID 环境变量")
    sys.exit(1)

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock"

# ================= 🔐 鉴权模块 =================
def get_access_token():
    print(">>> [Auth] 刷新令牌...")
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
            print(f"!! 令牌刷新失败: {r.text}")
            sys.exit(1)
        return r.json()['access_token']
    except Exception as e:
        print(f"!! 网络异常: {e}")
        sys.exit(1)

# ================= 🚧 抢占逻辑 (核心：保证全天仅4条记录) =================
def try_lock(token, today_str, current_period):
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    # 锁文件名包含日期和小时，确保每个时段只有一个赢家
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
        else:
            print(f"❌ [Lock] 错误: 请确保网盘已手动创建 {LOCK_FOLDER} 文件夹。")
            sys.exit(1)
    except:
        return False

# ================= 🚀 业务模块 =================

def task_read_calendar(token):
    print("\n>>> [Task 1] 读取日历保活")
    headers = {'Authorization': f'Bearer {token}'}
    try:
        requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers)
        print("    ✅ 操作成功")
    except: pass

def task_update_log(token):
    print("\n>>> [Task 2] 更新日志 (CSV)")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    try:
        old_content = "Time,Repo,Event"
        r = requests.get(log_url, headers=headers)
        if r.status_code == 200:
            old_content = r.text
            # 自动修剪日志，防止单文件过大
            lines = old_content.splitlines()
            if len(lines) > 100: old_content = "\n".join(lines[:1] + lines[-99:])
        
        new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
        requests.put(log_url, headers=headers, data=(old_content + new_row).encode('utf-8'))
        print("    ✅ 日志更新成功")
        return old_content
    except: return ""

def task_send_mail(token, old_log, today_str):
    # 核心：检查日志中今天是否已经发过信
    if f"{today_str},MAIL_SENT" in old_log:
        print("\n>>> [Task 3] 邮件跳过：今日已有账号发过信。")
        return
    
    print("\n>>> [Task 3] 发送每日提醒邮件")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    try:
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers)
        my_email = r_me.json().get('userPrincipalName')
        mail_data = {
            "message": {
                "subject": f"Office365 KeepAlive: {today_str}",
                "body": {"contentType": "Text", "content": f"系统运行正常。\n执行账号: {GITHUB_REPO}"},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        if requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data).status_code in [200, 202]:
            # 发信成功后，在日志中打标，防止其他账号重复发信
            mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
            r_latest = requests.get(f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content', headers=headers)
            requests.put(f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content', headers=headers, data=(r_latest.text + mark).encode('utf-8'))
            print(f"    ✅ 邮件已发送至 {my_email}")
    except: pass

def task_upload_large_file(token):
    print("\n>>> [Task 4] 模拟大文件操作")
    headers = {'Authorization': f'Bearer {token}'}
    file_size = random.randint(1, 5) * 1024 * 1024 # 1-5MB
    file_name = f"Auto_{int(time.time())}.bin"
    try:
        session_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/createUploadSession'
        r_session = requests.post(session_url, headers=headers, json={"item": {"@microsoft.graph.conflictBehavior": "rename"}})
        upload_url = r_session.json()['uploadUrl']
        requests.put(upload_url, data=b'\0' * file_size, headers={'Content-Length': str(file_size), 'Content-Range': f'bytes 0-{file_size-1}/{file_size}'})
        print(f"    ✅ 上传完成: {file_name}")
    except: pass

    # ================= 🏁 Task 5: 激进清理 (满35留3) =================
    print("\n>>> [Task 5] 网盘自动清理 (满35留3)")
    
    def cleanup(folder, suffix):
        try:
            url = f'{GRAPH_URL}/me/drive/root:{folder}:/children?$select=id,name,createdDateTime'
            items = [x for x in requests.get(url, headers=headers).json().get('value', []) if x['name'].endswith(suffix)]
            if len(items) >= 35:
                items.sort(key=lambda x: x['createdDateTime'])
                for item in items[:(len(items) - 3)]:
                    requests.delete(f'{GRAPH_URL}/me/drive/items/{item["id"]}', headers=headers)
                    print(f"        -> 清理旧文件: {item['name']}")
        except: pass

    cleanup(DATA_FOLDER, ".bin") # 清理大文件
    cleanup(LOCK_FOLDER, ".json") # 清理锁文件

# ================= 🏁 主入口 =================
def main():
    token = get_access_token()
    # 使用 UTC 时间确保多账号判定标准一致
    today_str = datetime.utcnow().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H')

    # 1. 尝试抢锁（抢不到说明本时段已有其他账号在干活）
    if try_lock(token, today_str, current_period):
        task_read_calendar(token)
        old_log = task_update_log(token)
        task_send_mail(token, old_log, today_str)
        task_upload_large_file(token)
        print("\n>>> [Done] 本次任务圆满完成。")
    else:
        print("\n>>> [Skip] 任务已由其他账号执行，本账号退出。")
        sys.exit(0)

if __name__ == '__main__':
    main()
