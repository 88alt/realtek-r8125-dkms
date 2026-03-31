import os
import random
import string
import requests
import time
import sys
from datetime import datetime

# ================= 🛡️ 核心配置区域 (通用版) =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
# 获取仓库名用于标识抢占者
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

# API 端点
if not TENANT_ID:
    print("!! 致命错误: 缺少 TENANT_ID 环境变量")
    sys.exit(1)

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock" # 锁文件存放位置

# ================= 🔐 鉴权模块 (Offline Access) =================
def get_access_token():
    print(">>> [Auth] 正在刷新令牌...")
    if not CLIENT_ID or not CLIENT_SECRET or not REFRESH_TOKEN:
        print("!! Secrets 读取失败，请检查 GitHub Workflow 的 env 映射")
        sys.exit(1)

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
            print(f"!! 详情: {r.text}")
            sys.exit(1)
        return r.json()['access_token']
    except Exception as e:
        print(f"!! 网络请求异常: {e}")
        sys.exit(1)

# ================= 🚧 抢占逻辑 (解决多账号冲突) =================
def try_lock(token, today_str, current_period):
    headers = get_headers(token)
    lock_file = f"lock_{today_str}_{current_period}.json"
    # 使用 fail 策略：文件存在则返回 409
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content?@microsoft.graph.conflictBehavior=fail"
    
    lock_data = {"locked_by": GITHUB_REPO, "time": datetime.utcnow().strftime('%H:%M:%S')}
    
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data)
        if r.status_code == 201:
            print(f"✅ [Lock] 抢占成功！本次时段由 {GITHUB_REPO} 执行。")
            return True
        elif r.status_code == 409:
            print(f"ℹ️ [Lock] 抢占失败：此时间段已有其他账号在执行。")
            return False
        elif r.status_code == 404:
            print(f"❌ [Lock] 路径不存在，请先在网盘建立 {LOCK_FOLDER} 文件夹。")
            sys.exit(1)
        return False
    except:
        return False

# ================= 🛠️ 辅助工具箱 (不变) =================
def get_headers(token):
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

def random_string(length=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# ================= 🚀 业务模块 (原封不动) =================

def task_read_calendar(token):
    print("\n>>> [Task 1] 读取日历")
    headers = get_headers(token)
    try:
        headers['Prefer'] = 'outlook.timezone="China Standard Time"'
        requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers)
        print("    ✅ 日历读取成功")
    except Exception as e:
        print(f"    ⚠️ 日历读取跳过: {e}")

def task_update_log(token):
    print("\n>>> [Task 2] 更新日志 (CSV)")
    headers = get_headers(token)
    filename = "ActivityLog.csv"
    path_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{filename}'
    
    try:
        content_url = f'{path_url}:/content'
        old_content = "Time,Status,ID"
        r = requests.get(content_url, headers=headers)
        if r.status_code == 200:
            old_content = r.text
            lines = old_content.splitlines()
            if len(lines) > 200:
                old_content = "\n".join(lines[:1] + lines[-199:])

        new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},AutoRun,OK,{random.randint(1000,9999)}"
        final_data = old_content + new_row
        requests.put(content_url, headers=headers, data=final_data.encode('utf-8'))
        print("    ✅ 日志更新成功")
        return old_content # 返回用于邮件判断
    except Exception as e:
        print(f"    ⚠️ 日志操作异常: {e}")
        return ""

def task_send_mail(token, old_log_content, today_str):
    # 增加逻辑：每天只发一封
    if f"{today_str},MAIL_SENT" in old_log_content:
        print("\n>>> [Task 3] 邮件跳过：今日已发送。")
        return

    print("\n>>> [Task 3] 发送邮件")
    headers = get_headers(token)
    try:
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers)
        if r_me.status_code != 200: return
        my_email = r_me.json().get('userPrincipalName')
        
        data = {
            "message": {
                "subject": f"KeepAlive: {today_str}",
                "body": {"contentType": "Text", "content": f"Success. By {GITHUB_REPO}"},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        res = requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=data)
        if res.status_code in [200, 202]:
            print(f"    ✅ 邮件已发送至 {my_email}")
            # 打标：记录到日志中
            log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
            mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
            r_latest = requests.get(log_url, headers=headers)
            requests.put(log_url, headers=headers, data=(r_latest.text + mark).encode('utf-8'))
    except:
        pass

def task_upload_large_file(token):
    print("\n>>> [Task 4] 大文件分片上传 (CreateUploadSession)")
    headers = get_headers(token)
    
    size_mb = random.randint(1, 50)
    file_size = size_mb * 1024 * 1024
    print(f"    📄 准备生成文件: {size_mb} MB")
    file_name = f"Auto_{int(time.time())}_{random_string(4)}.bin"
    
    try:
        session_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/createUploadSession'
        session_data = {"item": {"@microsoft.graph.conflictBehavior": "rename"}}
        r_session = requests.post(session_url, headers=headers, json=session_data)
        if r_session.status_code != 200:
            print(f"    ❌ 创建会话失败: {r_session.text}")
            return
            
        upload_url = r_session.json()['uploadUrl']
        CHUNK_SIZE = 10 * 1024 * 1024
        
        with open("/dev/zero", "rb") as f:
            current_pos = 0
            while current_pos < file_size:
                bytes_left = file_size - current_pos
                this_chunk_size = min(bytes_left, CHUNK_SIZE)
                chunk_data = b'\0' * this_chunk_size
                end_pos = current_pos + this_chunk_size - 1
                headers_chunk = {
                    'Content-Length': str(this_chunk_size),
                    'Content-Range': f'bytes {current_pos}-{end_pos}/{file_size}'
                }
                r_chunk = requests.put(upload_url, headers=headers_chunk, data=chunk_data)
                if r_chunk.status_code not in [200, 201, 202]:
                    print(f"    ❌ 分片上传失败: {r_chunk.status_code}")
                    return
                current_pos += this_chunk_size
                print(f"        -> 已上传: {current_pos / 1024 / 1024:.2f} MB")
        print(f"    ✅ 上传完成: {file_name}")
        
    except Exception as e:
        print(f"    ⚠️ 上传异常: {e}")

    print("\n>>> [Task 5] 文件轮替清理")
    try:
        list_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}:/children?$select=id,name,createdDateTime'
        r_list = requests.get(list_url, headers=headers)
        if r_list.status_code == 200:
            items = r_list.json().get('value', [])
            bins = [x for x in items if x['name'].startswith("Auto_") and x['name'].endswith(".bin")]
            if len(bins) > 25:
                bins.sort(key=lambda x: x['createdDateTime'])
                to_delete = bins[:len(bins)-25]
                for item in to_delete:
                    requests.delete(f'{GRAPH_URL}/me/drive/items/{item["id"]}', headers=headers)
                    print(f"        -> 删除: {item['name']}")
    except:
        pass

# ================= 🏁 主入口 =================
def main():
    token = get_access_token()
    today_str = datetime.now().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H') # 以小时作为时段锁

    # 1. 尝试抢占
    if not try_lock(token, today_str, current_period):
        sys.exit(0) # 没抢到直接停

    # 2. 执行任务
    task_read_calendar(token)
    old_log = task_update_log(token)
    task_send_mail(token, old_log, today_str) # 内部包含频率控制
    task_upload_large_file(token)
    
    print("\n>>> 所有操作成功完成")

if __name__ == '__main__':
    main()
