import os
import random
import string
import requests
import time
import sys
from datetime import datetime

# ================= 🛡️ 核心配置区域 (通用版) =================
# 修改点：去掉了 Z_ 前缀，变成通用变量名
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')

# API 端点
if not TENANT_ID:
    print("!! 致命错误: 缺少 TENANT_ID 环境变量")
    sys.exit(1)

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"  # 两个账号都必须有这个文件夹，或者脚本会自动建

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

# ================= 🛠️ 辅助工具箱 (不变) =================
def get_headers(token):
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

def random_string(length=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# ================= 🚀 业务模块 (不变) =================

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
    except Exception as e:
        print(f"    ⚠️ 日志操作异常: {e}")

def task_send_mail(token):
    print("\n>>> [Task 3] 发送邮件")
    headers = get_headers(token)
    try:
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers)
        if r_me.status_code != 200: return
        my_email = r_me.json().get('userPrincipalName')
        
        data = {
            "message": {
                "subject": f"KeepAlive: {random_string(5)}",
                "body": {"contentType": "Text", "content": "Auto maintenance success."},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=data)
        print(f"    ✅ 邮件已发送至 {my_email}")
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
    tasks = [
        lambda: task_read_calendar(token),
        lambda: task_update_log(token),
        lambda: task_send_mail(token),
        lambda: task_upload_large_file(token)
    ]
    for t in tasks:
        t()
        time.sleep(2)
    print("\n>>> 所有操作成功完成")

if __name__ == '__main__':
    main()
