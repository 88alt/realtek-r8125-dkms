import os
import random
import string
import requests
import time
import json
from datetime import datetime, timezone, timedelta

# ================= 云端配置 =================
# 直接从 GitHub Secrets 读取，无需本地配置
CLIENT_ID = os.environ.get('Z_CLIENT_ID')
CLIENT_SECRET = os.environ.get('Z_CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('Z_REFRESH_TOKEN')

# Graph API 端点
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"  # 你指定的 Data 文件夹

# ================= 辅助函数 =================

def get_access_token():
    """云端获取令牌：用 Refresh Token 换 Access Token"""
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data)
        r.raise_for_status()
        return r.json()['access_token']
    except Exception as e:
        print(f"!! [Cloud Error] 获取令牌失败，请检查 Secrets: {e}")
        exit(1)

def get_me(headers):
    """查询当前跑脚本的账号邮箱"""
    try:
        r = requests.get(f'{GRAPH_URL}/me', headers=headers)
        if r.status_code == 200:
            return r.json().get('userPrincipalName')
    except:
        pass
    return "unknown_user"

def random_str(length=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# ================= 业务逻辑 =================

def task_calendar(headers):
    """【读】读取日历"""
    print(">>> [Task] Reading Calendar...")
    try:
        # 随机读取未来 1-7 天的日程
        requests.get(f'{GRAPH_URL}/me/events?$top={random.randint(1,5)}', headers=headers)
    except:
        pass

def task_mail(headers, email):
    """【发】给自己发邮件"""
    print(f">>> [Task] Sending Mail to {email}...")
    if not email or "unknown" in email: return
    
    subject = f"Cloud Report: {random_str(5)}"
    body = f"Automatic maintenance executed at {datetime.now(timezone.utc)}.\nRandom Seed: {random.random()}"
    
    mail_json = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": email}}]
        },
        "saveToSentItems": "false" # 不污染发件箱
    }
    requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_json)

def task_excel_log(headers):
    """【写】更新 Excel/CSV 日志"""
    print(">>> [Task] Updating Log (CSV)...")
    filename = "ActivityLog.csv"
    # 生成一行日志: Date, Time, Action, ID
    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    new_row = f"\n{now_str},AutoRun,OK,{random.randint(100,999)}"
    
    # 获取文件内容并追加 (模拟真实编辑)
    url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{filename}:/content'
    current_content = ""
    
    # 1. 尝试读取旧文件
    r_get = requests.get(url, headers=headers)
    if r_get.status_code == 200:
        current_content = r_get.text
        # 如果日志太长(超过500行)，裁剪一下，防止文件无限大
        lines = current_content.splitlines()
        if len(lines) > 500:
            current_content = "\n".join(lines[-400:]) # 保留最近400行

    # 2. 覆盖上传新内容
    final_data = current_content + new_row
    requests.put(url, headers=headers, data=final_data)

def task_file_rotation(headers):
    """【存】上传随机大文件并轮替清理 (保持25个)"""
    print(">>> [Task] File Rotation...")
    
    # 1. 生成 1MB - 50MB 随机大小的文件
    size_mb = random.randint(1, 50)
    print(f"    - Generating {size_mb} MB data in memory...")
    # 云端内存生成：使用空字节填充，速度快且不占带宽
    file_content = b'\0' * 1024 * 1024 * size_mb
    
    file_name = f"AutoData_{int(time.time())}_{random_str(4)}.bin"
    upload_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/content'
    
    # 2. 上传
    r_put = requests.put(upload_url, headers=headers, data=file_content)
    if r_put.status_code not in [200, 201]:
        print(f"    !! Upload failed: {r_put.text}")
        return

    # 3. 清理旧文件 (Files.ReadWrite 权限的高频操作)
    print("    - Checking file count...")
    list_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}:/children?$select=id,name,createdDateTime'
    r_list = requests.get(list_url, headers=headers)
    
    if r_list.status_code == 200:
        items = r_list.json().get('value', [])
        # 只筛选我们要管理的 .bin 文件
        target_files = [x for x in items if x['name'].startswith("AutoData_") and x['name'].endswith(".bin")]
        
        count = len(target_files)
        print(f"    - Current files: {count}")
        
        if count > 25:
            # 按时间排序：旧的在前
            target_files.sort(key=lambda x: x['createdDateTime'])
            delete_count = count - 25
            print(f"    - Deleting {delete_count} old files...")
            
            for i in range(delete_count):
                file_id = target_files[i]['id']
                requests.delete(f'{GRAPH_URL}/me/drive/items/{file_id}', headers=headers)
                # 稍微停顿一下，避免并发太快
                time.sleep(1)

# ================= 主流程 =================

def main():
    # 随机启动延迟检查 (双重保险，虽然YAML里也有延迟)
    # 这里不做长时间休眠，主要逻辑交给 YAML 调度
    
    token = get_access_token()
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    me = get_me(headers)
    print(f"User: {me}")
    
    # 执行全套动作 (顺序打乱，更像人)
    tasks = [
        lambda: task_calendar(headers),
        lambda: task_excel_log(headers),
        lambda: task_mail(headers, me),
        lambda: task_file_rotation(headers)
    ]
    random.shuffle(tasks)
    
    for t in tasks:
        t()
        # 动作之间随机休息 5-20 秒
        time.sleep(random.randint(5, 20))

if __name__ == '__main__':
    main()
