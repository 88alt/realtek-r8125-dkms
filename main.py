import os
import random
import string
import requests
import time
from datetime import datetime

# ================= 配置区域 =================
# 从 Secrets 读取变量
CLIENT_ID = os.environ.get('Z_CLIENT_ID')
CLIENT_SECRET = os.environ.get('Z_CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('Z_REFRESH_TOKEN')

# API 端点
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"  # 对应你的 .../Documents/Data

# ================= 核心功能函数 =================

def get_access_token():
    """使用 Refresh Token 换取 Access Token"""
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
        r.raise_for_status()
        return r.json()['access_token']
    except Exception as e:
        print(f"!! 获取令牌失败: {e}")
        print("请检查 GitHub Secrets 中的 Z_REFRESH_TOKEN 是否正确。")
        exit(1)

def get_me(headers):
    """获取当前账号邮箱"""
    try:
        r = requests.get(f'{GRAPH_URL}/me', headers=headers)
        if r.status_code == 200:
            return r.json().get('userPrincipalName')
    except:
        pass
    return None

def random_string(length=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# --- 任务 1: 读日历 ---
def task_read_calendar(headers):
    print("\n>>> [Task 1] 读取日历 (Calendar.Read)")
    try:
        # 随机读取未来 1-3 天的日程
        r = requests.get(f'{GRAPH_URL}/me/events?$top={random.randint(1,3)}', headers=headers)
        if r.status_code == 200:
            print("    成功读取日程数据。")
    except Exception as e:
        print(f"    任务执行异常: {e}")

# --- 任务 2: 写 Excel/CSV 日志 ---
def task_update_log(headers):
    print("\n>>> [Task 2] 更新日志表格 (Files.ReadWrite)")
    # 使用 CSV 代替 XLSX，效果一样且更稳定
    filename = "ActivityLog.csv"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # CSV 格式: Time, Type, Status, RandomID
    new_row = f"\n{timestamp},AutoRun,Success,{random.randint(1000,9999)}"
    
    file_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{filename}:/content'
    
    try:
        # 1. 尝试读取旧内容
        current_content = "Time,Type,Status,ID"
        r_get = requests.get(file_url, headers=headers)
        if r_get.status_code == 200:
            current_content = r_get.text
            # 防止日志无限增长，只保留最后 100 行
            lines = current_content.splitlines()
            if len(lines) > 100:
                current_content = "\n".join(lines[-100:])
        
        # 2. 追加并覆盖上传
        final_content = current_content + new_row
        requests.put(file_url, headers=headers, data=final_content)
        print("    日志表格更新完成。")
    except Exception as e:
        print(f"    更新日志失败: {e}")

# --- 任务 3: 发邮件 ---
def task_send_mail(headers, my_email):
    print("\n>>> [Task 3] 发送邮件 (Mail.Send)")
    if not my_email: 
        print("    未获取到邮箱地址，跳过。")
        return

    subject = f"KeepAlive Report: {random_string(5)}"
    body = f"Automatic maintenance executed at {datetime.now()}.\nRandom Seed: {random.random()}"
    
    mail_data = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": my_email}}]
        },
        "saveToSentItems": "false" # 不保存到发件箱
    }
    
    try:
        r = requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data)
        if r.status_code == 202:
            print(f"    邮件已发送至 {my_email}")
        else:
            print(f"    发送失败: {r.text}")
    except Exception as e:
        print(f"    邮件发送异常: {e}")

# --- 任务 4: 文件生成与清理 (核心需求) ---
def task_file_manage(headers):
    print("\n>>> [Task 4] 文件生成与轮替 (Files.ReadWrite)")
    
    # A. 生成随机文件 (1MB - 50MB)
    size_mb = random.randint(1, 50)
    print(f"    - 正在生成 {size_mb} MB 随机数据...")
    
    # 内存生成空数据 (全0)，节省 GitHub Runner 资源但占用 OneDrive 空间
    file_content = b'\0' * 1024 * 1024 * size_mb
    file_name = f"KeepAlive_{int(time.time())}_{random_string(4)}.bin"
    
    upload_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/content'
    
    # B. 上传
    try:
        r_up = requests.put(upload_url, headers=headers, data=file_content)
        if r_up.status_code in [200, 201]:
            print(f"    - 上传成功: {file_name}")
        else:
            print(f"    - 上传失败: {r_up.text}")
            return # 上传失败就不清理了
    except Exception as e:
        print(f"    上传异常: {e}")
        return

    # C. 检查数量并删除旧文件 (保留 25 个)
    print("    - 正在检查文件数量...")
    list_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}:/children?$select=id,name,createdDateTime'
    
    try:
        r_list = requests.get(list_url, headers=headers)
        if r_list.status_code == 200:
            items = r_list.json().get('value', [])
            # 只统计 .bin 文件
            bin_files = [x for x in items if x['name'].startswith("KeepAlive_") and x['name'].endswith(".bin")]
            count = len(bin_files)
            print(f"    - 当前 .bin 文件数: {count}")
            
            if count > 25:
                # 按创建时间排序 (旧的在前)
                bin_files.sort(key=lambda x: x['createdDateTime'])
                delete_count = count - 25
                print(f"    - 需要删除 {delete_count} 个旧文件...")
                
                for i in range(delete_count):
                    item = bin_files[i]
                    print(f"      删除: {item['name']}")
                    del_url = f'{GRAPH_URL}/me/drive/items/{item["id"]}'
                    requests.delete(del_url, headers=headers)
                    time.sleep(1) # 避免 API 速率限制
    except Exception as e:
        print(f"    文件清理异常: {e}")

# ================= 主程序 =================
def main():
    token = get_access_token()
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    my_email = get_me(headers)
    print(f"当前用户: {my_email}")
    
    # 执行所有任务
    task_read_calendar(headers)
    task_update_log(headers)
    task_send_mail(headers, my_email)
    task_file_manage(headers)
    
    print("\n>>> 所有任务执行完毕。")

if __name__ == '__main__':
    main()
