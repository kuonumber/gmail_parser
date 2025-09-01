#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gmail 郵件內容下載工具
支援主旨篩選、時間範圍搜尋、附件下載與郵件內文下載
"""

import os
import base64
import pickle
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Gmail API 相關
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# 環境變數支援
from dotenv import load_dotenv

# 全域變數
TOKEN_FILE = None
GMAIL_CREDENTIALS_DIR = None
GMAIL_CREDENTIAL_PATTERNS = None
GMAIL_SUBJECTS = None
GMAIL_DOWNLOAD_DIR = None
GMAIL_SUBJECT_FOLDER_MAPPING = None
GMAIL_FILE_TYPES = None
GMAIL_DOWNLOAD_CONTENT = None
GMAIL_DATE_RANGE = None
GMAIL_START_DATE = None
GMAIL_END_DATE = None

# Gmail API 權限範圍
SCOPES = ['https://mail.google.com/']

def load_environment_variables():
    """載入環境變數到全域變數"""
    global TOKEN_FILE, GMAIL_CREDENTIALS_DIR, GMAIL_CREDENTIAL_PATTERNS
    global GMAIL_SUBJECTS, GMAIL_DOWNLOAD_DIR, GMAIL_SUBJECT_FOLDER_MAPPING
    global GMAIL_FILE_TYPES, GMAIL_DOWNLOAD_CONTENT
    global GMAIL_DATE_RANGE, GMAIL_START_DATE, GMAIL_END_DATE
    
    # 基本設定
    TOKEN_FILE = os.getenv('GMAIL_TOKEN_FILE', './token.pickle')
    GMAIL_CREDENTIALS_DIR = os.getenv('GMAIL_CREDENTIALS_DIR', './')
    GMAIL_CREDENTIAL_PATTERNS = os.getenv('GMAIL_CREDENTIAL_PATTERNS', 'gmail_cred.json,mp_gmail.json,credentials.json,client_secret.json')
    GMAIL_SUBJECTS = os.getenv('GMAIL_SUBJECTS', '')
    GMAIL_DOWNLOAD_DIR = os.getenv('GMAIL_DOWNLOAD_DIR', './downloads')
    GMAIL_SUBJECT_FOLDER_MAPPING = os.getenv('GMAIL_SUBJECT_FOLDER_MAPPING', '')
    GMAIL_FILE_TYPES = os.getenv('GMAIL_FILE_TYPES', 'pdf,xls,xlsx,csv')
    GMAIL_DOWNLOAD_CONTENT = os.getenv('GMAIL_DOWNLOAD_CONTENT', 'true').lower() == 'true'
    
    # 時間範圍設定
    GMAIL_DATE_RANGE = os.getenv('GMAIL_DATE_RANGE', '')
    GMAIL_START_DATE = os.getenv('GMAIL_START_DATE', '')
    GMAIL_END_DATE = os.getenv('GMAIL_END_DATE', '')

def get_credentials():
    """取得 Gmail API 認證"""
    creds = None
    
    # 檢查是否有快取的 token
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    # 如果沒有有效的認證，則重新認證
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # 尋找認證檔案
            credential_patterns = [p.strip() for p in GMAIL_CREDENTIALS_DIR.split(',')]
            credential_file = None
            
            for pattern in credential_patterns:
                if pattern.startswith('./'):
                    pattern = pattern[2:]
                if pattern.startswith('/'):
                    pattern = pattern[1:]
                
                # 搜尋認證檔案
                for file in os.listdir(GMAIL_CREDENTIALS_DIR):
                    if pattern in file and file.endswith('.json'):
                        credential_file = os.path.join(GMAIL_CREDENTIALS_DIR, file)
                        break
                
                if credential_file:
                    break
            
            if not credential_file:
                raise FileNotFoundError(f"找不到認證檔案，請確認 {GMAIL_CREDENTIALS_DIR} 目錄中有符合 {GMAIL_CREDENTIALS_DIR} 模式的 JSON 檔案")
            
            flow = InstalledAppFlow.from_client_secrets_file(credential_file, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # 儲存認證
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def build_date_query():
    """建立日期查詢條件"""
    if GMAIL_DATE_RANGE:
        now = datetime.now()
        
        if GMAIL_DATE_RANGE == 'today':
            start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        elif GMAIL_DATE_RANGE == 'yesterday':
            start_date = (now - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        elif GMAIL_DATE_RANGE == 'week':
            start_date = (now - timedelta(days=7))
        elif GMAIL_DATE_RANGE == 'month':
            start_date = (now - timedelta(days=30))
        elif GMAIL_DATE_RANGE == 'year':
            start_date = (now - timedelta(days=365))
        else:
            # 嘗試解析相對天數 (例如: 7d, 14d)
            match = re.match(r'(\d+)d', GMAIL_DATE_RANGE)
            if match:
                days = int(match.group(1))
                start_date = (now - timedelta(days=days))
            else:
                start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        
        end_date = now
    elif GMAIL_START_DATE and GMAIL_END_DATE:
        try:
            start_date = datetime.strptime(GMAIL_START_DATE, '%Y/%m/%d')
            end_date = datetime.strptime(GMAIL_END_DATE, '%Y/%m/%d')
        except ValueError:
            print(f"日期格式錯誤，請使用 YYYY/MM/DD 格式")
            return ""
    else:
        return ""
    
    # 轉換為 RFC 3339 格式
    start_str = start_date.strftime('%Y/%m/%d')
    end_str = end_date.strftime('%Y/%m/%d')
    
    return f"after:{start_str} before:{end_str}"

def get_download_prefix(subject: str, message_date: str = None) -> str:
    """根據主旨取得下載目錄前綴"""
    if GMAIL_SUBJECT_FOLDER_MAPPING:
        mappings = [m.strip() for m in GMAIL_SUBJECT_FOLDER_MAPPING.split(',')]
        for mapping in mappings:
            if ':' in mapping:
                key, folder = mapping.split(':', 1)
                if key.strip().lower() in subject.lower():
                    return folder.strip()
    
    # 如果沒有主旨對應映射，存入 all/日期 資料夾
    if message_date:
        try:
            # 解析日期格式 (例如: "Tue, 25 Feb 2025 15:20:37 +0800")
            date_obj = datetime.strptime(message_date.split('+')[0].strip(), '%a, %d %b %Y %H:%M:%S')
            date_folder = date_obj.strftime('%Y-%m-%d')
            return f"all/{date_folder}"
        except:
            # 如果日期解析失敗，使用 today 作為備用
            today = datetime.now().strftime('%Y-%m-%d')
            return f"all/{today}"
    
    # 如果沒有日期資訊，使用今天的日期
    today = datetime.now().strftime('%Y-%m-%d')
    return f"all/{today}"

def query_sub(service, subject_keywords: List[str], date_query: str = "") -> List[Dict]:
    """根據主旨關鍵字查詢郵件"""
    messages = []
    
    # 如果沒有主旨關鍵字，只搜尋時間範圍
    if not subject_keywords or not any(subject_keywords):
        query = date_query if date_query else ""
        try:
            results = service.users().messages().list(userId='me', q=query).execute()
            
            if 'messages' not in results or not results['messages']:
                print(f"指定時間範圍的郵件: 0 封")
                return []
            
            print(f"指定時間範圍的郵件: {len(results['messages'])} 封")
            messages.extend(results['messages'])
            
        except HttpError as error:
            print(f"查詢郵件時發生錯誤: {error}")
            return []
        
        return messages
    
    # 有主旨關鍵字時的搜尋邏輯
    for keyword in subject_keywords:
        if not keyword.strip():  # 跳過空字串
            continue
            
        query = f"subject:{keyword}"
        if date_query:
            query += f" {date_query}"
        
        try:
            results = service.users().messages().list(userId='me', q=query).execute()
            
            if 'messages' not in results or not results['messages']:
                print(f"主旨包含 '{keyword}' 的郵件: 0 封")
                continue
            
            print(f"主旨包含 '{keyword}' 的郵件: {len(results['messages'])} 封")
            messages.extend(results['messages'])
            
        except HttpError as error:
            print(f"查詢主旨 '{keyword}' 時發生錯誤: {error}")
    
    return messages

def download_attachment(service, message_id: str, attachment_id: str, filename: str, target_path: Path):
    """下載附件"""
    try:
        attachment = service.users().messages().attachments().get(
            userId='me', messageId=message_id, id=attachment_id
        ).execute()
        
        if 'data' in attachment:
            file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
            
            # 確保目標目錄存在
            target_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(target_path, 'wb') as f:
                f.write(file_data)
            
            print(f"  附件已下載: {target_path.name}")
            return True
        else:
            print(f"  附件資料無效: {filename}")
            return False
            
    except HttpError as error:
        print(f"  下載附件失敗: {filename}, 錯誤: {error}")
        return False

def download_email_content(service, message_id: str, target_path: Path):
    """下載郵件內文"""
    try:
        message = service.users().messages().get(userId='me', id=message_id, format='full').execute()
        
        content = ""
        headers = message['payload'].get('headers', [])
        
        # 提取郵件資訊
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '無主旨')
        sender = next((h['value'] for h in headers if h['name'] == 'From'), '無寄件者')
        date = next((h['value'] for h in headers if h['name'] == 'Date'), '無日期')
        
        content += f"主題: {subject}\n"
        content += f"寄件者: {sender}\n"
        content += f"日期: {date}\n"
        content += f"郵件ID: {message_id}\n"
        content += "-" * 50 + "\n\n"
        
        # 解析郵件內容
        def extract_text(part):
            if part.get('mimeType') == 'text/plain':
                data = part['body'].get('data', '')
                if data:
                    return base64.urlsafe_b64decode(data.encode('UTF-8')).decode('UTF-8', errors='ignore')
            elif part.get('mimeType') == 'text/html':
                data = part['body'].get('data', '')
                if data:
                    html_content = base64.urlsafe_b64decode(data.encode('UTF-8')).decode('UTF-8', errors='ignore')
                    # 簡單的 HTML 標籤移除
                    text_content = re.sub(r'<[^>]+>', '', html_content)
                    text_content = re.sub(r'\s+', ' ', text_content).strip()
                    return text_content
            return ""
        
        def process_part(part):
            text = ""
            if 'parts' in part:
                for subpart in part['parts']:
                    text += process_part(subpart)
            else:
                text = extract_text(part)
            return text
        
        if 'payload' in message:
            text_content = process_part(message['payload'])
            if text_content:
                content += text_content
            else:
                content += "無法解析郵件內容"
        
        # 確保目標目錄存在
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(target_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"  郵件內文已下載: {target_path.name}")
        return True
        
    except HttpError as error:
        print(f"  下載郵件內文失敗: {message_id}, 錯誤: {error}")
        return False

def get_attachments(service, message_id: str, message_data: Dict, download_dir: Path):
    """取得並下載附件"""
    attachments_downloaded = 0
    
    if 'parts' in message_data['payload']:
        for part in message_data['payload']['parts']:
            if part.get('filename'):
                filename = part['filename']
                attachment_id = part['body'].get('attachmentId')
                
                if attachment_id:
                    # 檢查檔案類型
                    file_ext = Path(filename).suffix.lower()
                    allowed_types = [f".{t.strip()}" for t in GMAIL_FILE_TYPES.split(',')]
                    
                    if file_ext in allowed_types:
                        target_path = download_dir / f"{message_id}_{filename}"
                        if download_attachment(service, message_id, attachment_id, filename, target_path):
                            attachments_downloaded += 1
                    else:
                        print(f"  跳過不支援的檔案類型: {filename} ({file_ext})")
    
    elif 'filename' in message_data['payload'] and message_data['payload']['filename']:
        filename = message_data['payload']['filename']
        attachment_id = message_data['payload']['body'].get('attachmentId')
        
        if attachment_id:
            file_ext = Path(filename).suffix.lower()
            allowed_types = [f".{t.strip()}" for t in GMAIL_FILE_TYPES.split(',')]
            
            if file_ext in allowed_types:
                target_path = download_dir / f"{message_id}_{filename}"
                if download_attachment(service, message_id, attachment_id, filename, target_path):
                    attachments_downloaded += 1
            else:
                print(f"  跳過不支援的檔案類型: {filename} ({file_ext})")
    
    return attachments_downloaded

def main():
    """主程式"""
    print("Gmail 郵件內容下載工具")
    print("=" * 50)
    
    # 載入環境變數
    config_file = os.getenv('GMAIL_CONFIG_FILE')
    if config_file:
        print(f"載入組態檔: {config_file}")
        load_dotenv(config_file, override=True)
    else:
        print("載入預設組態")
        load_dotenv()
    
    load_environment_variables()
    
    # 顯示組態
    print(f"主旨關鍵字: {GMAIL_SUBJECTS}")
    print(f"檔案類型: {GMAIL_FILE_TYPES}")
    print(f"下載目錄: {GMAIL_DOWNLOAD_DIR}")
    print(f"下載內文: {GMAIL_DOWNLOAD_CONTENT}")
    if GMAIL_DATE_RANGE:
        print(f"時間範圍: {GMAIL_DATE_RANGE}")
    elif GMAIL_START_DATE and GMAIL_END_DATE:
        print(f"自訂日期: {GMAIL_START_DATE} 到 {GMAIL_END_DATE}")
    print("-" * 50)
    
    try:
        # 取得認證
        print("正在取得 Gmail API 認證...")
        creds = get_credentials()
        
        # 建立服務
        service = build('gmail', 'v1', credentials=creds)
        print("Gmail API 服務已建立")
        
        # 建立日期查詢
        date_query = build_date_query()
        if date_query:
            print(f"日期查詢: {date_query}")
        
        # 解析主旨關鍵字
        subject_keywords = [kw.strip() for kw in GMAIL_SUBJECTS.split(',')]
        
        # 查詢郵件
        print("\n正在查詢郵件...")
        messages = query_sub(service, subject_keywords, date_query)
        
        if not messages:
            print("沒有找到符合條件的郵件")
            return
        
        # 去重
        unique_messages = {msg['id']: msg for msg in messages}.values()
        print(f"找到 {len(unique_messages)} 封唯一郵件")
        
        # 詢問處理數量
        q_times_input = input(f"\n請輸入檔案處理的筆數(預設為{min(10, len(unique_messages))}): ")
        q_times = int(q_times_input) if q_times_input.strip() else min(10, len(unique_messages))
        q_times = min(q_times, len(unique_messages))
        
        print(f"將處理 {q_times} 封郵件")
        
        # 檢查已處理的郵件
        already_parsed_file = 'already_parsed_mails.txt'
        already_parsed = set()
        if os.path.exists(already_parsed_file):
            with open(already_parsed_file, 'r', encoding='utf-8') as f:
                already_parsed = set(line.strip() for line in f)
        
        # 處理郵件
        processed_count = 0
        total_attachments = 0
        total_content = 0
        
        for message in list(unique_messages)[:q_times]:
            message_id = message['id']
            
            if message_id in already_parsed:
                print(f"跳過已處理的郵件: {message_id}")
                continue
            
            print(f"\n處理郵件 {processed_count + 1}/{q_times}: {message_id}")
            
            try:
                # 取得郵件詳細資料
                message_data = service.users().messages().get(userId='me', id=message_id, format='full').execute()
                
                # 取得主旨和日期
                headers = message_data['payload'].get('headers', [])
                subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '無主旨')
                message_date = next((h['value'] for h in headers if h['name'] == 'Date'), '')
                print(f"  主旨: {subject}")
                print(f"  日期: {message_date}")
                
                # 建立下載目錄
                download_prefix = get_download_prefix(subject, message_date)
                download_dir = Path(GMAIL_DOWNLOAD_DIR) / download_prefix
                download_dir.mkdir(parents=True, exist_ok=True)
                
                # 下載附件
                attachments = get_attachments(service, message_id, message_data, download_dir)
                total_attachments += attachments
                
                # 下載郵件內文
                if GMAIL_DOWNLOAD_CONTENT:
                    content_path = download_dir / f"{message_id}_content.txt"
                    if download_email_content(service, message_id, content_path):
                        total_content += 1
                
                # 記錄已處理
                with open(already_parsed_file, 'a', encoding='utf-8') as f:
                    f.write(f"{message_id}\n")
                already_parsed.add(message_id)
                
                processed_count += 1
                
            except HttpError as error:
                print(f"  處理郵件時發生錯誤: {error}")
                continue
        
        print(f"\n處理完成！")
        print(f"處理郵件: {processed_count} 封")
        print(f"下載附件: {total_attachments} 個")
        print(f"下載內文: {total_content} 個")
        
    except Exception as e:
        print(f"發生錯誤: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
