import os
import pandas as pd
import sys
import base64
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import logging
from logging.handlers import TimedRotatingFileHandler
import traceback
import time
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
handler = TimedRotatingFileHandler('app.log', when='midnight', interval=1, backupCount=30) # 產生名稱為app的log檔案，並且會每日會備份且建立新的一份，最多存30天
handler.setLevel(logging.INFO)  # 層級比較重要的執行過程才寫入app.log中，可以看看logging套件的官網說明
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')  # 設置 handler 的格式
handler.setFormatter(formatter)
logger.addHandler(handler) 

def custom_exception_handler(exc_type, exc_value, exc_trackback, messageId=None, logger = None):
    ''' 
    將問題記錄下來比較好回查，此程式可以將except的內容搭配logging記錄下來並且重複使用。
    '''
    try:
        if logger is None:
            logger = logging.getLogger(__name__)
        if issubclass(exc_type, PermissionError):
            logger.error(f"PermissionError: {exc_value} - 請確認目前user是否有寫入權限.")
        else:
            logger.error(f"Exception: {exc_type.__name__} - {exc_value}")
            stack_trace = "Stack trace:" + ''.join(traceback.format_tb(exc_trackback))
            logger.error(stack_trace)
    except Exception as e:
        logger.error(f"Error in custom_exception_handler: {e}, wrong Message id: {messageId}")
    

def format_seconds(seconds):
    '''
    將執行過程的秒數換算成幾分幾秒的個格式
    '''
    minutes, seconds = divmod(seconds, 60)
    return "{:02d}:{:02d}".format(int(minutes), int(seconds))

q_times = input("請輸入檔案處理的筆數(預設為10): ") or 10
# 計算此次要處理的檔案數量

before = time.perf_counter()
# 計算花了多少時間跑完程式

SCOPES = ['https://mail.google.com/']

creds = None

# 定義文件路徑
mail_token = Path("./token.pickle")

def create_new_token(json_path):
    json_file = Path(json_path)
    flow = InstalledAppFlow.from_client_secrets_file(json_file, SCOPES)
    creds = flow.run_local_server(port=8081)
    with mail_token.open('wb') as token:
        pickle.dump(creds, token)

# gmail credential JSON 文件的路徑
user_specified_json_path = "./gmail_cred.json"
create_new_token(user_specified_json_path)

try:
    if mail_token.exists():
        with mail_token.open('rb') as token:
            creds = pickle.load(token)
            logger.info(f'token 到期? {creds.expired}')

        if creds.expired:
            try:
                creds.refresh(Request())
            except Exception as e:
                logger.error("Token需要重製，可能已更改信箱密碼，需要登入再次驗證")
                create_new_token()
                logger.info(f"已更新 {mail_token}")
                with mail_token.open('rb') as token:
                    creds = pickle.load(token)
    else:
        logger.info(f"已建立 {mail_token}")
        create_new_token()
        logger.info(f"已建立 {mail_token}")
        with mail_token.open('rb') as token:
            creds = pickle.load(token)

    logger.info(f'token 有效日期: {creds.expiry + pd.Timedelta(hours=8)}')

except Exception as e:
    custom_exception_handler(type(e), e, e.__traceback__, logger)

def start_service(creds):
    service = build('gmail', 'v1', credentials=creds, cache_discovery=False)
    #https://stackoverflow.com/questions/55561354/modulenotfounderror-no-module-named-google-appengine
    return service

service = start_service(creds)
logger.info(f"服務成功啟動")

def create_folder(folder_name):
    """
    創造資料夾。
    """
    folder_path = Path(folder_name)
    
    if folder_path.exists():
        logger.info(f'已經存在資料夾 {folder_name}---繼續執行')
    else:
        folder_path.mkdir()
        logger.info(f'不存在資料夾 {folder_name}，建立成功---繼續執行')

create_folder('einvoice')



def GetAttachments(service, user_id, msg_id, prefix="", en_name=""):
    """Get and store attachment from Message with given id.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: ID of Message containing attachment.
    prefix: prefix which is added to the attachment filename on saving
    """
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()

    for part in message['payload']['parts']:
        if part['filename']:
            if 'data' in part['body']:
                data = part['body']['data']
            else:
                att_id = part['body']['attachmentId']
                att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
                data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            path = prefix + msg_id + '_' + part['filename']
            logger.info(f"目前處理的物件ID是{msg_id},檔案名稱是:{part['filename']}")
            with open(path, 'wb') as f:
                f.write(file_data)



def record_done(messageId):
    with open('./already_parsed_mails.txt', 'a') as f: 
        f.write(messageId + "\n")

def query_sub(subject):
    ''' 搜尋信件主題，並下載至指定資料夾'''
    query = f"(subject:{subject})+ (filename:pdf OR filename:xls OR filename:xlsx OR filename:csv) " 
    logger.info(f"目前正在下載:{subject}")
    # q_times = 1000
    results = service.users().messages().list(userId = 'me', q = query, maxResults=q_times).execute()
    msgs = results['messages']
    msg_ids = [msg['id'] for msg in msgs]
    logger.info(f"{subject}的近{q_times}筆檔案獨立ID列表:{msg_ids}")

    if not os.path.exists('./already_parsed_mails.txt'):
        with open('./already_parsed_mails.txt', 'a'):pass
        
    if os.path.exists('./already_parsed_mails.txt'):
        with open('./already_parsed_mails.txt', 'r') as f: 
            list_done = f.read().split()
# msg_tids = [msg['threadId'] for msg in msgs]
    for messageId in msg_ids:
        try:

            if messageId not in list_done:
                if subject == '財政部電子發票整合服務平台':
                    GetAttachments(service, "me", messageId, prefix="./einvoice/")
                    record_done(messageId)

        except Exception as e:
            custom_exception_handler(type(e), e, e.__traceback__, messageId, logger)

sub = ['財政部電子發票整合服務平台']

[query_sub(s) for s in sub]

download_end = time.perf_counter()
download_spent = download_end - before 
logger.info(f'Download總共花費:{format_seconds(download_spent)}')

logger.info(f'最新{(sub)}的{q_times}筆檔案下載完畢')