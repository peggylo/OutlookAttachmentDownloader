import win32com.client
import os

# 設定桌面路徑並創建 file 資料夾
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
file_folder_path = os.path.join(desktop_path, 'file')
os.makedirs(file_folder_path, exist_ok=True)

# 連接到 Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 選擇收件匣
inbox = outlook.GetDefaultFolder(6)  # 6 代表收件匣
messages = inbox.Items

# 按接收時間排序，最新的在前
messages.Sort("[ReceivedTime]", True)

# 取得最新的第8封郵件
message = messages[7]  # 索引從0開始，所以第8封郵件是索引7

# 打印第8封郵件的標題和附件資訊
print(f"第 8 封郵件的標題: {message.Subject}")
attachment_count = message.Attachments.Count
print(f"是否有附件: {'有' if attachment_count > 0 else '無'}")
if attachment_count > 0:
    for attachment in message.Attachments:
        attachment_path = os.path.join(file_folder_path, attachment.FileName)
        attachment.SaveAsFile(attachment_path)
        print(f"附件 {attachment.FileName} 已下載到 {attachment_path}")
print("-" * 40)
