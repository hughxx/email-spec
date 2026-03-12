import logging
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime
from typing import Optional
from dataclasses import dataclass


def convert_pywin_datetime(pywin_dt: any) -> datetime:
    """将 pywintypes.DateTime 转换为 Python datetime"""
    # pywin32 返回的是 pywintypes.DateTime 对象
    # 它有 year, month, day, hour, minute, second 属性
    try:
        return datetime(
            pywin_dt.year,
            pywin_dt.month,
            pywin_dt.day,
            pywin_dt.hour,
            pywin_dt.minute,
            pywin_dt.second
        )
    except Exception:
        # 如果转换失败，尝试直接返回
        return pywin_dt

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


@dataclass
class EmailItem:
    entry_id: str
    subject: str
    conversation_topic: str
    sent_on: datetime
    sender: str
    body: str
    html_body: str
    attachments: list


class OutlookClient:
    def __init__(self) -> None:
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        logger.info("Outlook client initialized")

    def get_folder_tree(self) -> dict[str, str]:
        """获取文件夹树形结构"""
        folders = {}
        root = self.namespace.Folders.Item(1)
        self._collect_folders(root, folders, "")
        logger.info(f"Found {len(folders)} folders")
        return folders

    def _collect_folders(self, folder, result: dict, path: str) -> None:
        """递归收集文件夹"""
        full_path = f"{path}/{folder.Name}" if path else folder.Name
        result[full_path] = folder.Name
        try:
            for subfolder in folder.Folders:
                self._collect_folders(subfolder, result, full_path)
        except Exception as e:
            logger.warning(f"Error collecting folders: {e}")

    def _find_folder_by_path(self, root_folder, path_parts: list) -> Optional[any]:
        """递归查找嵌套文件夹"""
        if not path_parts:
            return root_folder

        folder_name = path_parts[0]
        try:
            for subfolder in root_folder.Folders:
                if subfolder.Name == folder_name:
                    if len(path_parts) == 1:
                        return subfolder
                    else:
                        return self._find_folder_by_path(subfolder, path_parts[1:])
        except Exception as e:
            logger.warning(f"Error finding folder {folder_name}: {e}")
        return None

    def get_folder(self, folder_path: str) -> Optional[any]:
        """根据路径获取文件夹"""
        if not folder_path or folder_path == "收件箱":
            return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        path_parts = folder_path.split("/")
        root = self.namespace.Folders.Item(1)
        return self._find_folder_by_path(root, path_parts)

    def get_default_folder(self) -> any:
        """获取默认收件箱"""
        return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

    def get_emails(
        self,
        folder_path: str,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
        keyword: Optional[str] = None
    ) -> list[EmailItem]:
        """获取邮件列表"""
        folder = self.get_folder(folder_path)
        if not folder:
            logger.error(f"Folder not found: {folder_path}")
            return []

        items = folder.Items
        items.Sort("[SentOn]", True)

        emails = []
        seen_topics = set()

        for item in items:
            if item.Class != 43:  # 43 = olMail
                continue

            # 日期筛选 - 转换 pywintypes.datetime 为 Python datetime
            sent_on = convert_pywin_datetime(item.SentOn)
            if start_date and sent_on.date() < start_date:
                continue
            if end_date:
                end_day = end_date.replace(hour=23, minute=59, second=59)
                if sent_on.date() > end_day:
                    continue

            # 关键词筛选
            if keyword and keyword.lower() not in (item.Subject or "").lower():
                continue

            # 去重：同一 ConversationTopic 只保留最新
            topic = item.ConversationTopic
            if topic in seen_topics:
                continue
            seen_topics.add(topic)

            email = EmailItem(
                entry_id=item.EntryID,
                subject=item.Subject,
                conversation_topic=topic,
                sent_on=sent_on,
                sender=item.SenderName,
                body=item.Body,
                html_body=item.HTMLBody,
                attachments=[]
            )
            emails.append(email)

        logger.info(f"Retrieved {len(emails)} emails from {folder_path}")
        return emails
