import logging
import requests
from typing import Optional

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

BASE_URL: str = "http://localhost:8000"
TIMEOUT: int = 30


class APIClient:
    def __init__(self, base_url: str = BASE_URL) -> None:
        self.base_url = base_url
        self.session = requests.Session()
        self.session.timeout = TIMEOUT

    def upload_word(self, file_path: str) -> dict:
        """上传 Word 文件，返回 task_id"""
        logger.info(f"Uploading {file_path} to {self.base_url}")
        with open(file_path, "rb") as f:
            files = {
                "file": (
                    file_path,
                    f,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            }
            response = self.session.post(f"{self.base_url}/upload", files=files)
            response.raise_for_status()
            result = response.json()
            logger.info(f"Upload successful, task_id: {result.get('task_id')}")
            return result

    def get_task_status(self, task_id: str) -> Optional[dict]:
        """查询任务状态"""
        try:
            response = self.session.get(f"{self.base_url}/task/{task_id}")
            if response.status_code == 404:
                logger.warning(f"Task {task_id} not found")
                return None
            response.raise_for_status()
            result = response.json()
            logger.info(f"Task {task_id} status: {result.get('status')}")
            return result
        except requests.RequestException as e:
            logger.error(f"Failed to get task status: {e}")
            return None
