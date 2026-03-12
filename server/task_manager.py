import logging
import json
import os
import uuid
import threading
import time
from typing import Optional

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

TASK_FILE: str = "server/storage/tasks.json"


class TaskManager:
    def __init__(self) -> None:
        self.tasks: dict[str, dict] = {}
        self._lock = threading.Lock()
        self._load_tasks()

    def _load_tasks(self) -> None:
        """从磁盘加载任务状态"""
        if os.path.exists(TASK_FILE):
            with open(TASK_FILE, "r", encoding="utf-8") as f:
                self.tasks = json.load(f)
            logger.info(f"Loaded {len(self.tasks)} tasks from disk")

    def _save_tasks(self) -> None:
        """持久化任务状态到磁盘"""
        os.makedirs(os.path.dirname(TASK_FILE), exist_ok=True)
        with open(TASK_FILE, "w", encoding="utf-8") as f:
            json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        logger.info(f"Tasks saved to disk")

    def create_task(self) -> str:
        """创建新任务，返回 task_id"""
        task_id = str(uuid.uuid4())
        with self._lock:
            self.tasks[task_id] = {
                "status": "pending",
                "progress": 0
            }
            self._save_tasks()
        logger.info(f"Task {task_id} created")
        # 后台异步处理
        threading.Thread(target=self._process_task, args=(task_id,)).start()
        return task_id

    def _process_task(self, task_id: str) -> None:
        """模拟处理任务"""
        logger.info(f"Task {task_id} processing...")
        with self._lock:
            self.tasks[task_id]["status"] = "processing"
            self._save_tasks()

        # mock 处理：延迟约 3 秒
        time.sleep(3)

        with self._lock:
            self.tasks[task_id]["status"] = "completed"
            self.tasks[task_id]["progress"] = 100
            self._save_tasks()
        logger.info(f"Task {task_id} completed")

    def get_status(self, task_id: str) -> Optional[dict]:
        """获取任务状态"""
        with self._lock:
            task = self.tasks.get(task_id)
            if not task:
                return None
            return {
                "task_id": task_id,
                "status": task["status"],
                "progress": task.get("progress", 0)
            }
