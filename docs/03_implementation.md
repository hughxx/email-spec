# 03_implementation.md - 实现清单

## 1. 项目文件结构

```
spec-email/
├── docs/
│   ├── 01_requirements.md
│   └── 02_interface.md
├── client/
│   ├── main.py                 # [修改] 主界面
│   ├── email_window.py         # [新增] 邮件提取窗口
│   ├── outlook_client.py       # [新增] Outlook 操作
│   ├── word_generator.py       # [新增] Word 生成
│   └── api_client.py           # [新增] 后端 API 调用
└── server/
    ├── main.py                 # [新增] FastAPI 入口
    ├── task_manager.py         # [新增] 任务管理
    └── storage/                # [目录] 任务状态持久化
        └── tasks.json          # [自动生成]
```

---

## 2. 前端实现

### 2.1 client/main.py - 主界面

**修改内容**：
- 添加主窗口，显示两个按钮

**函数签名**：
```python
class MainWindow(QMainWindow):
    def __init__(self) -> None: ...
    def _open_email_window(self) -> None: ...
```

---

### 2.2 client/outlook_client.py - Outlook 操作

**新增文件**

**类和函数签名**：
```python
import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

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
    def __init__(self) -> None: ...
    def get_folder_tree(self) -> dict[str, str]: ...
    def get_folder(self, folder_path: str) -> Folder: ...
    def get_emails(
        self,
        folder_path: str,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
        keyword: Optional[str] = None
    ) -> list[EmailItem]: ...
```

---

### 2.3 client/word_generator.py - Word 生成

**新增文件**

**类和函数签名**：
```python
import logging
from typing import Tuple

class WordGenerator:
    @staticmethod
    def clean_filename(name: str) -> str: ...

    @staticmethod
    def extract_images_from_html(html_body: str) -> list[Tuple[str, str]]: ...

    @staticmethod
    def generate_word(email: EmailItem, output_dir: str) -> str: ...
```

---

### 2.4 client/api_client.py - 后端 API 调用

**新增文件**

**类和函数签名**：
```python
import logging
import requests
from typing import Optional

BASE_URL: str = "http://localhost:8000"
TIMEOUT: int = 30

class APIClient:
    def __init__(self, base_url: str = BASE_URL) -> None: ...
    def upload_word(self, file_path: str) -> dict: ...
    def get_task_status(self, task_id: str) -> Optional[dict]: ...
```

---

### 2.5 client/email_window.py - 邮件提取窗口

**新增文件**

**类和函数签名**：
```python
import logging
import os
import tempfile
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem,
    QListWidget, QListWidgetItem, QPushButton, QDateEdit, QLabel,
    QLineEdit, QProgressBar, QMessageBox
)
from PyQt5.QtCore import Qt, QDate

class EmailWindow(QDialog):
    def __init__(self) -> None: ...
    def _init_ui(self) -> None: ...
    def _load_folders(self) -> None: ...
    def _on_folder_selected(self, item: QTreeWidgetItem, column: int) -> None: ...
    def _on_search(self) -> None: ...
    def _on_extract(self) -> None: ...
    def _show_task_status(self) -> None: ...
```

---

## 3. 后端实现

### 3.1 server/task_manager.py - 任务管理

**新增文件**

**常量**：
```python
TASK_FILE: str = "server/storage/tasks.json"
```

**类和函数签名**：
```python
import logging
import json
import os
import uuid
import threading
import time
from typing import Optional

class TaskManager:
    def __init__(self) -> None: ...
    def _load_tasks(self) -> None: ...
    def _save_tasks(self) -> None: ...
    def create_task(self) -> str: ...
    def _process_task(self, task_id: str) -> None: ...
    def get_status(self, task_id: str) -> Optional[dict]: ...
```

---

### 3.2 server/main.py - FastAPI 入口

**新增文件**

**类和函数签名**：
```python
import logging
import os
import tempfile
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import JSONResponse

app: FastAPI = FastAPI()
task_manager: TaskManager = TaskManager()

@app.post("/upload")
async def upload_file(file: UploadFile) -> JSONResponse: ...

@app.get("/task/{task_id}")
def get_task_status(task_id: str) -> dict: ...
```

---

## 4. 启动方式

### 4.1 启动后端

```bash
cd server
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

### 4.2 启动前端

```bash
cd client
python main.py
```

---

## 5. 注意事项

- 前端依赖：`pip install PyQt5 pywin32 python-docx requests beautifulsoup4`
- 后端依赖：`pip install fastapi uvicorn python-multipart`
- Outlook 需保持登录状态
- 任务状态持久化到 `server/storage/tasks.json`
