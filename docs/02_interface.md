# 02_interface.md - 技术栈与接口定义

## 1. 技术栈

| 层级 | 技术 | 版本/说明 |
|------|------|-----------|
| 前端 GUI | PyQt5 | Python 桌面 GUI 框架 |
| Outlook 集成 | pywin32 | 操作 Outlook 邮件 |
| Word 生成 | python-docx | 生成 Word 文档（含图片） |
| HTTP 请求 | requests | 发送请求到后端 |
| 后端框架 | FastAPI | Python Web 框架 |
| 后端任务队列 | asyncio + threading | 异步处理任务 |
| 日志 | logging | 统一日志记录 |

---

## 2. 前端数据结构

### 2.1 邮件数据结构

```python
class EmailItem:
    entry_id: str              # 邮件唯一标识
    subject: str               # 邮件主题
   ConversationTopic: str     # 会话主题
    sent_on: datetime          # 发送时间
    sender: str                # 发件人
    body: str                  # 邮件正文（RTF/HTML 解析后）
    attachments: list          # 附件列表 [{name, data}]
```

### 2.2 筛选条件

```python
class FilterOptions:
    start_date: datetime | None    # 开始日期（闭区间）
    end_date: datetime | None      # 结束日期（闭区间）
    folder_path: str               # 文件夹路径，如 "收件箱"
    keyword: str | None            # 关键词搜索（搜索主题）
```

### 2.3 任务状态

```python
class TaskStatus:
    task_id: str
    status: Literal["pending", "processing", "completed", "failed"]
    progress: int                  # 进度 0-100
```

---

## 3. 后端接口定义

### 3.1 上传接口

**接口**：POST /upload

**请求**：
- Content-Type: multipart/form-data
- Body:
  - file: binary（Word 文件，图片已嵌入）

**响应**：
```json
{
    "task_id": "uuid-string",
    "status": "pending",
    "message": "任务已创建"
}
```

**说明**：后端收到 Word 文件后，mock 处理（延迟约 3 秒）后删除文件
- API 请求失败：弹窗提示用户重试

### 3.2 任务状态查询接口

**接口**：GET /task/{task_id}

**响应**：
```json
{
    "task_id": "uuid-string",
    "status": "pending" | "processing" | "completed" | "failed",
    "progress": 0
}
```

---

## 4. 项目文件结构

```
spec-email/
├── docs/
│   ├── 01_requirements.md
│   └── 02_interface.md
├── client/                      # 前端 Python 客户端
│   ├── main.py                 # 主界面
│   ├── email_window.py         # 邮件提取窗口
│   ├── outlook_client.py       # Outlook 操作封装
│   ├── word_generator.py       # Word 生成工具
│   └── api_client.py           # 后端 API 调用
└── server/                     # 后端 FastAPI
    ├── main.py                 # FastAPI 入口
    ├── task_manager.py         # 任务管理
    └── storage/                # 文件存储目录
```

---

## 5. 注意事项

- Word 文件名：ConversationTopic 过滤掉非法字符（\ / : * ? " < > |）
- Outlook 连接失败：弹窗提示用户
- 进度条：显示当前文件上传进度（如 2/10 表示第 2 个文件上传中）
- 任务状态持久化：重启后端后任务状态不丢失（写入磁盘文件）
- Word 文档中的图片需要从邮件 RTF/HTML 正文中提取并嵌入到 Word 中
- 同一 ConversationTopic 只保留最新一封（按 sent_on 排序）
- 日期区间为闭区间 [start_date, end_date]
- 后端地址：http://localhost:8000
- 任务状态存储在内存中（后续可扩展为数据库）
- Word 文件处理后删除，不持久化存储
- progress 计算：当前处理第 n 封 / 总数 × 100
