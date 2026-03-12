import logging
import os
import tempfile
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem,
    QListWidget, QListWidgetItem, QPushButton, QDateEdit, QLabel,
    QLineEdit, QProgressBar, QMessageBox
)
from PyQt5.QtCore import Qt, QDate
from outlook_client import OutlookClient
from word_generator import WordGenerator
from api_client import APIClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class EmailWindow(QDialog):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("邮件提取")
        self.resize(900, 600)
        self.outlook_client = OutlookClient()
        self.api_client = APIClient()
        self.folder_tree = {}
        self.emails = []
        self.task_ids = []
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout()

        # 筛选条件
        filter_layout = QHBoxLayout()

        # 文件夹选择
        filter_layout.addWidget(QLabel("文件夹:"))
        self.folder_tree_widget = QTreeWidget()
        self.folder_tree_widget.setHeaderHidden(True)
        self._load_folders()
        self.folder_tree_widget.itemClicked.connect(self._on_folder_selected)
        filter_layout.addWidget(self.folder_tree_widget)

        # 日期筛选
        filter_layout.addWidget(QLabel("开始日期:"))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addMonths(-1))
        filter_layout.addWidget(self.start_date)

        filter_layout.addWidget(QLabel("结束日期:"))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        filter_layout.addWidget(self.end_date)

        # 关键词搜索
        filter_layout.addWidget(QLabel("关键词:"))
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("搜索主题...")
        filter_layout.addWidget(self.keyword_input)

        # 搜索按钮
        self.search_btn = QPushButton("搜索")
        self.search_btn.clicked.connect(self._on_search)
        filter_layout.addWidget(self.search_btn)

        layout.addLayout(filter_layout)

        # 邮件列表
        self.email_list = QListWidget()
        self.email_list.setSelectionMode(QListWidget.MultiSelection)
        layout.addWidget(QLabel("邮件列表:"))
        layout.addWidget(self.email_list)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.extract_btn = QPushButton("提取")
        self.extract_btn.clicked.connect(self._on_extract)
        self.extract_btn.setEnabled(False)
        btn_layout.addWidget(self.extract_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        btn_layout.addWidget(self.progress_bar)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _load_folders(self) -> None:
        """加载文件夹树"""
        self.folder_tree = self.outlook_client.get_folder_tree()
        # 默认收件箱
        self.folder_tree_widget.addTopLevelItem(QTreeWidgetItem(["收件箱"]))
        # 添加其他文件夹
        for path in self.folder_tree:
            if path != "收件箱":
                self.folder_tree_widget.addTopLevelItem(QTreeWidgetItem([path]))

    def _on_folder_selected(self, item: QTreeWidgetItem, column: int) -> None:
        self.selected_folder = item.text(0)

    def _on_search(self) -> None:
        """搜索邮件"""
        try:
            folder = getattr(self, "selected_folder", "收件箱") or "收件箱"
            start = self.start_date.date().toPyDate()
            end = self.end_date.date().toPyDate()
            keyword = self.keyword_input.text() or None

            self.emails = self.outlook_client.get_emails(
                folder, start, end, keyword
            )

            # 显示邮件
            self.email_list.clear()
            for email in self.emails:
                item = QListWidgetItem(f"{email.sent_on.strftime('%Y-%m-%d')} - {email.conversation_topic}")
                item.setData(Qt.UserRole, email)
                self.email_list.addItem(item)

            self.extract_btn.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"搜索失败: {str(e)}")

    def _on_extract(self) -> None:
        """提取选中的邮件"""
        selected = self.email_list.selectedItems()
        if not selected:
            QMessageBox.warning(self, "提示", "请先选择邮件")
            return

        self.extract_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(selected))
        self.progress_bar.setValue(0)
        self.task_ids = []

        output_dir = tempfile.mkdtemp()

        try:
            for i, item in enumerate(selected):
                email = item.data(Qt.UserRole)
                word_path = None
                try:
                    # 生成 Word
                    word_path = WordGenerator.generate_word(email, output_dir)

                    # 上传到后端
                    result = self.api_client.upload_word(word_path)
                    self.task_ids.append(result["task_id"])

                except Exception as e:
                    QMessageBox.warning(
                        self, "上传失败",
                        f"{email.conversation_topic}: 请重试"
                    )
                finally:
                    # 立即删除临时 Word 文件
                    if word_path and os.path.exists(word_path):
                        os.remove(word_path)

                self.progress_bar.setValue(i + 1)

        finally:
            # 清理临时目录
            if os.path.exists(output_dir):
                import shutil
                shutil.rmtree(output_dir, ignore_errors=True)

        self.extract_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        if self.task_ids:
            QMessageBox.information(
                self, "成功",
                f"已提交 {len(self.task_ids)} 个任务\n点击确定查看状态",
                QMessageBox.Ok
            )
            self._show_task_status()

    def _show_task_status(self) -> None:
        """显示任务状态"""
        status_text = ""
        for task_id in self.task_ids:
            result = self.api_client.get_task_status(task_id)
            if result:
                status_text += f"{task_id}: {result['status']}\n"
        QMessageBox.information(self, "任务状态", status_text or "查询失败")
