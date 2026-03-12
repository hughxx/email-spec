import logging
import os
import sys

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QPushButton, QMessageBox, QApplication


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("邮件提取客户端")
        self.resize(400, 200)

        central = QWidget()
        layout = QVBoxLayout()

        # 邮件提取按钮
        self.email_btn = QPushButton("邮件提取")
        self.email_btn.clicked.connect(self._open_email_window)
        layout.addWidget(self.email_btn)

        # 待规划按钮（灰色不可点击）
        self.plan_btn = QPushButton("待规划")
        self.plan_btn.setEnabled(False)
        self.plan_btn.setStyleSheet("background-color: gray; color: white")
        layout.addWidget(self.plan_btn)

        central.setLayout(layout)
        self.setCentralWidget(central)

        logger.info("Main window initialized")

    def _open_email_window(self) -> None:
        from email_window import EmailWindow
        w = EmailWindow()
        w.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
