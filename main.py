import sys
import os
import subprocess
import webbrowser
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGroupBox, QSizePolicy
)
from PyQt5.QtCore import QRunnable, pyqtSlot, pyqtSignal, QObject, QThreadPool, Qt
from PyQt5.QtGui import QFont


# 尝试导入网络依赖
try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    requests = None
    BeautifulSoup = None


class WorkerSignals(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    success = pyqtSignal(str, str)  # (url, h2_text)


class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @pyqtSlot()
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
            if result is not None:
                url, h2_text = result
                self.signals.success.emit(url, h2_text)
        except Exception as e:
            self.signals.error.emit(str(e))
        finally:
            self.signals.finished.emit()


# === 任务函数 ===
def generate_key_task(username, password, exe_path, output_dir):
    if not username.strip():
        raise ValueError("请输入用户名")
    if not password.strip():
        raise ValueError("请输入密码")
    exe = Path(exe_path)
    if not exe.exists():
        raise FileNotFoundError(f"找不到EXE文件: {exe}")

    cmd = f'& "{exe}" "{username.replace(chr(34), "`" + chr(34))}" "{password.replace(chr(34), "`" + chr(34))}"'
    result = subprocess.run(
        ["powershell", "-Command", cmd],
        capture_output=True, text=True, encoding='gbk', shell=True
    )
    
    output_path = Path(output_dir) / "rarreg.key"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(result.stdout)
    
    return f"程序执行完成，密钥已保存到:\n{output_path}" + (f"\n错误: {result.stderr}" if result.stderr else "")


def fetch_chinese_url_task():
    if not requests or not BeautifulSoup:
        raise ImportError("缺少 requests 或 beautifulsoup4")
    resp = requests.get('https://www.win-rar.com/download.html', headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.content, 'html.parser')
    for h2 in soup.find_all('h2'):
        if 'Chinese Simplified' in h2.get_text():
            a = h2.find('a')
            if a and a.get('href'):
                href = a['href']
                full_url = 'https://www.win-rar.com' + href if href.startswith('/') else href
                return full_url, h2.get_text(strip=True)
    raise RuntimeError("未找到包含 'Chinese Simplified' 的下载项")


# === 主窗口 ===
class WinRARExecutor(QWidget):
    def __init__(self):
        super().__init__()
        self.current_url = ""
        self.threadpool = QThreadPool()
        self.appdata_winrar = Path(os.getenv('APPDATA')) / "WinRAR"
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("WinRAR密钥生成器")
        self.setFixedSize(500, 380)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(18, 15, 18, 15)

        # ===== 区域1：获取中文版 WinRAR =====
        group1 = QGroupBox("📥 获取中文版 WinRAR")
        group1.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        layout1 = QVBoxLayout()
        layout1.setSpacing(6)

        # 第一行：最新版本标签 + 刷新按钮
        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        
        self.version_label = QLabel("最新版本：")
        self.version_label.setStyleSheet("color: #555;")
        
        self.version_text = QLabel("点击“刷新”获取版本信息")
        self.version_text.setStyleSheet("color: #000")
        self.version_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.setFixedWidth(70)
        self.refresh_btn.clicked.connect(self.start_refresh_urls)

        top_row.addWidget(self.version_label)
        top_row.addWidget(self.version_text)
        top_row.addWidget(self.refresh_btn)
        layout1.addLayout(top_row)

        # URL 输入框
        self.url_display = QLineEdit()
        self.url_display.setPlaceholderText("刷新后自动填充官方下载链接")
        self.url_display.setReadOnly(True)
        self.url_display.setFixedHeight(26)
        layout1.addWidget(self.url_display)

        # 下方提示文字
        self.tip_label = QLabel("点击“浏览器下载”将在默认浏览器中打开此链接")
        self.tip_label.setStyleSheet("color: #666; font-size: 9pt;")
        self.tip_label.setAlignment(Qt.AlignLeft)
        layout1.addWidget(self.tip_label)

        # 下载按钮
        self.download_btn = QPushButton("在浏览器中下载")
        self.download_btn.clicked.connect(self.open_in_browser)
        self.download_btn.setFixedHeight(28)
        layout1.addWidget(self.download_btn)

        group1.setLayout(layout1)
        group1.setStyleSheet("QGroupBox { padding-top: 8px; margin-top: 5px; }")
        main_layout.addWidget(group1)

        # ===== 区域2：生成注册密钥 =====
        group2 = QGroupBox("🔑 授权信息")
        group2.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        layout2 = QVBoxLayout()
        layout2.setSpacing(10)

        user_layout = QHBoxLayout()
        user_layout.setSpacing(6)
        user_layout.addWidget(QLabel("授权用户:"))
        self.username_entry = QLineEdit()
        self.username_entry.setPlaceholderText("例如：Your Name")
        user_layout.addWidget(self.username_entry)
        layout2.addLayout(user_layout)

        pwd_layout = QHBoxLayout()
        pwd_layout.setSpacing(6)
        pwd_layout.addWidget(QLabel("授权信息:"))
        self.password_entry = QLineEdit()
        self.password_entry.setPlaceholderText("请填写授权信息")
        self.password_entry.setEchoMode(QLineEdit.Password)
        pwd_layout.addWidget(self.password_entry)
        layout2.addLayout(pwd_layout)

        self.generate_btn = QPushButton("生成授权文件")
        self.generate_btn.clicked.connect(self.start_generate_key)
        layout2.addWidget(self.generate_btn)

        group2.setLayout(layout2)
        group2.setStyleSheet("QGroupBox { padding-top: 8px; margin-top: 8px; }")
        main_layout.addWidget(group2)

        self.setLayout(main_layout)
        self.setFont(QFont("Microsoft YaHei", 9))

    def run_task(self, task_fn, *args, on_success=None, on_error=None):
        worker = Worker(task_fn, *args)
        if on_success:
            worker.signals.success.connect(on_success)
        if on_error:
            worker.signals.error.connect(on_error)
        self.threadpool.start(worker)

    def start_refresh_urls(self):
        self.version_text.setText("正在获取...")
        self.url_display.clear()
        self.run_task(
            fetch_chinese_url_task,
            on_success=self.on_url_received,
            on_error=lambda err: self.on_fetch_error(err)
        )

    def on_url_received(self, url, h2_text):
        self.current_url = url
        self.url_display.setText(url)
        self.version_text.setText(h2_text)
        # 不再弹窗，信息直接显示在界面上

    def on_fetch_error(self, err):
        self.version_text.setText("获取失败")
        self.url_display.setPlaceholderText("获取失败，请重试")

    def open_in_browser(self):
        if not self.current_url:
            self.version_text.setText("请先刷新获取链接")
            return
        try:
            webbrowser.open(self.current_url)
        except Exception as e:
            self.version_text.setText("浏览器启动失败")

    def start_generate_key(self):
        # 修改：exe_path现在与main.py在同一目录下
        exe_path = Path(__file__).parent / "winrar-keygen-x64.exe"
        self.run_task(
            generate_key_task,
            self.username_entry.text(),
            self.password_entry.text(),
            exe_path,
            str(self.appdata_winrar),
            on_success=lambda msg: self.show_message("成功", msg),
            on_error=lambda err: self.show_message("错误", f"生成失败:\n{err}")
        )

    def show_message(self, title, text):
        # 保留必要弹窗（生成结果），因无其他位置显示长路径
        from PyQt5.QtWidgets import QMessageBox
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(text)
        box.exec_()


def main():
    app = QApplication(sys.argv)
    if sys.platform == "win32":
        app.setStyle("vista")
    window = WinRARExecutor()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
