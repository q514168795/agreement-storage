#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import json
import zipfile
import subprocess
import datetime
import shutil
import xml.etree.ElementTree as ET

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPoint
from PyQt6.QtGui import QFont, QColor, QPalette, QDragEnterEvent, QDropEvent, QIcon, QPainter, QBrush, QPen
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QFrame, QSizePolicy, QGraphicsDropShadowEffect
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AGREEMENT_JSON_PATH = os.path.join(BASE_DIR, "agreement.json")


def extract_docx_content(docx_path):
    """
    通过原生 zipfile 与 XML 解析 .docx 文档，提取段落文本
    """
    paragraphs = []
    with zipfile.ZipFile(docx_path) as docx:
        xml_content = docx.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        for p in tree.iterfind('.//w:p', ns):
            texts = [node.text for node in p.iterfind('.//w:t', ns) if node.text]
            full_text = ''.join(texts).strip()
            if full_text:
                paragraphs.append(full_text)
    return paragraphs


class GitPushWorker(QThread):
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, commit_msg, extra_files=None):
        super().__init__()
        self.commit_msg = commit_msg
        self.extra_files = extra_files or []

    def run(self):
        try:
            # 1. git add agreement.json 与可能存在的 APK 更新包及课件附件
            files_to_add = ["agreement.json"]
            apk_path = os.path.join(BASE_DIR, "AgreementSign_Release.apk")
            if os.path.exists(apk_path):
                files_to_add.append("AgreementSign_Release.apk")
            for extra in self.extra_files:
                if extra not in files_to_add and os.path.exists(os.path.join(BASE_DIR, extra)):
                    files_to_add.append(extra)

            res_add = subprocess.run(
                ["git", "add"] + files_to_add,
                cwd=BASE_DIR,
                capture_output=True,
                text=True
            )
            if res_add.returncode != 0:
                self.finished_signal.emit(False, f"Git add 失败: {res_add.stderr}")
                return

            # 2. git commit
            res_commit = subprocess.run(
                ["git", "commit", "-m", self.commit_msg],
                cwd=BASE_DIR,
                capture_output=True,
                text=True
            )
            if res_commit.returncode != 0 and "nothing to commit" not in res_commit.stdout and "无文件要提交" not in res_commit.stdout and "干净的工作区" not in res_commit.stdout:
                self.finished_signal.emit(False, f"Git commit 失败: {res_commit.stderr}")
                return

            # 3. git push
            res_push = subprocess.run(
                ["git", "push", "origin", "main"],
                cwd=BASE_DIR,
                capture_output=True,
                text=True
            )
            if res_push.returncode != 0:
                self.finished_signal.emit(False, f"Git push 失败: {res_push.stderr}")
                return

            self.finished_signal.emit(True, "发布成功！全网客户手机打开 App 即可自动同步最新协议与版本。")
        except Exception as e:
            self.finished_signal.emit(False, f"操作异常: {str(e)}")


class ShadcnToast(QWidget):
    """
    极简 shadcn/ui 风格浮动 Toast 提示
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.SubWindow | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(14, 10, 14, 10)

        self.label = QLabel("", self)
        self.label.setFont(QFont("Noto Sans", 10, QFont.Weight.Medium))
        self.layout.addWidget(self.label)

        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.hide)

        self.hide()

    def show_message(self, text, is_success=True):
        self.label.setText(text)
        if is_success:
            self.setStyleSheet("""
                QWidget {
                    background-color: #09090B;
                    border: 1px solid #27272A;
                    border-radius: 6px;
                }
                QLabel {
                    color: #FAFAFA;
                }
            """)
        else:
            self.setStyleSheet("""
                QWidget {
                    background-color: #7F1D1D;
                    border: 1px solid #991B1B;
                    border-radius: 6px;
                }
                QLabel {
                    color: #FEF2F2;
                }
            """)
        self.adjustSize()

        # 定位在窗口右上角
        if self.parent():
            parent_rect = self.parent().rect()
            x = parent_rect.width() - self.width() - 24
            y = 20
            self.move(x, y)

        self.show()
        self.raise_()
        self.timer.start(3500)


class DropArea(QFrame):
    """
    Word 文档拖拽上传区 (shadcn 虚线框风格)
    """
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("dropArea")
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 24, 20, 24)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.title_label = QLabel("将 Word (.docx) 文档拖拽至此处", self)
        self.title_label.setFont(QFont("Noto Sans", 11, QFont.Weight.DemiBold))
        self.title_label.setStyleSheet("color: #09090B;")
        layout.addWidget(self.title_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.sub_label = QLabel("或者点击浏览本地文件，自动提取标题与条款内容", self)
        self.sub_label.setFont(QFont("Noto Sans", 9))
        self.sub_label.setStyleSheet("color: #71717A; margin-top: 4px;")
        layout.addWidget(self.sub_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.btn_select = QPushButton("选择 Word 文件", self)
        self.btn_select.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        self.btn_select.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_select.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #09090B;
                border: 1px solid #E4E4E7;
                border-radius: 6px;
                padding: 6px 14px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background-color: #F4F4F5;
            }
        """)
        self.btn_select.clicked.connect(self.select_file)
        layout.addWidget(self.btn_select, alignment=Qt.AlignmentFlag.AlignCenter)

        self.apply_style(False)

    def apply_style(self, is_hover=False):
        border_color = "#18181B" if is_hover else "#D4D4D8"
        bg_color = "#F4F4F5" if is_hover else "#FFFFFF"
        self.setStyleSheet(f"""
            QFrame#dropArea {{
                background-color: {bg_color};
                border: 1.5px dashed {border_color};
                border-radius: 8px;
            }}
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(".docx"):
                event.acceptProposedAction()
                self.apply_style(True)
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.apply_style(False)

    def dropEvent(self, event: QDropEvent):
        self.apply_style(False)
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(".docx"):
                self.file_dropped.emit(file_path)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.select_file()

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 协议文档", "", "Word 文档 (*.docx)"
        )
        if file_path:
            self.file_dropped.emit(file_path)


class DocDropArea(QFrame):
    """
    课件与附件 (PPT / PDF) 拖拽与配置区 (shadcn 虚线框风格)
    """
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("docDropArea")
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.title_label = QLabel("将 课件 (PPT / PDF) 拖拽至此处", self)
        self.title_label.setFont(QFont("Noto Sans", 10, QFont.Weight.DemiBold))
        self.title_label.setStyleSheet("color: #09090B;")
        layout.addWidget(self.title_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.sub_label = QLabel("选填：配置后手机端将开启双锁模式与封面存证", self)
        self.sub_label.setFont(QFont("Noto Sans", 8))
        self.sub_label.setStyleSheet("color: #71717A; margin-top: 2px;")
        layout.addWidget(self.sub_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.btn_select = QPushButton("选择课件文件", self)
        self.btn_select.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        self.btn_select.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_select.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #09090B;
                border: 1px solid #E4E4E7;
                border-radius: 6px;
                padding: 5px 12px;
                margin-top: 8px;
            }
            QPushButton:hover {
                background-color: #F4F4F5;
            }
        """)
        self.btn_select.clicked.connect(self.select_file)
        layout.addWidget(self.btn_select, alignment=Qt.AlignmentFlag.AlignCenter)

        self.apply_style(False)

    def apply_style(self, is_hover=False):
        border_color = "#18181B" if is_hover else "#D4D4D8"
        bg_color = "#F4F4F5" if is_hover else "#FFFFFF"
        self.setStyleSheet(f"""
            QFrame#docDropArea {{
                background-color: {bg_color};
                border: 1.5px dashed {border_color};
                border-radius: 8px;
            }}
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile().lower()
                if path.endswith(".pdf") or path.endswith(".pptx") or path.endswith(".ppt"):
                    event.acceptProposedAction()
                    self.apply_style(True)
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.apply_style(False)

    def dropEvent(self, event: QDropEvent):
        self.apply_style(False)
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            path_lower = file_path.lower()
            if path_lower.endswith(".pdf") or path_lower.endswith(".pptx") or path_lower.endswith(".ppt"):
                self.file_dropped.emit(file_path)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.select_file()

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择培训课件 / 附件文档", "", "课件文档 (*.pdf *.pptx *.ppt);;PDF 文档 (*.pdf);;PowerPoint (*.pptx *.ppt)"
        )
        if file_path:
            self.file_dropped.emit(file_path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EVObike 协议发布助手 • shadcn/ui")
        self.resize(860, 740)
        self.setMinimumSize(780, 620)
        self.service_active = True
        self.attached_doc_path = ""
        self.current_doc_url = ""
        self.current_doc_title = ""

        # 全局样式采用 shadcn/ui 风格
        self.setStyleSheet("""
            QMainWindow {
                background-color: #FAFAFA;
            }
            QWidget {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans", sans-serif;
            }
            QLineEdit, QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #E4E4E7;
                border-radius: 6px;
                padding: 8px 12px;
                color: #09090B;
                selection-background-color: #E4E4E7;
                selection-color: #09090B;
            }
            QLineEdit:focus, QTextEdit:focus {
                border: 1px solid #09090B;
            }
            QScrollBar:vertical {
                border: none;
                background: #F4F4F5;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #D4D4D8;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #A1A1AA;
            }
        """)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(28, 24, 28, 24)
        main_layout.setSpacing(18)

        # 1. 顶部 Header 区
        header_layout = QHBoxLayout()
        header_text_layout = QVBoxLayout()

        badges_layout = QHBoxLayout()
        badge_brand = QLabel("⚡ EVOBIKE CLOUD", self)
        badge_brand.setStyleSheet("""
            background-color: #09090B;
            color: #FAFAFA;
            border-radius: 4px;
            padding: 3px 8px;
            font-size: 11px;
            font-weight: bold;
        """)
        badges_layout.addWidget(badge_brand)

        self.badge_status = QLabel("已连接 GitHub 存储库", self)
        self.badge_status.setStyleSheet("""
            background-color: #ECFDF5;
            color: #059669;
            border: 1px solid #A7F3D0;
            border-radius: 4px;
            padding: 2px 7px;
            font-size: 11px;
            font-weight: 500;
        """)
        badges_layout.addWidget(self.badge_status)
        badges_layout.addStretch()

        header_text_layout.addLayout(badges_layout)

        title_label = QLabel("EVObike 协议云端发布与终端管理中心", self)
        title_label.setFont(QFont("Noto Sans", 16, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #09090B; margin-top: 4px;")
        header_text_layout.addWidget(title_label)

        sub_label = QLabel("导入或拖入 Word (.docx) 协议文档，自动解析条款并一键推送到 GitHub，手机端免装/自动升级实时生效。", self)
        sub_label.setFont(QFont("Noto Sans", 10))
        sub_label.setStyleSheet("color: #71717A; margin-top: 2px;")
        header_text_layout.addWidget(sub_label)

        header_layout.addLayout(header_text_layout)
        header_layout.addStretch()

        main_layout.addLayout(header_layout)

        # 分割线
        divider = QFrame(self)
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setStyleSheet("color: #E4E4E7;")
        main_layout.addWidget(divider)

        # 2. 导入区：左侧 Word 条款导入，右侧 PPT/PDF 课件导入
        import_row_layout = QHBoxLayout()
        import_row_layout.setSpacing(12)

        self.drop_area = DropArea(self)
        self.drop_area.file_dropped.connect(self.on_file_loaded)
        import_row_layout.addWidget(self.drop_area, stretch=1)

        self.doc_drop_area = DocDropArea(self)
        self.doc_drop_area.file_dropped.connect(self.on_doc_file_selected)
        import_row_layout.addWidget(self.doc_drop_area, stretch=1)

        main_layout.addLayout(import_row_layout)

        # 课件状态微调栏 (shadcn 极简微卡片)
        self.doc_status_frame = QFrame(self)
        self.doc_status_frame.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #E4E4E7;
                border-radius: 6px;
            }
        """)
        doc_status_layout = QHBoxLayout(self.doc_status_frame)
        doc_status_layout.setContentsMargins(14, 8, 14, 8)
        doc_status_layout.setSpacing(10)

        self.lbl_doc_badge = QLabel("课件状态:", self.doc_status_frame)
        self.lbl_doc_badge.setFont(QFont("Noto Sans", 9, QFont.Weight.DemiBold))
        self.lbl_doc_badge.setStyleSheet("color: #18181B; border: none;")
        doc_status_layout.addWidget(self.lbl_doc_badge)

        self.lbl_doc_info = QLabel("未添加课件附件（手机端将自动以纯文字单锁模式运行）", self.doc_status_frame)
        self.lbl_doc_info.setFont(QFont("Noto Sans", 9))
        self.lbl_doc_info.setStyleSheet("color: #71717A; border: none;")
        doc_status_layout.addWidget(self.lbl_doc_info, stretch=1)

        self.edit_doc_title = QLineEdit(self.doc_status_frame)
        self.edit_doc_title.setPlaceholderText("课件展示名称 (如: EVOBIKE 员工规章培训)")
        self.edit_doc_title.setVisible(False)
        doc_status_layout.addWidget(self.edit_doc_title, stretch=2)

        self.btn_clear_doc = QPushButton("× 移除课件", self.doc_status_frame)
        self.btn_clear_doc.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        self.btn_clear_doc.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_clear_doc.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #DC2626;
                border: 1px solid #FECACA;
                border-radius: 4px;
                padding: 4px 10px;
            }
            QPushButton:hover {
                background-color: #FEF2F2;
            }
        """)
        self.btn_clear_doc.clicked.connect(self.clear_attached_doc)
        self.btn_clear_doc.setVisible(False)
        doc_status_layout.addWidget(self.btn_clear_doc)

        main_layout.addWidget(self.doc_status_frame)

        # 3. 协议编辑卡片 (白底、1px细线、微圆角)
        card_frame = QFrame(self)
        card_frame.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #E4E4E7;
                border-radius: 8px;
            }
        """)
        card_layout = QVBoxLayout(card_frame)
        card_layout.setContentsMargins(18, 16, 18, 16)
        card_layout.setSpacing(12)

        # 标题行 + 版本 + 日期
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(12)

        # 协议标题
        vbox_title = QVBoxLayout()
        lbl_t = QLabel("协议主标题 (移动端将居中显示)", card_frame)
        lbl_t.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        lbl_t.setStyleSheet("color: #18181B; border: none;")
        vbox_title.addWidget(lbl_t)

        self.edit_title = QLineEdit(card_frame)
        self.edit_title.setPlaceholderText("例如：EVObike 标准服务与保密协议")
        vbox_title.addWidget(self.edit_title)
        row1_layout.addLayout(vbox_title, stretch=3)

        # 版本号
        vbox_ver = QVBoxLayout()
        lbl_v = QLabel("版本号", card_frame)
        lbl_v.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        lbl_v.setStyleSheet("color: #18181B; border: none;")
        vbox_ver.addWidget(lbl_v)

        self.edit_version = QLineEdit(card_frame)
        self.edit_version.setPlaceholderText("2.0")
        vbox_ver.addWidget(self.edit_version)
        row1_layout.addLayout(vbox_ver, stretch=1)

        # 更新日期
        vbox_date = QVBoxLayout()
        lbl_d = QLabel("更新日期", card_frame)
        lbl_d.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        lbl_d.setStyleSheet("color: #18181B; border: none;")
        vbox_date.addWidget(lbl_d)

        self.edit_date = QLineEdit(card_frame)
        self.edit_date.setPlaceholderText("YYYY-MM-DD")
        vbox_date.addWidget(self.edit_date)
        row1_layout.addLayout(vbox_date, stretch=1)

        card_layout.addLayout(row1_layout)

        # 协议正文内容预览与编辑
        vbox_content = QVBoxLayout()
        lbl_c = QLabel("协议正文条款 (用户在手机端需完整由上至下滑动浏览以解锁签署)", card_frame)
        lbl_c.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        lbl_c.setStyleSheet("color: #18181B; border: none;")
        vbox_content.addWidget(lbl_c)

        self.edit_content = QTextEdit(card_frame)
        self.edit_content.setFont(QFont("Noto Sans", 10))
        self.edit_content.setPlaceholderText("导入 Word 文档后将在此处呈现正文条款，您可以直接编辑或补充条款...")
        vbox_content.addWidget(self.edit_content)

        card_layout.addLayout(vbox_content)
        main_layout.addWidget(card_frame, stretch=1)

        # 4. 底部操作栏
        action_layout = QHBoxLayout()

        # 终端状态一键控制按钮 (隐蔽式服务停用与数据重置)
        self.btn_toggle_service = QPushButton(self)
        self.btn_toggle_service.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        self.btn_toggle_service.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle_service.clicked.connect(self.toggle_service_status)
        action_layout.addWidget(self.btn_toggle_service)

        action_layout.addStretch()

        self.btn_reload = QPushButton("重新载入本地原版", self)
        self.btn_reload.setFont(QFont("Noto Sans", 9, QFont.Weight.Medium))
        self.btn_reload.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reload.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #09090B;
                border: 1px solid #E4E4E7;
                border-radius: 6px;
                padding: 9px 16px;
            }
            QPushButton:hover {
                background-color: #F4F4F5;
            }
        """)
        self.btn_reload.clicked.connect(self.load_current_json)
        action_layout.addWidget(self.btn_reload)

        self.btn_publish = QPushButton("一键发布到云端 (GitHub)", self)
        self.btn_publish.setFont(QFont("Noto Sans", 10, QFont.Weight.DemiBold))
        self.btn_publish.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_publish.setStyleSheet("""
            QPushButton {
                background-color: #18181B;
                color: #FAFAFA;
                border: 1px solid #18181B;
                border-radius: 6px;
                padding: 9px 20px;
            }
            QPushButton:hover {
                background-color: #27272A;
            }
            QPushButton:disabled {
                background-color: #A1A1AA;
                border-color: #A1A1AA;
            }
        """)
        self.btn_publish.clicked.connect(self.publish_to_cloud)
        action_layout.addWidget(self.btn_publish)

        main_layout.addLayout(action_layout)

        # 浮动 Toast 组件
        self.toast = ShadcnToast(self)

        # 加载初始数据
        self.load_current_json()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.toast.isVisible():
            x = self.width() - self.toast.width() - 24
            y = 20
            self.toast.move(x, y)

    def update_service_ui(self):
        if self.service_active:
            self.btn_toggle_service.setText("终端协议: 正常运行 (点击一键注销)")
            self.btn_toggle_service.setStyleSheet("""
                QPushButton {
                    background-color: #FFFFFF;
                    color: #09090B;
                    border: 1px solid #E4E4E7;
                    border-radius: 6px;
                    padding: 8px 14px;
                }
                QPushButton:hover {
                    background-color: #F4F4F5;
                }
            """)
        else:
            self.btn_toggle_service.setText("终端协议: 已注销停用 (点击恢复)")
            self.btn_toggle_service.setStyleSheet("""
                QPushButton {
                    background-color: #7F1D1D;
                    color: #FAFAFA;
                    border: 1px solid #991B1B;
                    border-radius: 6px;
                    padding: 8px 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #991B1B;
                }
            """)

    def toggle_service_status(self):
        self.service_active = not self.service_active
        self.update_service_ui()
        state_str = "正常开启" if self.service_active else "已注销并清理本地缓存"
        self.toast.show_message(f"终端协议状态已切换为: {state_str}，点击发布后全网生效", is_success=self.service_active)

    def on_doc_file_selected(self, file_path):
        self.attached_doc_path = file_path
        self.current_doc_url = ""
        base_name = os.path.basename(file_path)
        title_guess = os.path.splitext(base_name)[0]
        self.set_doc_ui_bound(f"已选本地课件: {base_name}", title_guess)
        self.toast.show_message(f"已选课件: {base_name}", is_success=True)

    def set_doc_ui_bound(self, info_text, title_text):
        self.lbl_doc_info.setText(info_text)
        self.lbl_doc_info.setStyleSheet("color: #059669; font-weight: 500; border: none;")
        self.edit_doc_title.setText(title_text)
        self.edit_doc_title.setVisible(True)
        self.btn_clear_doc.setVisible(True)

    def clear_attached_doc(self, silent=False):
        self.attached_doc_path = ""
        self.current_doc_url = ""
        self.current_doc_title = ""
        self.lbl_doc_info.setText("未添加课件附件（手机端将自动以纯文字单锁模式运行）")
        self.lbl_doc_info.setStyleSheet("color: #71717A; border: none;")
        self.edit_doc_title.clear()
        self.edit_doc_title.setVisible(False)
        self.btn_clear_doc.setVisible(False)
        if not silent:
            self.toast.show_message("已清空课件，手机端将以纯文字模式下发", is_success=True)

    def load_current_json(self):
        if os.path.exists(AGREEMENT_JSON_PATH):
            try:
                with open(AGREEMENT_JSON_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.edit_title.setText(data.get("title", "EVObike 标准图文服务与保密协议"))
                    self.edit_version.setText(data.get("version", "2.1"))
                    self.edit_date.setText(data.get("updated_at", str(datetime.date.today())))
                    self.edit_content.setPlainText(data.get("content", ""))
                    self.service_active = data.get("service_active", True)

                    # 载入课件配置状态
                    doc_url = data.get("doc_url", data.get("ppt_url", data.get("pdf_url", "")))
                    doc_title = data.get("doc_title", data.get("ppt_title", data.get("pdf_title", "")))
                    if doc_url:
                        self.current_doc_url = doc_url
                        self.current_doc_title = doc_title
                        self.attached_doc_path = ""
                        self.set_doc_ui_bound(f"云端已绑定课件: {os.path.basename(doc_url)}", doc_title)
                    else:
                        self.clear_attached_doc(silent=True)

                    self.update_service_ui()
            except Exception as e:
                self.toast.show_message(f"加载已有配置失败: {str(e)}", is_success=False)
        else:
            self.edit_title.setText("EVObike 标准图文服务与保密协议")
            self.edit_version.setText("2.1")
            self.edit_date.setText(str(datetime.date.today()))
            self.service_active = True
            self.clear_attached_doc(silent=True)
            self.update_service_ui()

    def on_file_loaded(self, file_path):
        try:
            paragraphs = extract_docx_content(file_path)
            if not paragraphs:
                self.toast.show_message("文档解析为空，请检查 Word 文件内容", is_success=False)
                return

            first = paragraphs[0]
            if len(first) <= 40:
                self.edit_title.setText(first)
                body_paragraphs = paragraphs[1:]
            else:
                body_paragraphs = paragraphs

            content_text = "\n\n".join(body_paragraphs)
            self.edit_content.setPlainText(content_text)

            cur_ver = self.edit_version.text().strip()
            try:
                parts = cur_ver.split(".")
                if len(parts) >= 2:
                    new_ver = f"{parts[0]}.{int(parts[1]) + 1}"
                else:
                    new_ver = f"{int(cur_ver) + 1}.0"
            except Exception:
                new_ver = "2.1"
            self.edit_version.setText(new_ver)
            self.edit_date.setText(str(datetime.date.today()))

            self.toast.show_message(f"已成功解析: {os.path.basename(file_path)}", is_success=True)
        except Exception as e:
            self.toast.show_message(f"解析 Word 失败: {str(e)}", is_success=False)

    def publish_to_cloud(self):
        title = self.edit_title.text().strip()
        version = self.edit_version.text().strip()
        date_str = self.edit_date.text().strip()
        content = self.edit_content.toPlainText().strip()

        if not title:
            self.toast.show_message("请填写协议标题", is_success=False)
            return
        if not content:
            self.toast.show_message("协议正文不能为空", is_success=False)
            return

        extra_files = []
        doc_url = ""
        doc_title = self.edit_doc_title.text().strip()
        doc_type = "pdf"

        # 如果用户选取了本地课件文件，复制到仓库并在 git 中同步推送
        if self.attached_doc_path and os.path.exists(self.attached_doc_path):
            ext = os.path.splitext(self.attached_doc_path)[1].lower()
            target_name = f"doc_attachment{ext}"
            target_path = os.path.join(BASE_DIR, target_name)
            try:
                shutil.copy2(self.attached_doc_path, target_path)
                extra_files.append(target_name)
                doc_url = f"https://raw.githubusercontent.com/q514168795/agreement-storage/main/{target_name}"
                doc_type = "pptx" if ext in [".ppt", ".pptx"] else "pdf"
                if not doc_title:
                    doc_title = os.path.splitext(os.path.basename(self.attached_doc_path))[0]
            except Exception as e:
                self.toast.show_message(f"复制课件文件失败: {str(e)}", is_success=False)
                return
        elif self.current_doc_url:
            doc_url = self.current_doc_url
            if not doc_title:
                doc_title = self.current_doc_title
            doc_type = "pptx" if doc_url.lower().endswith((".ppt", ".pptx")) else "pdf"

        # 保留可能存在的既有多语言字段
        existing_data = {}
        if os.path.exists(AGREEMENT_JSON_PATH):
            try:
                with open(AGREEMENT_JSON_PATH, "r", encoding="utf-8") as f:
                    existing_data = json.load(f)
            except Exception:
                pass

        payload = {
            **existing_data,
            "version": version or "2.1",
            "title": title,
            "content": content,
            "updated_at": date_str or str(datetime.date.today()),
            "app_version_code": 3,
            "app_version_name": "2.1.0",
            "app_apk_url": "https://raw.githubusercontent.com/q514168795/agreement-storage/main/AgreementSign_Release.apk",
            "app_update_notes": "1. 全新 shadcn/ui 极简高对比设计规范\n2. 支持 PDF/PPT 课件与双锁阅读达标拦截\n3. 凭证长图自适应合成 PPT 封面微缩与核验卡\n4. 换装 EVObike 官方品牌标识与防伪印鉴",
            "service_active": self.service_active,
            "doc_url": doc_url,
            "doc_title": doc_title,
            "doc_type": doc_type
        }

        # 写入本地 agreement.json
        try:
            with open(AGREEMENT_JSON_PATH, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.toast.show_message(f"保存文件失败: {str(e)}", is_success=False)
            return

        self.btn_publish.setEnabled(False)
        self.btn_publish.setText("正在同步至云端...")
        self.badge_status.setText("正在推送至 GitHub...")

        commit_msg = f"Update EVObike agreement: {title} v{version}"
        self.worker = GitPushWorker(commit_msg, extra_files=extra_files)
        self.worker.finished_signal.connect(self.on_push_finished)
        self.worker.start()

    def on_push_finished(self, success, message):
        self.btn_publish.setEnabled(True)
        self.btn_publish.setText("一键发布到云端 (GitHub)")
        if success:
            self.badge_status.setText("已同步最新版")
            self.toast.show_message(message, is_success=True)
        else:
            self.badge_status.setText("同步失败")
            self.toast.show_message(message, is_success=False)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
