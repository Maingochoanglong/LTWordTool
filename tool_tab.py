#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tool_tab.py
===========
Khung sườn Qt6 (PySide6) dùng chung cho mọi tab công cụ trong gui.py --
tách ra từ gui.py để thêm 1 công cụ mới (tab mới) không phải viết lại
phần luồng nền / nút bấm / nhật ký, vốn giống hệt nhau ở mọi tab.

3 thứ export ra:

  FilePickerRow  - 1 hàng chọn file, tự nhớ đường dẫn qua QSettings.
  CallableWorker - QThread tổng quát: chạy 1 hàm bất kỳ (func(*a, **kw))
                   ở luồng riêng, emit finished_ok(kết quả) hoặc
                   failed(thông báo lỗi) -- KHÔNG cần viết 1 class
                   QThread riêng cho mỗi công cụ như trước nữa.
  ToolTab        - lớp nền cho 1 tab công cụ: tự dựng khung UI (nhóm
                   chọn file, nút Chạy + Mở kết quả, ô nhật ký), tự nối
                   CallableWorker, tự bật/tắt nút, tự bắt lỗi hiện
                   QMessageBox. Lớp con (1 công cụ cụ thể) CHỈ cần viết
                   2 hàm:
                     collect_call()   - kiểm tra input, trả về
                                        (func, args, kwargs) để chạy nền
                                        (hoặc None nếu thiếu, tự hiện
                                        cảnh báo trước khi return None)
                     format_result(x) - kết quả func trả về -> list các
                                        dòng text để ghi log khi xong
                   Xem gui.py (ReplaceTab, MathTypeFixTab) để có ví dụ.

Không phụ thuộc replace_docx.py hay fix_mathtype_parens.py -- file này
không biết và không cần biết có những công cụ gì, chỉ lo phần khung UI.
"""

import os

from PySide6.QtCore import QSettings, QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

DOCX_FILTER = "Word Document (*.docx)"
ORG_NAME = "DocxReplacer"
APP_NAME = "DocxReplacer"


def _app_settings():
    """QSettings dùng file .ini riêng của ứng dụng thay vì định dạng
    "native" cũ (trên Windows, native = registry). Dùng constructor 4 tham
    số (IniFormat, UserScope, org, app) để Qt tự chọn đường dẫn chuẩn theo
    từng hệ điều hành, vd %APPDATA%\\DocxReplacer\\DocxReplacer.ini trên
    Windows, ~/.config/DocxReplacer/DocxReplacer.ini trên Linux."""
    return QSettings(QSettings.IniFormat, QSettings.UserScope, ORG_NAME, APP_NAME)


class FilePickerRow(QWidget):
    """1 hàng: nhãn + ô đường dẫn (chỉ đọc) + nút chọn file.

    Nếu có settings_key, đường dẫn được tự nhớ lại (qua QSettings) giữa các
    lần mở ứng dụng - không phải chọn lại từ đầu mỗi lần mở GUI.
    """

    def __init__(self, label_text, save_mode=False, default_name="",
                 settings_key="", placeholder="Chưa chọn file..."):
        super().__init__()
        self.save_mode = save_mode
        self.default_name = default_name
        self.settings_key = settings_key

        label = QLabel(label_text)
        label.setFixedWidth(190)

        self.edit = QLineEdit()
        self.edit.setReadOnly(True)
        self.edit.setPlaceholderText(placeholder)

        button = QPushButton("Chọn file...")
        button.clicked.connect(self.pick_file)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(label)
        layout.addWidget(self.edit, 1)
        layout.addWidget(button)

        self.load_setting()

    def pick_file(self):
        start = self.path() or self.default_name
        if self.save_mode:
            path, _ = QFileDialog.getSaveFileName(self, "Chọn nơi lưu file", start, DOCX_FILTER)
            if path and not path.lower().endswith(".docx"):
                path += ".docx"
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Chọn file Word", start, DOCX_FILTER)
        if path:
            self.edit.setText(path)
            self.save_setting()

    def path(self):
        return self.edit.text().strip()

    def save_setting(self):
        if self.settings_key:
            _app_settings().setValue(self.settings_key, self.path())

    def load_setting(self):
        if self.settings_key:
            saved = _app_settings().value(self.settings_key, "")
            if saved:
                self.edit.setText(saved)


class CallableWorker(QThread):
    """Chạy func(*args, **kwargs) ở luồng riêng để giao diện không bị đơ.
    Dùng chung cho MỌI công cụ -- không cần viết 1 class QThread riêng
    cho từng công cụ như trước (ReplaceWorker/MathTypeFixWorker cũ)."""

    finished_ok = Signal(object)   # kết quả func trả về, kiểu gì cũng được
    failed = Signal(str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
        except Exception as e:  # noqa: BLE001 - hiển thị mọi lỗi ra giao diện
            self.failed.emit(str(e))
        else:
            self.finished_ok.emit(result)


class ToolTab(QWidget):
    """Khung sườn dùng chung cho 1 tab công cụ: nhóm 'Chọn file' (từ
    file_rows), nút chạy + nút 'Mở file kết quả', ô nhật ký. Lo hết phần
    luồng nền (CallableWorker), bật/tắt nút, và bắt lỗi hiện QMessageBox.

    Thêm 1 công cụ mới = kế thừa lớp này, cung cấp file_rows/result_row,
    rồi viết đúng 2 hàm collect_call() và format_result() (bắt buộc) và
    tuỳ chọn extra_widget() nếu muốn thêm ghi chú/widget khác dưới log."""

    def __init__(self, run_label, file_rows, result_row, log_height=150):
        """run_label: chữ trên nút chạy.
        file_rows: danh sách FilePickerRow của tab, theo đúng thứ tự hiển thị.
        result_row: FilePickerRow nào giữ đường dẫn file kết quả (để nút
        'Mở file kết quả' biết mở gì); truyền None nếu tab không có khái
        niệm 1 file kết quả duy nhất (nút Mở sẽ luôn tắt)."""
        super().__init__()
        self.worker = None
        self.file_rows = file_rows
        self.result_row = result_row

        pick_box = QGroupBox("Chọn file")
        pick_layout = QVBoxLayout(pick_box)
        for row in file_rows:
            pick_layout.addWidget(row)

        self.run_button = QPushButton(run_label)
        self.run_button.clicked.connect(self._on_run_clicked)

        self.open_button = QPushButton("Mở file kết quả")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self.open_result)

        button_row = QHBoxLayout()
        button_row.addWidget(self.run_button)
        button_row.addWidget(self.open_button)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(log_height)

        layout = QVBoxLayout(self)
        layout.addWidget(pick_box)
        layout.addLayout(button_row)
        layout.addWidget(QLabel("Nhật ký:"))
        layout.addWidget(self.log)
        extra = self.extra_widget()
        if extra is not None:
            layout.addWidget(extra)

    # ---- lớp con override khi cần -------------------------------------

    def extra_widget(self):
        """Trả về 1 QWidget để thêm dưới ô nhật ký (vd ghi chú), hoặc None
        (mặc định) nếu không cần gì thêm."""
        return None

    def collect_call(self):
        """PHẢI override. Kiểm tra input hiện tại, trả về (func, args,
        kwargs) để chạy nền -- args là tuple, kwargs là dict. Nếu input
        thiếu/sai: tự hiện QMessageBox cảnh báo (dùng self làm parent) rồi
        return None để huỷ, KHÔNG được raise."""
        raise NotImplementedError

    def format_result(self, result):
        """PHẢI override. result là đúng giá trị mà func ở collect_call()
        trả về khi chạy xong. Trả về list các dòng text để ghi vào nhật ký."""
        raise NotImplementedError

    # ---- phần dùng chung, lớp con không cần đụng vào -------------------

    def log_msg(self, text):
        self.log.append(text)

    def _on_run_clicked(self):
        call = self.collect_call()
        if call is None:
            return
        func, args, kwargs = call

        self.run_button.setEnabled(False)
        self.open_button.setEnabled(False)
        self.log_msg("Đang xử lý...")

        self.worker = CallableWorker(func, *args, **kwargs)
        self.worker.finished_ok.connect(self._on_success)
        self.worker.failed.connect(self._on_error)
        self.worker.start()

    def _on_success(self, result):
        self.run_button.setEnabled(True)
        self.open_button.setEnabled(True)
        for line in self.format_result(result):
            self.log_msg(line)

    def _on_error(self, message):
        self.run_button.setEnabled(True)
        self.log_msg(f"✘ Lỗi: {message}")
        QMessageBox.critical(self, "Lỗi", message)

    def open_result(self):
        path = self.result_row.path() if self.result_row else ""
        if path and os.path.isfile(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
