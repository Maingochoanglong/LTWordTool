#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
panel.py
========
Khung Qt6 (PySide6) DÙNG CHUNG cho phần "vỏ" của app: luồng nền, nút
bấm, nhật ký, nhớ đường dẫn, dịch lỗi -- tách khỏi gui.py để phần Qt
thuần tuý (không biết gì về docx/MTEF) không lẫn với phần UI/logic cụ
thể của app (checkbox nào, gọi hàm nào). App hiện chỉ có 1 panel
(CombinedPanel), không có tab nào cả (đã bỏ QTabWidget từ lâu) -- nhưng
tách vẫn có lợi: gộp vào gui.py sẽ ra 1 file lẫn lộn 2 việc khác hẳn
nhau (khung Qt chung vs. nghiệp vụ cụ thể của app).

(Trước đây file này tên tool_tab.py, class ToolTab -- đổi tên vì app đã
bỏ QTabWidget từ lâu, chữ "tab" không còn phản ánh gì trong UI nữa.)

6 thứ export ra:

  FilePickerRow        - 1 hàng chọn file, tự nhớ đường dẫn qua QSettings.
  save_path_pairs/
  load_path_pairs      - lưu & đọc lại 1 DANH SÁCH cặp (path1, path2) bất
                          kỳ qua QSettings (dạng mảng) -- dùng cho panel
                          nào cần nhớ NHIỀU cặp đường dẫn cùng lúc, khác
                          FilePickerRow (chỉ nhớ được 1 đường dẫn/khoá).
                          Không biết và không cần biết 2 vế của cặp mang ý
                          nghĩa gì (tìm/thay hay bất kỳ gì khác) -- chỉ lo
                          cơ chế lưu/đọc chung.
  friendly_error_text  - dịch vài lỗi HỆ THỐNG hay gặp sang tiếng Việt
                          (xem docstring hàm để biết lỗi nào KHÔNG dịch).
  CallableWorker       - QThread tổng quát: chạy 1 hàm bất kỳ (func(*a,
                          **kw)) ở luồng riêng, emit finished_ok(kết quả)
                          hoặc failed(thông báo lỗi đã dịch).
  RunPanel             - lớp nền cho 1 panel: tự dựng khung UI (khu vực
                          chọn file, nút Chạy + Mở kết quả, ô nhật ký),
                          tự nối CallableWorker, tự bật/tắt nút, tự bắt
                          lỗi hiện QMessageBox. Lớp con CHỈ cần viết đúng
                          2 hàm:
                            collect_call()   - kiểm tra input, trả về
                                               (func, args) để chạy nền
                                               (hoặc None nếu thiếu, tự
                                               hiện cảnh báo trước khi
                                               return None)
                            format_result(x) - kết quả func trả về ->
                                               list dòng text để ghi log
                          Xem gui.py (CombinedPanel) để có ví dụ.

Không phụ thuộc replace_docx.py hay fix_mathtype_parens.py -- file này
không biết và không cần biết app đang xử lý docx/MTEF gì, chỉ lo khung UI.
"""

import os
import zipfile

from PySide6.QtCore import QSettings, QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
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
# Tên chính thức của app. ORG_NAME/APP_NAME dùng chung giá trị này (app
# cá nhân, không có "tổ chức" riêng) -- quyết định nơi QSettings lưu
# đường dẫn đã chọn trên máy. Đổi giá trị này lần nữa sẽ lại mất đường
# dẫn đã lưu dưới tên cũ, y như lần đổi DocxReplacer -> DocxTools trước.
ORG_NAME = "LTWordTool"
APP_NAME = "LTWordTool"


def _app_settings():
    """QSettings dùng file .ini riêng của ứng dụng thay vì định dạng
    "native" cũ (trên Windows, native = registry). Dùng constructor 4 tham
    số (IniFormat, UserScope, org, app) để Qt tự chọn đường dẫn chuẩn theo
    từng hệ điều hành, vd %APPDATA%\\LTWordTool\\LTWordTool.ini trên
    Windows, ~/.config/LTWordTool/LTWordTool.ini trên Linux."""
    return QSettings(QSettings.IniFormat, QSettings.UserScope, ORG_NAME, APP_NAME)


def save_path_pairs(group_key, pairs):
    """Lưu danh sách cặp (path1, path2) vào QSettings dưới group_key, để
    lần mở app sau đọc lại đúng y hệt bằng load_path_pairs() cùng
    group_key. Không biết/không cần biết ý nghĩa của path1/path2 -- lớp
    gọi hàm này (vd CombinedPanel) tự hiểu path1 là gì, path2 là gì.

    remove(group_key) trước khi ghi: dọn sạch mảng cũ, để nếu danh sách
    mới NGẮN hơn lần lưu trước thì không sót lại cặp thừa ở cuối (chỉ
    truyền size cho beginWriteArray không chắc dọn hết tuỳ backend)."""
    settings = _app_settings()
    settings.remove(group_key)
    settings.beginWriteArray(group_key, len(pairs))
    for i, (first, second) in enumerate(pairs):
        settings.setArrayIndex(i)
        settings.setValue("first", first)
        settings.setValue("second", second)
    settings.endArray()


def load_path_pairs(group_key):
    """Đọc lại danh sách cặp (path1, path2) đã lưu bằng save_path_pairs()
    ở lần chạy trước, dưới đúng group_key. Trả về [] nếu chưa từng lưu
    (group_key lạ, hoặc lần lưu gần nhất là danh sách rỗng)."""
    settings = _app_settings()
    count = settings.beginReadArray(group_key)
    pairs = []
    for i in range(count):
        settings.setArrayIndex(i)
        first = settings.value("first", "")
        second = settings.value("second", "")
        pairs.append((first, second))
    settings.endArray()
    return pairs


class FilePickerRow(QWidget):
    """1 hàng: nhãn + ô đường dẫn + nút chọn file. Ô đường dẫn mặc định
    CHỈ ĐỌC (chỉ đổi được qua nút chọn file/nơi lưu) -- truyền
    editable=True để cho gõ tay trực tiếp vào ô (vd file kết quả, nơi
    người dùng có thể muốn tự đặt tên/sửa lại mà không cần mở hộp thoại
    mỗi lần).

    Nếu có settings_key, đường dẫn được tự nhớ lại (qua QSettings) giữa các
    lần mở ứng dụng - không phải chọn lại từ đầu mỗi lần mở GUI. Với ô
    editable=True, gõ tay xong (rời ô hoặc bấm Enter) cũng tự lưu lại y hệt
    lúc chọn qua hộp thoại, không chỉ lúc chọn qua hộp thoại mới được nhớ.
    """

    def __init__(self, label_text, save_mode=False, default_name="",
                 settings_key="", placeholder="Chưa chọn file...", editable=False):
        super().__init__()
        self.save_mode = save_mode
        self.default_name = default_name
        self.settings_key = settings_key

        label = QLabel(label_text)
        label.setFixedWidth(190)

        self.edit = QLineEdit()
        self.edit.setReadOnly(not editable)
        self.edit.setPlaceholderText(placeholder)
        if editable:
            # Gõ tay xong (rời ô hoặc Enter) cũng tự lưu qua QSettings y hệt
            # lúc chọn qua hộp thoại (xem pick_file() -> save_setting())
            # -- không thì gõ tay xong tắt/mở lại app sẽ mất, chỉ path chọn
            # qua hộp thoại mới được nhớ, gây khó hiểu (2 cách nhập, chỉ 1
            # cách được nhớ).
            self.edit.editingFinished.connect(self.save_setting)

        # Chữ trên nút phải khớp việc nó THẬT SỰ làm -- save_mode mở hộp
        # thoại LƯU (xem pick_file() bên dưới), không phải hộp mở file.
        button = QPushButton("Chọn nơi lưu..." if save_mode else "Chọn file...")
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


def friendly_error_text(exc):
    """Diễn giải vài loại lỗi HỆ THỐNG hay gặp nhất (quyền ghi, file không
    hợp lệ...) sang câu tiếng Việt dễ hiểu cho người dùng cuối, thay vì
    hiện thẳng thông báo kỹ thuật (thường tiếng Anh) của Python/thư viện.

    CHỈ áp dụng cho lỗi hệ thống chung chung -- lỗi NGHIỆP VỤ mà
    replace_docx.py/fix_mathtype_parens.py đã tự raise kèm sẵn câu tiếng
    Việt (vd "Không tìm thấy file: ...", "... không phải file Word .docx
    thật...") thì GIỮ NGUYÊN, không dịch chồng lên -- tránh mất chi tiết
    cụ thể (tên file, lý do) mà chính module đó đã biết rõ hơn ở đây.

    Lỗi không nhận diện được vẫn hiện nguyên str(exc) như cũ -- không giấu
    thông tin, chỉ là chưa dịch được, để còn debug/báo lỗi khi cần."""
    if isinstance(exc, PermissionError):
        return (
            "Không có quyền ghi vào nơi lưu đã chọn -- có thể file kết quả "
            "đang mở trong Word, hoặc thư mục không cho phép ghi. Đóng file "
            "đó lại (nếu đang mở) hoặc chọn nơi lưu khác rồi thử lại."
        )
    if isinstance(exc, zipfile.BadZipFile):
        return (
            "File không đọc được như 1 file .docx hợp lệ -- có thể file bị "
            "hỏng hoặc không phải file Word thật."
        )
    if isinstance(exc, KeyError):
        return (
            "File có vẻ không đúng cấu trúc 1 file .docx chuẩn (thiếu phần "
            "nội dung bên trong) -- có thể đây không phải file Word thật, "
            "dù đuôi file là .docx."
        )
    if isinstance(exc, FileNotFoundError):
        # replace_docx.py đã tự kèm câu tiếng Việt cho A/B/C (vd "Không
        # tìm thấy file: ..."); trường hợp khác (vd thư mục lưu kết quả
        # không tồn tại) thường không có, nên thêm 1 gợi ý chung ở đây,
        # KHÔNG thay hẳn thông báo gốc (giữ nguyên để không mất chi tiết).
        return f"{exc}\n(Kiểm tra lại: file có tồn tại, và thư mục lưu kết quả có thật không.)"
    if isinstance(exc, OSError):
        return f"Lỗi hệ thống khi đọc/ghi file: {exc}"

    return str(exc)


class CallableWorker(QThread):
    """Chạy func(*args) ở luồng riêng để giao diện không bị đơ.
    Dùng chung cho MỌI công cụ -- không cần viết 1 class QThread riêng
    cho từng công cụ như trước (ReplaceWorker/MathTypeFixWorker cũ)."""

    finished_ok = Signal(object)   # kết quả func trả về, kiểu gì cũng được
    failed = Signal(str)

    def __init__(self, func, *args):
        super().__init__()
        self.func = func
        self.args = args

    def run(self):
        try:
            result = self.func(*self.args)
        except Exception as e:  # noqa: BLE001 - hiển thị mọi lỗi ra giao diện
            self.failed.emit(friendly_error_text(e))
        else:
            self.finished_ok.emit(result)


class RunPanel(QWidget):
    """Khung sườn dùng chung cho 1 panel công cụ: khu vực chọn file (từ
    file_rows, không viền), nút chạy + nút 'Mở file kết quả', ô nhật ký.
    Lo hết phần luồng nền (CallableWorker), bật/tắt nút, và bắt lỗi hiện
    QMessageBox.

    Kế thừa lớp này, cung cấp file_rows/result_row, rồi viết đúng 2 hàm
    collect_call() và format_result() -- xem gui.py (CombinedPanel)."""

    def __init__(self, run_label, file_rows, result_row, log_height=150):
        """run_label: chữ trên nút chạy.
        file_rows: danh sách FilePickerRow của panel, theo đúng thứ tự hiển thị.
        result_row: FilePickerRow giữ đường dẫn file kết quả (để nút 'Mở
        file kết quả' biết mở gì)."""
        super().__init__()
        self.worker = None
        self.file_rows = file_rows
        self.result_row = result_row

        # Không bọc QGroupBox (không viền) -- chỉ 1 layout dọc thường, để
        # giao diện đơn giản, phẳng. file_rows nhận widget bất kỳ (không
        # chỉ FilePickerRow) nên lớp con có thể chèn thêm checkbox/ghi chú
        # xen kẽ vào đây, xem CombinedPanel trong gui.py làm ví dụ.
        pick_layout = QVBoxLayout()
        pick_layout.setContentsMargins(0, 0, 0, 6)
        pick_layout.setSpacing(6)
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
        layout.addLayout(pick_layout)
        layout.addLayout(button_row)
        layout.addWidget(QLabel("Nhật ký:"))
        layout.addWidget(self.log)

    # ---- lớp con override khi cần -------------------------------------

    def collect_call(self):
        """PHẢI override. Kiểm tra input hiện tại, trả về (func, args) để
        chạy nền -- args là tuple. Nếu input thiếu/sai: tự hiện QMessageBox
        cảnh báo (dùng self làm parent) rồi return None để huỷ, KHÔNG được
        raise."""
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
        func, args = call

        self.run_button.setEnabled(False)
        self.open_button.setEnabled(False)
        self.log_msg("Đang xử lý...")

        self.worker = CallableWorker(func, *args)
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
        path = self.result_row.path()
        if path and os.path.isfile(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
