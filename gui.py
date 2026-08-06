#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui.py
======
Giao diện Qt6 (PySide6) cho 2 công cụ xử lý file Word, mỗi công cụ 1 tab
trong cùng 1 cửa sổ:

  1. "Thay thế nội dung" (replace_docx.py): chọn file A, B, C, D bằng hộp
     thoại chọn file, bấm 1 nút để chạy thay thế - không cần dòng lệnh.
  2. "Sửa ngoặc MathType" (fix_mathtype_parens.py): chọn 1 file .docx, bấm
     1 nút để chuyển toàn bộ ngoặc tròn cứng trong công thức MathType
     thành ngoặc tự co giãn.

File này CHỈ còn phần "lắp ráp" riêng của từng công cụ (chọn ô nào, gọi
hàm nào, hiện log ra sao) -- toàn bộ phần khung dùng chung (nút chạy, nút
mở kết quả, luồng nền, ô nhật ký, nhớ đường dẫn qua QSettings...) nằm ở
tool_tab.py. Muốn thêm 1 công cụ mới: viết 1 module logic mới (như
replace_docx.py/fix_mathtype_parens.py), rồi thêm 1 class kế thừa
ToolTab ở đây với đúng 2 hàm collect_call()/format_result() -- xem 2 tab
bên dưới làm ví dụ, không cần đụng vào tool_tab.py.

replace_docx.py và fix_mathtype_parens.py phải nằm cùng thư mục (logic
xử lý không viết lại ở đây - gui.py chỉ gọi lại hàm replace_docx() và
fix_mathtype_parens_in_docx() của 2 file đó).

Yêu cầu: pip install PySide6 lxml olefile
Chạy:    python gui.py
"""

import os
import sys

from PySide6.QtWidgets import QApplication, QLabel, QMessageBox, QTabWidget, QVBoxLayout, QWidget

from tool_tab import FilePickerRow, ToolTab
from replace_docx import replace_docx
from fix_mathtype_parens import fix_mathtype_parens_in_docx


class ReplaceTab(ToolTab):
    """Tab 'Thay thế nội dung': A + B → C = D. Xem docstring replace_docx()
    trong replace_docx.py để biết chi tiết nguyên tắc xử lý."""

    def __init__(self):
        self.row_a = FilePickerRow("File gốc (A):", settings_key="path_a")
        self.row_b = FilePickerRow("Nội dung cần tìm (B):", settings_key="path_b")
        self.row_c = FilePickerRow("Nội dung dùng để thay (C):", settings_key="path_c")
        self.row_d = FilePickerRow(
            "File kết quả (D):", save_mode=True, default_name="D.docx", settings_key="path_d",
            placeholder="Bỏ trống = tự lưu D.docx cạnh file A",
        )
        super().__init__(
            run_label="Thực hiện thay thế",
            file_rows=[self.row_a, self.row_b, self.row_c, self.row_d],
            result_row=self.row_d,
        )

    def collect_call(self):
        a, b, c, d = self.row_a.path(), self.row_b.path(), self.row_c.path(), self.row_d.path()

        missing = [name for name, p in (("A", a), ("B", b), ("C", c)) if not p]
        if missing:
            QMessageBox.warning(
                self, "Thiếu thông tin", f"Vui lòng chọn đủ file: {', '.join(missing)}"
            )
            return None

        if not d:
            d = os.path.join(os.path.dirname(a) or ".", "D.docx")
            self.row_d.edit.setText(d)
            self.row_d.save_setting()
            self.log_msg(f"(Chưa chọn nơi lưu D - tự lưu tại: {d})")

        return replace_docx, (a, b, c, d), {}

    def format_result(self, count):
        return [f"✔ Đã thay {count} chỗ → {self.row_d.path()}"]


class MathTypeFixTab(ToolTab):
    """Tab 'Sửa ngoặc MathType': chuyển ngoặc tròn cứng trong công thức
    MathType của 1 file .docx thành ngoặc tự co giãn. Xem docstring
    fix_mathtype_parens_in_docx() trong fix_mathtype_parens.py để biết
    chi tiết nguyên tắc xử lý."""

    def __init__(self):
        self.row_input = FilePickerRow("File Word cần sửa:", settings_key="mtf_input")
        self.row_output = FilePickerRow(
            "File kết quả:", save_mode=True, default_name="fixed.docx", settings_key="mtf_output",
            placeholder="Bỏ trống = tự lưu <tên gốc>_fixed.docx cạnh file gốc",
        )
        super().__init__(
            run_label="Sửa ngoặc MathType",
            file_rows=[self.row_input, self.row_output],
            result_row=self.row_output,
            log_height=180,
        )

    def extra_widget(self):
        note = QLabel(
            "Lưu ý: ảnh xem trước của công thức trong Word không tự vẽ lại "
            "ngay sau khi sửa (do cơ chế cache của OLE) - cần double-click "
            "từng công thức trong Word thật để thấy ngoặc co giãn đúng."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #666; font-size: 11px;")
        return note

    def collect_call(self):
        input_path = self.row_input.path()
        output_path = self.row_output.path()

        if not input_path:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Word cần sửa.")
            return None

        if not output_path:
            stem = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(os.path.dirname(input_path) or ".", f"{stem}_fixed.docx")
            self.row_output.edit.setText(output_path)
            self.row_output.save_setting()
            self.log_msg(f"(Chưa chọn nơi lưu - tự lưu tại: {output_path})")

        return fix_mathtype_parens_in_docx, (input_path, output_path), {}

    def format_result(self, report):
        return ["✔ Xong:"] + report.summary_lines()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Công cụ xử lý file Word")
        self.setMinimumSize(620, 440)

        tabs = QTabWidget()
        tabs.addTab(ReplaceTab(), "Thay thế nội dung (A+B→C=D)")
        tabs.addTab(MathTypeFixTab(), "Sửa ngoặc MathType")

        layout = QVBoxLayout(self)
        layout.addWidget(tabs)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
