#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui.py
======
Giao diện Qt6 (PySide6) cho 1 cửa sổ duy nhất xử lý file Word: chọn 1
file gốc, tích chọn thao tác cần làm, bấm 1 nút để chạy -- không cần
dòng lệnh, không chia tab, không tách 2 công cụ riêng nữa. Mọi khu vực
nhập liệu LUÔN HIỆN SẴN (không ẩn/hiện theo checkbox) -- 2 checkbox dưới
đây chỉ quyết định bước đó có THỰC SỰ CHẠY hay không khi bấm nút, không
quyết định phần nào hiện/ẩn trên màn hình:

  [ ] Thay thế nội dung: thêm 1 hay NHIỀU cặp (nội dung cần tìm, nội dung
      dùng để thay) vào danh sách. Nhiều cặp chạy TUẦN TỰ theo đúng thứ
      tự trong danh sách (kéo-thả để đổi thứ tự). Xem docstring
      replace_docx() trong replace_docx.py.
  [ ] Sửa ngoặc MathType: chuyển ngoặc tròn cứng trong công thức MathType
      thành ngoặc tự co giãn, chạy thẳng trên file đang có. Xem docstring
      fix_mathtype_parens_in_docx() trong fix_mathtype_parens.py.

Tích cả 2: chạy TUẦN TỰ theo đúng thứ tự liệt kê ở trên (thay thế trước
-- hết toàn bộ danh sách cặp -- sửa ngoặc sau, để công thức mới chèn từ
các cặp thay thế cũng được sửa ngoặc theo) -- xem pipeline.py
(run_combined()) để biết chi tiết cách nối 2 bước.

File này CHỈ còn phần "lắp ráp" (chọn ô nào, gọi hàm nào, hiện log ra
sao) -- khung dùng chung (nút chạy, nút mở kết quả, luồng nền, ô nhật
ký, nhớ đường dẫn qua QSettings...) nằm ở panel.py; phần nối 2 bước
lại với nhau nằm ở pipeline.py; xử lý docx/MTEF thật sự nằm ở
replace_docx.py và fix_mathtype_parens.py -- không viết lại ở đây.

Style: cố tình không dùng QGroupBox/viền bao quanh khu vực chọn file
(xem panel.py) -- phân biệt các phần bằng khoảng cách + thụt lề nhẹ
(_indented/_note_label bên dưới) thay vì đường viền, cho gọn mắt.

Yêu cầu: pip install PySide6 lxml olefile
Chạy:    python gui.py
"""

import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from panel import FilePickerRow, RunPanel, load_path_pairs, save_path_pairs
from pipeline import run_combined


def _indented(*widgets):
    """Gói nhiều widget vào 1 khối thụt lề nhẹ, KHÔNG viền -- dùng để phân
    nhóm trực quan các trường thuộc về 1 checkbox liên quan (vd danh sách
    cặp tìm/thay thuộc về 'Thay thế nội dung'), LUÔN hiện sẵn cùng lúc chứ
    không ẩn/hiện theo checkbox -- thụt lề chỉ để mắt dễ thấy nhóm nào
    thuộc về mục nào, không phải cơ chế hiện/ẩn."""
    box = QWidget()
    box_layout = QVBoxLayout(box)
    box_layout.setContentsMargins(20, 2, 0, 4)
    box_layout.setSpacing(4)
    for w in widgets:
        box_layout.addWidget(w)
    return box


def _note_label(text):
    """1 dòng ghi chú nhỏ, màu xám, thụt lề theo nhóm liên quan -- không
    viền, không QGroupBox."""
    note = QLabel(text)
    note.setWordWrap(True)
    note.setStyleSheet("color: #666; font-size: 11px; margin-left: 20px;")
    return note


def _button_row(*buttons):
    """Gói nhiều QPushButton vào 1 hàng ngang, làm 1 widget duy nhất --
    để có thể đưa vào _indented()/file_rows (chỉ nhận widget, không nhận
    layout trực tiếp)."""
    row = QWidget()
    layout = QHBoxLayout(row)
    layout.setContentsMargins(0, 0, 0, 0)
    for b in buttons:
        layout.addWidget(b)
    return row


class CombinedPanel(RunPanel):
    """Panel duy nhất của app: 1 file gốc + các checkbox thao tác.
    collect_call() luôn trả về đúng 1 lệnh gọi run_combined() (pipeline.py)
    dù người dùng tích 1 hay cả 2 thao tác, và dù danh sách cặp tìm/thay
    có 1 hay nhiều cặp -- CallableWorker dùng chung ở panel.py không cần
    biết bên trong chạy mấy bước/mấy cặp, nó chỉ thấy 1 hàm, 1 kết quả,
    y hệt mọi công cụ khác."""

    # Group key để save_path_pairs()/load_path_pairs() (panel.py) lưu
    # danh sách cặp tìm/thay riêng, không lẫn với các settings_key khác
    # (path_source, path_out...) trong cùng file LTWordTool.ini.
    PAIRS_SETTINGS_KEY = "pairs_replace"

    def __init__(self):
        self.row_source = FilePickerRow("File gốc:", settings_key="path_source")

        self.chk_replace = QCheckBox("Thay thế nội dung (theo danh sách tìm/thay)")

        # 2 ô nhập 1 cặp tại 1 thời điểm -- không dùng settings_key ở đây
        # vì đây chỉ là chỗ NHẬP TẠM trước khi bấm "Thêm vào danh sách";
        # thứ thật sự cần nhớ giữa các lần mở app là danh sách cặp bên
        # dưới (xem pair_list), không phải nội dung đang gõ dở.
        self.row_find = FilePickerRow("Nội dung cần tìm:")
        self.row_replacement = FilePickerRow("Nội dung dùng để thay:")

        self.btn_add_pair = QPushButton("+ Thêm vào danh sách")
        self.btn_add_pair.clicked.connect(self._on_add_pair)
        self.btn_remove_pair = QPushButton("- Xoá dòng đã chọn")
        self.btn_remove_pair.clicked.connect(self._on_remove_pair)

        # Chiều cao CỐ ĐỊNH (không phình theo số cặp đã thêm) -- tự cuộn
        # nội bộ khi danh sách dài hơn chỗ hiện, giữ cửa sổ chính luôn
        # gọn dù người dùng thêm bao nhiêu cặp.
        self.pair_list = QListWidget()
        self.pair_list.setFixedHeight(90)
        self.pair_list.setSelectionMode(QAbstractItemView.SingleSelection)
        # Ép màu tô khi CHỌN 1 dòng luôn rõ, không phụ thuộc pair_list có
        # đang giữ focus hay không -- mặc định Qt (rõ nhất trên style
        # Windows) làm NHẠT hẳn màu tô khi widget MẤT focus, mà người dùng
        # mất focus khỏi pair_list đúng vào lúc di chuột sang bấm nút "Xoá
        # dòng đã chọn" -- tức đúng lúc cần nhìn rõ nhất để biết đang xoá
        # dòng nào lại là lúc khó thấy nhất nếu không ép màu ở đây.
        self.pair_list.setStyleSheet(
            "QListWidget::item:selected { background-color: #3874d6; color: white; }"
        )
        # Kéo-thả đổi thứ tự NGAY TRONG danh sách -- thứ tự hiển thị chính
        # là thứ tự sẽ chạy tuần tự (xem replace_docx() trong
        # replace_docx.py), không cần đánh số thủ công trong nhãn từng
        # dòng (số thứ tự sẽ lệch ngay khi kéo-thả nếu có).
        self.pair_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.pair_list.model().rowsMoved.connect(lambda *_args: self._save_pairs())

        for find_path, replacement_path in load_path_pairs(self.PAIRS_SETTINGS_KEY):
            self._add_pair_item(find_path, replacement_path)

        self.replace_fields = _indented(
            self.row_find,
            self.row_replacement,
            _button_row(self.btn_add_pair, self.btn_remove_pair),
            self.pair_list,
        )

        self.chk_mathtype = QCheckBox("Sửa ngoặc MathType (ngoặc cứng → tự co giãn)")
        self.mathtype_note = _note_label(
            "Lưu ý: sau khi sửa, cần double-click từng công thức trong Word "
            "thật thì ảnh xem trước mới vẽ lại đúng ngoặc co giãn."
        )

        self.order_note = _note_label(
            "Tích cả 2: sẽ thay thế nội dung trước (hết toàn bộ danh sách "
            "cặp), rồi sửa ngoặc MathType trên kết quả đó (để công thức "
            "mới chèn từ các cặp thay thế cũng được sửa)."
        )

        self.row_out = FilePickerRow(
            "File kết quả:", save_mode=True, default_name="ket_qua.docx", settings_key="path_out",
            placeholder="Bỏ trống = tự đặt tên cạnh file gốc",
            editable=True,
        )

        # Nối tín hiệu ở đây an toàn (chỉ ĐĂNG KÝ, chưa GỌI) -- tới lúc
        # người dùng thật sự tích checkbox thì self đã dựng xong từ lâu.
        self.chk_replace.toggled.connect(self._on_feature_toggled)
        self.chk_mathtype.toggled.connect(self._on_feature_toggled)

        super().__init__(
            run_label="Thực hiện",
            file_rows=[
                self.row_source,
                self.chk_replace, self.replace_fields,
                self.chk_mathtype, self.mathtype_note, self.order_note,
                self.row_out,
            ],
            result_row=self.row_out,
            log_height=200,
        )

    # ---- danh sách cặp tìm/thay -----------------------------------------

    def _add_pair_item(self, find_path, replacement_path):
        """Thêm 1 dòng vào pair_list hiển thị cặp (find_path,
        replacement_path) -- dùng chung cho lúc người dùng bấm 'Thêm vào
        danh sách' lẫn lúc nạp lại danh sách đã lưu từ lần chạy trước,
        tránh viết code tạo dòng 2 lần ở 2 chỗ."""
        label = f"{os.path.basename(find_path)} → {os.path.basename(replacement_path)}"
        item = QListWidgetItem(label)
        item.setData(Qt.UserRole, (find_path, replacement_path))
        self.pair_list.addItem(item)

    def _current_pairs(self):
        """Danh sách (find_path, replacement_path) hiện có trong pair_list,
        ĐÚNG THEO THỨ TỰ hiển thị (đã kéo-thả sắp xếp nếu có) -- đây cũng
        là thứ tự sẽ chạy tuần tự khi bấm nút Thực hiện."""
        return [self.pair_list.item(i).data(Qt.UserRole) for i in range(self.pair_list.count())]

    def _save_pairs(self):
        save_path_pairs(self.PAIRS_SETTINGS_KEY, self._current_pairs())

    def _on_add_pair(self):
        find_path = self.row_find.path()
        replacement_path = self.row_replacement.path()
        if not find_path or not replacement_path:
            return self._warn(
                "Vui lòng chọn đủ file cần tìm và file dùng để thay trước khi thêm vào danh sách."
            )

        self._add_pair_item(find_path, replacement_path)
        self._save_pairs()
        # Xoá trắng 2 ô để nhập cặp kế tiếp ngay, khỏi phải tự xoá tay.
        self.row_find.edit.clear()
        self.row_replacement.edit.clear()

    def _on_remove_pair(self):
        row = self.pair_list.currentRow()
        if row < 0:
            return self._warn("Vui lòng chọn 1 dòng trong danh sách để xoá.")
        self.pair_list.takeItem(row)
        self._save_pairs()

    # ---- phần còn lại -----------------------------------------------------

    def _on_feature_toggled(self, _checked=None):
        """Danh sách cặp, ghi chú MathType, ghi chú thứ tự chạy LUÔN hiện
        sẵn (không ẩn/hiện theo checkbox nữa) -- 2 checkbox giờ CHỈ quyết
        định bước đó có THỰC SỰ CHẠY hay không khi bấm nút, không còn
        quyết định phần nào hiện/ẩn. Vì vậy hàm này giờ chỉ còn 1 việc: đổi
        chữ trên nút chạy cho khớp đúng thao tác sẽ chạy."""
        replace_on = self.chk_replace.isChecked()
        math_on = self.chk_mathtype.isChecked()
        # Tên nút chạy đổi theo đúng thao tác đang tích -- 2 công cụ cũ có
        # tên nút rõ việc ("Thực hiện thay thế"/"Sửa ngoặc MathType"), gộp
        # lại vẫn muốn giữ độ rõ đó thay vì 1 chữ "Thực hiện" chung chung.
        self.run_button.setText(self._run_button_label(replace_on, math_on))

    @staticmethod
    def _run_button_label(replace_on, math_on):
        if replace_on and math_on:
            return "Thực hiện cả 2 bước"
        if replace_on:
            return "Thay thế nội dung"
        if math_on:
            return "Sửa ngoặc MathType"
        return "Thực hiện"

    def _warn(self, message):
        """Hiện cảnh báo 'Thiếu thông tin' rồi trả None -- dùng trực tiếp
        qua 'return self._warn(...)' trong collect_call(), khỏi lặp lại
        title + return None ở mỗi nhánh kiểm tra input."""
        QMessageBox.warning(self, "Thiếu thông tin", message)
        return None

    def collect_call(self):
        source = self.row_source.path()
        do_replace = self.chk_replace.isChecked()
        do_mathtype = self.chk_mathtype.isChecked()
        pairs = self._current_pairs()
        out = self.row_out.path()

        if not source:
            return self._warn("Vui lòng chọn file gốc.")
        if not do_replace and not do_mathtype:
            return self._warn("Vui lòng chọn ít nhất 1 thao tác.")
        if do_replace and not pairs:
            return self._warn(
                'Đã tích "Thay thế nội dung" - vui lòng thêm ít nhất 1 cặp tìm/thay vào danh sách.'
            )

        if not out:
            stem = os.path.splitext(os.path.basename(source))[0]
            out = os.path.join(os.path.dirname(source) or ".", f"{stem}_ket_qua.docx")
            self.row_out.edit.setText(out)
            self.row_out.save_setting()
            self.log_msg(f"(Chưa chọn nơi lưu - tự lưu tại: {out})")

        return run_combined, (source, do_replace, pairs, do_mathtype, out)

    def format_result(self, report):
        return report.summary_lines()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LTWordTool - Công cụ xử lý file Word")
        self.setMinimumSize(620, 400)

        layout = QVBoxLayout(self)
        layout.addWidget(CombinedPanel())

        # Mọi khu vực đều hiện sẵn ngay từ đầu (không còn ẩn/hiện theo
        # checkbox nữa) nên cửa sổ cần đủ cao NGAY LẦN show() ĐẦU TIÊN --
        # không còn cơ hội "giãn dần" theo thao tác người dùng như thiết kế
        # cũ. Gọi resize(sizeHint()) tường minh ở đây thay vì trông chờ Qt
        # tự canh đủ cao khi show() lần đầu: đã kiểm chứng thực nghiệm kích
        # thước mặc định lúc show() lần đầu có thể hụt so với sizeHint()
        # thật sự cần, cắt mất phần dưới cùng (nút chạy, ô nhật ký...).
        self.resize(self.sizeHint())
        self._center_on_screen()

    def _center_on_screen(self):
        """Đặt cửa sổ vào GIỮA màn hình khi mở app. Mặc định (không gọi
        hàm này) vị trí cửa sổ lúc mở tuỳ window manager của hệ điều hành
        quyết định -- có thể lệch góc, không nhất quán giữa các máy.

        Dùng availableGeometry() (không phải geometry() suông) vì nó đã
        trừ sẵn phần bị taskbar/dock chiếm -- tránh cửa sổ bị lệch xuống
        dưới taskbar. Gọi move() TRƯỚC show() (xem __init__) để tránh cửa
        sổ hiện ra ở 1 chỗ rồi "nhảy" sang giữa màn hình ngay sau đó."""
        screen = QApplication.primaryScreen()
        if screen is None:  # phòng hờ môi trường không có màn hình thật (vd chạy test/CI)
            return
        available = screen.availableGeometry()
        x = available.x() + (available.width() - self.width()) // 2
        y = available.y() + (available.height() - self.height()) // 2
        self.move(x, y)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
