#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui.py
======
Giao diện Qt6 (PySide6) cho 1 cửa sổ duy nhất xử lý file Word: chọn 1
file gốc, tích chọn thao tác cần làm, bấm 1 nút để chạy -- không cần
dòng lệnh, không chia tab, không tách 2 công cụ riêng nữa:
 
   [ ] Thay thế nội dung: thêm 1 hay NHIỀU cặp (nội dung cần tìm, nội dung
       dùng để thay, chuỗi dừng lan định dạng lùi/tiến RIÊNG của cặp đó)
      vào danh sách -- khu vực nhập cặp + danh sách LUÔN hiện sẵn, chỉ khoá
      (mờ đi, không thao tác được) tới khi tích mục này. Nhiều cặp chạy
      TUẦN TỰ theo đúng thứ tự trong danh sách (kéo-thả để đổi thứ tự) --
      xem docstring replace_docx() trong replace_docx.py để biết đầy đủ
      nguyên tắc so khớp/thay thế/lan định dạng.
  [ ] Sửa ngoặc MathType: chuyển ngoặc tròn cứng trong công thức MathType
      thành ngoặc tự co giãn, chạy thẳng trên file đang có. Xem docstring
      fix_mathtype_parens_in_docx() trong fix_mathtype_parens.py.
 
Tích cả 2: chạy TUẦN TỰ theo đúng thứ tự liệt kê ở trên -- xem
pipeline.py (run_combined()) để biết chi tiết cách nối 2 bước.
 
File này CHỈ còn phần "lắp ráp" (chọn ô nào, gọi hàm nào, hiện log ra
sao) -- khung dùng chung nằm ở panel.py; phần nối 2 bước lại với nhau
nằm ở pipeline.py; xử lý docx/MTEF thật sự nằm ở replace_docx.py và
fix_mathtype_parens.py -- không viết lại ở đây.
 
Style: cố tình không dùng QGroupBox/viền bao quanh khu vực chọn file
(xem panel.py) -- phân biệt các phần bằng khoảng cách + thụt lề nhẹ
(_indented bên dưới) thay vì đường viền, cho gọn mắt.
 
Yêu cầu: pip install PySide6 lxml olefile
Chạy:    python gui.py
"""
 
import os
import sys
 
from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QHBoxLayout,
    QLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
 
from panel import (
    FilePickerRow,
    RunPanel,
    load_checkbox_state,
    load_path_pairs,
    save_checkbox_state,
    save_path_pairs,
)
from pipeline import run_combined
 
 
_ROW_LABEL_WIDTH = 260
 
 
def _indented(*widgets):
    """Gói nhiều widget vào 1 khối thụt lề nhẹ, KHÔNG viền -- dùng cho các
    trường chỉ hiện khi 1 checkbox liên quan được tích (vd danh sách cặp
    tìm/thay của 'Thay thế nội dung')."""
    box = QWidget()
    box_layout = QVBoxLayout(box)
    box_layout.setContentsMargins(20, 2, 0, 4)
    box_layout.setSpacing(4)
    for w in widgets:
        box_layout.addWidget(w)
    return box
 
 
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
 
 
def _field_row(label_text, widget):
    """1 hàng: nhãn (bề rộng cố định, khớp FilePickerRow ở trên) + 1
    widget bất kỳ (vd QLineEdit) -- để layout vẫn thẳng hàng với các
    FilePickerRow phía trên dù không phải ô chọn file."""
    row = QWidget()
    label = QLabel(label_text)
    label.setFixedWidth(_ROW_LABEL_WIDTH)
    layout = QHBoxLayout(row)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(label)
    layout.addWidget(widget, 1)
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
    # Key lưu TRẠNG THÁI TÍCH của 2 checkbox (qua
    # save_checkbox_state()/load_checkbox_state() ở panel.py) -- để lần
    # mở app sau giữ nguyên đúng những gì đã để lại lần trước, không phải
    # tích lại từ đầu. Hai chuỗi dừng lan định dạng đi RIÊNG theo từng
    # cặp, lưu CHUNG với find_path/replacement_path trong PAIRS_SETTINGS_KEY
    # (xem _add_pair_item) -- không còn key riêng của nó nữa.
    CHK_REPLACE_KEY = "checked_replace"
    CHK_MATHTYPE_KEY = "checked_mathtype"
 
    def __init__(self):
        self.row_source = FilePickerRow("File gốc:", settings_key="path_source")
 
        self.chk_replace = QCheckBox("Thay thế nội dung (theo danh sách tìm/thay)")
        self.chk_replace.setChecked(load_checkbox_state(self.CHK_REPLACE_KEY, False))
 
        # 4 ô nhập 1 cặp tại 1 thời điểm -- không dùng settings_key/
        # QSettings riêng ở đây vì đây chỉ là chỗ NHẬP TẠM trước khi bấm
        # "Thêm vào danh sách"; thứ thật sự cần nhớ giữa các lần mở app là
        # danh sách cặp bên dưới (xem pair_list, đã gồm cả
        # chuỗi dừng của từng cặp), không phải nội dung đang gõ dở.
        self.row_find = FilePickerRow("Nội dung cần tìm:")
        self.row_replacement = FilePickerRow("Nội dung dùng để thay:")
        self.edit_backward_stop = QLineEdit()
        self.edit_backward_stop.setPlaceholderText("Trống = lan tới đầu đoạn")
        self.row_backward_stop = _field_row(
            "Định dạng lan VỀ TRƯỚC tới khi gặp:", self.edit_backward_stop
        )
        self.edit_forward_stop = QLineEdit()
        self.edit_forward_stop.setPlaceholderText("Trống = lan tới cuối đoạn")
        self.row_forward_stop = _field_row(
            "Định dạng lan VỀ SAU tới khi gặp:", self.edit_forward_stop
        )
 
        self.btn_add_pair = QPushButton("+ Thêm vào danh sách")
        self.btn_add_pair.clicked.connect(self._on_add_pair)
        self.btn_remove_pair = QPushButton("- Xoá cặp đã chọn")
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
        # cặp đã chọn" -- tức đúng lúc cần nhìn rõ nhất để biết đang xoá
        # cặp nào lại là lúc khó thấy nhất nếu không ép màu ở đây.
        self.pair_list.setStyleSheet(
            "QListWidget::item:selected { background-color: #3874d6; color: white; }"
        )
        # Kéo-thả đổi thứ tự NGAY TRONG danh sách -- thứ tự hiển thị chính
        # là thứ tự sẽ chạy tuần tự (xem replace_docx() trong
        # replace_docx.py), không cần đánh số thủ công trong nhãn từng
        # dòng (số thứ tự sẽ lệch ngay khi kéo-thả nếu có).
        self.pair_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.pair_list.model().rowsMoved.connect(lambda *_args: self._save_pairs())
 
        for find_path, replacement_path, backward_stop_text, forward_stop_text in load_path_pairs(
            self.PAIRS_SETTINGS_KEY, n_fields=4, missing_value=None
        ):
            # Cặp cũ chỉ có 1 chuỗi dừng: duy trì đúng hành vi cũ bằng cách
            # dùng nó cho cả hai chiều. Cặp mới đã lưu fourth="" thì giữ
            # nguyên chuỗi tiến rỗng đó, không tự điền lại.
            if forward_stop_text is None:
                forward_stop_text = backward_stop_text
            self._add_pair_item(
                find_path, replacement_path, backward_stop_text, forward_stop_text
            )
 
        self.replace_fields = _indented(
            self.row_find,
            self.row_replacement,
            self.row_backward_stop,
            self.row_forward_stop,
            _button_row(self.btn_add_pair, self.btn_remove_pair),
            self.pair_list,
        )
 
        self.chk_mathtype = QCheckBox("Sửa ngoặc MathType (ngoặc cứng → tự co giãn)")
        self.chk_mathtype.setChecked(load_checkbox_state(self.CHK_MATHTYPE_KEY, False))
 
        self.row_out = FilePickerRow(
            "File kết quả:", save_mode=True, default_name="ket_qua.docx", settings_key="path_out",
            placeholder="Bỏ trống = tự đặt tên cạnh file gốc",
            editable=True,
        )
 
        # Trạng thái ban đầu: khoá (không ẩn) phần phụ thuộc "Thay thế nội
        # dung" bằng setEnabled() theo ĐÚNG giá trị vừa nạp ở trên (có thể
        # đã tích sẵn từ lần trước) -- trường này LUÔN hiện sẵn, chỉ mờ
        # đi/không thao tác được khi chưa tích, không ẩn hẳn. Đặt trực
        # tiếp lên widget con (không đụng self) vì self chưa phải QWidget
        # hợp lệ tới khi super().__init__() chạy bên dưới.
        self.replace_fields.setEnabled(self.chk_replace.isChecked())
 
        # Nối tín hiệu ở đây an toàn (chỉ ĐĂNG KÝ, chưa GỌI) -- tới lúc
        # người dùng thật sự tích checkbox thì self đã dựng xong từ lâu.
        self.chk_replace.toggled.connect(self._on_feature_toggled)
        self.chk_mathtype.toggled.connect(self._on_feature_toggled)
 
        super().__init__(
            run_label="Thực hiện",
            file_rows=[
                self.row_source,
                self.chk_replace, self.replace_fields,
                self.chk_mathtype,
                self.row_out,
            ],
            result_row=self.row_out,
            log_height=200,
        )
 
        # Đồng bộ chữ trên nút Chạy theo ĐÚNG trạng thái đã nạp ở trên
        # (có thể đã tích sẵn từ lần chạy trước) -- phải gọi SAU
        # super().__init__() vì cần self.run_button đã tồn tại. Cũng lưu
        # lại state qua QSettings (idempotent, vô hại nếu trùng giá trị
        # vừa nạp) để 2 checkbox luôn đồng bộ 1 chỗ duy nhất.
        self._on_feature_toggled()
 
    # ---- danh sách cặp tìm/thay -----------------------------------------
 
    def _add_pair_item(
        self, find_path, replacement_path, backward_stop_text, forward_stop_text
    ):
        """Thêm 1 dòng vào pair_list hiển thị cặp (find_path,
        replacement_path, backward_stop_text, forward_stop_text) -- dùng
        chung cho lúc người dùng bấm 'Thêm vào danh sách' lẫn lúc nạp lại
        danh sách đã lưu từ lần chạy trước, tránh viết code tạo dòng 2 lần
        ở 2 chỗ."""
        label = f"{os.path.basename(find_path)} → {os.path.basename(replacement_path)}"
        if backward_stop_text:
            label += f'  [lùi: "{backward_stop_text}"]'
        if forward_stop_text:
            label += f'  [tiến: "{forward_stop_text}"]'
        item = QListWidgetItem(label)
        item.setData(
            Qt.UserRole,
            (find_path, replacement_path, backward_stop_text, forward_stop_text),
        )
        self.pair_list.addItem(item)
 
    def _current_pairs(self):
        """Danh sách (find_path, replacement_path, backward_stop_text,
        forward_stop_text)
        hiện có trong pair_list, ĐÚNG THEO THỨ TỰ hiển thị (đã kéo-thả sắp
        xếp nếu có) -- đây cũng là thứ tự sẽ chạy tuần tự khi bấm nút
        Thực hiện."""
        return [self.pair_list.item(i).data(Qt.UserRole) for i in range(self.pair_list.count())]
 
    def _save_pairs(self):
        save_path_pairs(self.PAIRS_SETTINGS_KEY, self._current_pairs())
 
    def _on_add_pair(self):
        find_path = self.row_find.path()
        replacement_path = self.row_replacement.path()
        backward_stop_text = self.edit_backward_stop.text()
        forward_stop_text = self.edit_forward_stop.text()
        if not find_path or not replacement_path:
            return self._warn(
                "Vui lòng chọn đủ file cần tìm và file dùng để thay trước khi thêm vào danh sách."
            )
 
        self._add_pair_item(
            find_path, replacement_path, backward_stop_text, forward_stop_text
        )
        self._save_pairs()
        # Xoá trắng cả 4 ô để nhập cặp kế tiếp ngay, khỏi phải tự xoá tay.
        self.row_find.edit.clear()
        self.row_replacement.edit.clear()
        self.edit_backward_stop.clear()
        self.edit_forward_stop.clear()
 
    def _on_remove_pair(self):
        row = self.pair_list.currentRow()
        if row < 0:
            return self._warn("Vui lòng chọn 1 cặp trong danh sách để xoá.")
        self.pair_list.takeItem(row)
        self._save_pairs()
 
    # ---- phần còn lại -----------------------------------------------------
 
    def _on_feature_toggled(self, _checked=None):
        """1 chỗ duy nhất quyết định phần nào BẬT/KHOÁ (setEnabled, KHÔNG
        ẩn) theo 2 checkbox -- gọi lại mỗi khi 1 trong 2 checkbox đổi
        trạng thái (và 1 lần lúc khởi tạo, xem cuối __init__). Vì không
        còn widget nào bị ẩn/hiện, kích thước layout không đổi theo
        checkbox -- không cần resize cửa sổ thủ công."""
        replace_on = self.chk_replace.isChecked()
        math_on = self.chk_mathtype.isChecked()
        save_checkbox_state(self.CHK_REPLACE_KEY, replace_on)
        save_checkbox_state(self.CHK_MATHTYPE_KEY, math_on)
        self.replace_fields.setEnabled(replace_on)
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
        layout = QVBoxLayout(self)
        # Khung Replace luôn hiện (chỉ bị khoá/mở theo checkbox), nên cửa
        # sổ không được thu nhỏ tới mức layout ép mất các ô chọn file hoặc
        # danh sách cặp. Người dùng vẫn có thể phóng to bình thường.
        layout.setSizeConstraint(QLayout.SetMinimumSize)
        layout.addWidget(CombinedPanel())
        self.resize(self.sizeHint().expandedTo(QSize(620, 400)))
        self._center_on_primary_screen()
 
    def _center_on_primary_screen(self):
        """Đặt cửa sổ vào giữa vùng hiển thị khả dụng của màn hình chính
        (không che taskbar) mỗi lần app được tạo."""
        screen = QApplication.primaryScreen()
        if screen is None:
            return
        frame = self.frameGeometry()
        frame.moveCenter(screen.availableGeometry().center())
        self.move(frame.topLeft())
 
 
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
 
 
if __name__ == "__main__":
    main()