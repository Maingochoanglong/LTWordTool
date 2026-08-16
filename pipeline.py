

"""
pipeline.py
===========
Nối 2 bước xử lý (replace_docx.py + fix_mathtype_parens.py) chạy TUẦN TỰ
trên cùng 1 file, thành 1 lượt chạy duy nhất với 1 file kết quả cuối --
để gui.py có đúng 1 hàm để gọi (run_combined()) dù người dùng tích 1 hay
cả 2 thao tác, không phải tự lo phối 2 bước lại với nhau trong lớp GUI.

Không viết lại xử lý docx/MTEF ở đây -- chỉ gọi lại replace_docx() và
fix_mathtype_parens_in_docx() của 2 module đó, đúng như gui.py vẫn làm.
"Thay thế nội dung" nhận NHIỀU cặp (tìm, thay) chạy tuần tự (xem
replace_docx() trong replace_docx.py), không chỉ 1 cặp như trước.

Thứ tự LUÔN CỐ ĐỊNH khi chạy cả 2 bước: Thay thế nội dung trước (hết
toàn bộ danh sách cặp), Sửa ngoặc MathType sau -- để công thức mới chèn
từ các cặp thay thế cũng được sửa ngoặc theo, không chỉ công thức có sẵn
trong file gốc.
"""

import os
import tempfile

from replace_docx import replace_docx
from fix_mathtype_parens import fix_mathtype_parens_in_docx

def _temp_docx_path():
    """Đường dẫn 1 file .docx tạm rỗng, dùng làm nơi ghi kết quả bước
    trung gian khi chạy cả 2 bước. mkstemp tạo sẵn file rỗng (đóng ngay,
    không cần giữ handle) -- replace_docx() sẽ ghi đè lên đó bằng
    zipfile.ZipFile(..., 'w') như ghi 1 file kết quả bình thường."""
    fd, path = tempfile.mkstemp(suffix=".docx", prefix="_buoc_trung_gian_")
    os.close(fd)
    return path

class CombinedReport:
    """Kết quả 1 lượt run_combined() -- gồm 0, 1 hoặc 2 bước tuỳ người
    dùng tích chọn gì. replace_counts/mathtype_report = None nghĩa là bước
    đó KHÔNG được chạy (không tích chọn), không phải chạy nhưng lỗi."""

    def __init__(self, replace_counts, mathtype_report, out_path):
        self.replace_counts = replace_counts
        self.mathtype_report = mathtype_report
        self.out_path = out_path

    def summary_lines(self):
        """Danh sách dòng text để GUI ghi vào ô nhật ký -- gộp kết quả của
        đúng những bước đã chạy, theo thứ tự đã chạy."""
        lines = []
        if self.replace_counts is not None:
            total = sum(self.replace_counts)
            lines.append(
                f"✔ Thay thế nội dung: đã thay {total} chỗ "
                f"({len(self.replace_counts)} cặp)"
            )
            for i, count in enumerate(self.replace_counts, start=1):
                lines.append(f"    - cặp {i}: {count} chỗ")
        if self.mathtype_report is not None:
            r = self.mathtype_report
            if r.total == 0:
                lines.append("✔ Sửa ngoặc MathType: không có công thức MathType nào trong file")
            else:
                lines.append(
                    f"✔ Sửa ngoặc MathType: {len(r.fixed)}/{r.total} công thức "
                    f"({r.total_pairs} cặp ngoặc), bỏ qua {len(r.skipped)}"
                )
                for name, err in r.skipped:
                    lines.append(f"    - bỏ qua {name}: {err}")
        lines.append(f"Đã lưu: {self.out_path}")
        return lines

def run_combined(source_path, do_replace, pairs, do_mathtype, out_path):
    """Chạy tuần tự trên source_path, LUÔN theo thứ tự: Thay thế nội dung
    trước (nếu do_replace -- chạy hết toàn bộ danh sách pairs tuần tự,
    xem replace_docx() trong replace_docx.py) rồi Sửa ngoặc MathType sau
    (nếu do_mathtype). Ghi kết quả CUỐI CÙNG ra out_path; nếu chạy cả 2
    bước, dùng 1 file tạm cho kết quả bước 1 rồi xoá đi ngay sau khi xong
    (kể cả khi bước 2 lỗi) -- không để rác lại. Nếu chỉ tích 1 trong 2,
    hàm tương ứng ghi thẳng ra out_path, không qua file tạm (giống hệt
    hành vi cũ khi mỗi công cụ còn chạy độc lập).

    pairs: danh sách (find_path, replacement_path, backward_stop_text,
    forward_stop_text) cho bước thay thế nội dung -- chỉ dùng khi
    do_replace=True (truyền None hoặc [] khi do_replace=False). Hai chuỗi
    dừng đi RIÊNG theo từng bộ: một chuỗi chặn lan LÙI, chuỗi còn lại chặn
    lan TIẾN; để trống thì lan tới đầu/cuối đoạn. Xem docstring
    replace_docx() để biết đầy đủ."""
    replace_counts = None
    mathtype_report = None
    tmp_path = None
    try:
        current = source_path
        if do_replace:
            target = _temp_docx_path() if do_mathtype else out_path
            replace_counts = replace_docx(current, pairs, target)
            if do_mathtype:
                tmp_path = target
            current = target

        if do_mathtype:
            mathtype_report = fix_mathtype_parens_in_docx(current, out_path)

        return CombinedReport(replace_counts, mathtype_report, out_path)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
