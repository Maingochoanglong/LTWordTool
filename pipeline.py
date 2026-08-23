"""
pipeline.py
===========
Nối các bước xử lý (replace_docx.py + fix_mathtype_parens.py) chạy TUẦN
TỰ trên cùng 1 file, thành 1 lượt chạy duy nhất với 1 file kết quả cuối
-- để gui.py có đúng 1 hàm để gọi (run_combined()) dù người dùng tích
bao nhiêu thao tác, không phải tự lo phối các bước lại với nhau trong
lớp GUI.

Không viết lại xử lý docx/MTEF ở đây -- chỉ gọi lại replace_docx() và
fix_mathtype_parens_in_docx() của các module đó, đúng như gui.py vẫn
làm. "Thay thế nội dung" nhận NHIỀU cặp (tìm, thay) chạy tuần tự (xem
replace_docx() trong replace_docx.py), không chỉ 1 cặp như trước.

Thứ tự LUÔN CỐ ĐỊNH khi chạy cả 2 bước: Thay thế nội dung trước (hết
toàn bộ danh sách cặp), Sửa ngoặc MathType sau (để công thức mới chèn
từ các cặp thay thế cũng được sửa ngoặc theo, không chỉ công thức có sẵn
trong file gốc).
"""
from __future__ import annotations

import os
import tempfile
from collections.abc import Callable

from replace_docx import FindReplacePair, replace_docx
from fix_mathtype_parens import FixReport, fix_mathtype_parens_in_docx
from mathtype_refresh import RefreshReport, refresh_equation_previews


def _temp_docx_path() -> str:
    """Đường dẫn 1 file .docx tạm rỗng, dùng làm nơi ghi kết quả bước
    trung gian khi chạy cả 2 bước. mkstemp tạo sẵn file rỗng (đóng ngay,
    không cần giữ handle) -- replace_docx() sẽ ghi đè lên đó bằng
    zipfile.ZipFile(..., 'w') như ghi 1 file kết quả bình thường."""
    fd, path = tempfile.mkstemp(suffix=".docx", prefix="_buoc_trung_gian_")
    os.close(fd)
    return path


class CombinedReport:
    """Kết quả 1 lượt run_combined() -- gồm 0 đến 2 bước tuỳ người dùng
    tích chọn gì (Thay thế nội dung, Sửa ngoặc MathType), cộng thêm 1
    bước con TỰ ĐỘNG của "Sửa ngoặc MathType" (vẽ lại ảnh xem trước, xem
    refresh_report bên dưới). replace_counts/mathtype_report = None
    nghĩa là bước đó KHÔNG được chạy (không tích chọn), không phải chạy
    nhưng lỗi."""

    def __init__(
        self,
        replace_counts: list[int] | None,
        mathtype_report: FixReport | None,
        refresh_report: RefreshReport | None,
        out_path: str,
    ) -> None:
        """refresh_report: kết quả bước vẽ lại ảnh xem trước MathType
        (RefreshReport của mathtype_refresh.py) -- None nghĩa là bước đó
        KHÔNG chạy, vì do_mathtype=False hoặc vì mathtype_report.fixed
        rỗng (không có công thức nào cần vẽ lại ảnh, xem run_combined())."""
        self.replace_counts = replace_counts
        self.mathtype_report = mathtype_report
        self.refresh_report = refresh_report
        self.out_path = out_path

    def summary_lines(self) -> list[str]:
        """Danh sách dòng text để GUI ghi vào ô nhật ký -- gộp kết quả của
        đúng những bước đã chạy, theo thứ tự đã chạy."""
        lines: list[str] = []
        if self.replace_counts is not None:
            total = sum(self.replace_counts)
            lines.append(
                f"[OK] Thay thế nội dung: đã thay {total} chỗ "
                f"({len(self.replace_counts)} cặp)"
            )
            for i, count in enumerate(self.replace_counts, start=1):
                lines.append(f"    - cặp {i}: {count} chỗ")
        if self.mathtype_report is not None:
            r = self.mathtype_report
            if r.total == 0:
                lines.append("[OK] Sửa ngoặc MathType: không có công thức MathType nào trong file")
            else:
                lines.append(
                    f"[OK] Sửa ngoặc MathType: {len(r.fixed)}/{r.total} công thức "
                    f"({r.total_pairs} cặp ngoặc), bỏ qua {len(r.skipped)}"
                )
                for name, err in r.skipped:
                    lines.append(f"[BỎ QUA] {name}: {err}")
            if self.refresh_report is not None:
                for line in self.refresh_report.summary_lines():
                    lines.append(f"    {line}")
        lines.append(f"Đã lưu: {self.out_path}")
        return lines


def run_combined(
    source_path: str,
    do_replace: bool,
    pairs: list[FindReplacePair] | None,
    do_mathtype: bool,
    out_path: str,
    log: Callable[[str], None] | None = None,
) -> CombinedReport:
    """Chạy tuần tự trên source_path, LUÔN theo thứ tự: Thay thế nội dung
    trước (nếu do_replace -- chạy hết toàn bộ danh sách pairs tuần tự,
    xem replace_docx() trong replace_docx.py), Sửa ngoặc MathType sau
    (nếu do_mathtype). Ghi kết quả CUỐI CÙNG ra out_path; bước trung gian
    (khi chạy cả 2) dùng file tạm rồi xoá ngay sau khi xong (kể cả khi
    bước sau lỗi) -- không để rác lại. Nếu chỉ tích 1 bước, hàm tương
    ứng ghi thẳng ra out_path, không qua file tạm (giống hệt hành vi cũ
    khi mỗi công cụ còn chạy độc lập).

    pairs: danh sách (find_path, replacement_path, backward_stop_text,
    forward_stop_text) cho bước thay thế nội dung -- chỉ dùng khi
    do_replace=True (truyền None hoặc [] khi do_replace=False). Hai chuỗi
    dừng đi RIÊNG theo từng bộ: một chuỗi chặn lan LÙI, chuỗi còn lại chặn
    lan TIẾN; để trống thì lan tới đầu/cuối đoạn. Xem docstring
    replace_docx() để biết đầy đủ.

    SAU KHI sửa ngoặc MathType xong, NẾU có ít nhất 1 công thức thực sự
    đổi ngoặc (mathtype_report.fixed khác rỗng): tự động mở Word (COM,
    chạy ẩn) để vẽ lại ảnh xem trước cho ĐÚNG những công thức đó (xem
    refresh_equation_previews() trong mathtype_refresh.py) -- không hỏi
    lại, vì đây là phần tự nhiên của việc "sửa ngoặc" (ngoặc đã đổi mà
    ảnh xem trước còn hiện dạng cũ thì coi như chưa xong việc). Bước này
    KHÔNG BAO GIỜ làm hỏng hay xoá mất out_path đã ghi trước đó -- lỗi ở
    bước này (thiếu pywin32, không có Word/MathType, COM lỗi...) chỉ được
    ghi lại trong log + refresh_report, không raise (xem docstring
    refresh_equation_previews()).

    log(text): callback ghi 1 dòng log NGAY tại thời điểm xảy ra, để GUI
    hiện tiến độ thời gian thực thay vì đợi xong mới thấy 1 cục tổng kết
    -- mặc định no-op nếu không truyền."""
    log = log or (lambda _msg: None)
    replace_counts: list[int] | None = None
    mathtype_report: FixReport | None = None
    refresh_report: RefreshReport | None = None
    tmp_path: str | None = None
    try:
        current = source_path

        if do_replace:
            log(f"Đang thay thế nội dung ({len(pairs)} cặp)...")
            target = _temp_docx_path() if do_mathtype else out_path
            replace_counts = replace_docx(current, pairs, target, log=log)
            log(f"[OK] Đã thay thế xong: {sum(replace_counts)} chỗ.")
            if do_mathtype:
                tmp_path = target
            current = target

        if do_mathtype:
            log("Đang sửa ngoặc MathType...")
            mathtype_report = fix_mathtype_parens_in_docx(current, out_path, log=log)

            fixed_names = [name for name, _n_repl in mathtype_report.fixed]
            if fixed_names:
                log("Đang mở Word để vẽ lại ảnh xem trước cho công thức vừa sửa...")
                refresh_report = refresh_equation_previews(out_path, fixed_names, log=log)

        return CombinedReport(replace_counts, mathtype_report, refresh_report, out_path)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
