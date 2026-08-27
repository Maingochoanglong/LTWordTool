"""
fix_mathtype_parens.py
=======================
Chuyển ngoặc tự co giãn trong công thức MathType thành ngoặc cứng khi phần
nội dung bên trong chỉ là chữ, số hoặc phép tính đơn giản.

Đây là module orchestration DUY NHẤT mà gui.py/pipeline.py import (qua
hàm fix_mathtype_parens_in_docx() bên dưới) -- phần lõi xử lý nhị phân
được TÁCH thành 3 module con theo domain riêng, mỗi module 1 trách
nhiệm:
  - mtef_parser.py    : đọc (KHÔNG sửa) cấu trúc nhị phân MTEF thành cây Rec.
  - mtef_transform.py : duyệt cây Rec, tìm & chuyển ngoặc tự co giãn "đơn
                        giản" thành ngoặc cứng (transform_mtef()).
  - cfb_builder.py    : đóng gói lại toàn bộ stream (kể cả Equation Native
                        đã sửa) thành 1 container OLE/CFB hợp lệ.
File này chỉ lo phần "ghép nối" các module con đó với thao tác trên file
.docx thật: mở từng oleObjectN.bin bằng olefile (process_one_equation()),
rồi duyệt/nén lại toàn bộ file .docx (giải nén -> sửa từng
oleObject*.bin -> nén lại, xem fix_mathtype_parens_in_docx()) và gom kết
quả thành FixReport để gui.py/pipeline.py hiển thị.

Yêu cầu: pip install olefile

Cách hoạt động: .docx là 1 file zip; mỗi công thức MathType là 1 "OLE
object" nhị phân trong word/embeddings/oleObjectN.bin. Bên trong là cấu
trúc nhị phân MTEF (định dạng riêng của MathType, xem spec chính thức
tại docs.wiris.com/.../mathtype-mtef-v5-mathtype-40-and-later) -- không
có thư viện Python/pip nào đọc VÀ ghi lại được định dạng này (đã tìm),
nên mtef_parser.py/mtef_transform.py (đọc + biến đổi) và cfb_builder.py
(ghi container OLE) là tự viết.

Công thức nào không tự tin đọc đúng 100% cấu trúc nhị phân (đọc lại phải
khớp chính xác từng byte) sẽ được giữ nguyên, không sửa, và liệt kê
trong báo cáo cuối -- an toàn hơn đoán và có thể làm hỏng công thức. Đã
test kỹ bằng LibreOffice, nhưng KHÔNG có MathType/Word thật để test, nên
mở thử bằng Word/MathType thật trên 1 bản sao trước khi tin tưởng hoàn
toàn, đặc biệt với file quan trọng.

Tương thích phiên bản MathType: định dạng nhị phân MTEF v5 (theo đúng
spec Wiris công bố) dùng cho "MathType 4.0 trở lên" -- tức toàn bộ dòng
MathType 4.x/5.x/6.x/7.x/365 hiện có (bản mới nhất hiện tại: 7.11.x),
nhiều hơn hẳn 10 phiên bản. Công thức tạo bởi Equation Editor 3.x hoặc
MathType 3.5 (cũ hơn 4.0) dùng định dạng khác hẳn (MTEF v3/v4) -- module
mtef_parser.py tự nhận diện qua byte phiên bản ghi sẵn trong header của
từng công thức và BỎ QUA an toàn (liệt kê rõ lý do trong báo cáo), không
cố đọc sai.
"""
from __future__ import annotations

import shutil
import tempfile
import uuid
import zipfile
import io
from collections.abc import Callable
from pathlib import Path

import olefile

from cfb_builder import build_cfb
from mtef_parser import MTEFError
from mtef_transform import transform_equation_native


def process_one_equation(bin_path: Path) -> tuple[bytes | None, int, str | None]:
    """Đọc 1 file oleObjectN.bin, thử sửa ngoặc, đóng gói lại. Trả về
    (bytes_kết_quả_hoặc_None, số_cặp_ngoặc_đã_sửa, lý_do_lỗi_hoặc_None):
    bytes_kết_quả là None khi không có gì để sửa (n_repl == 0) HOẶC khi
    có lỗi (lý_do_lỗi khác None) -- nơi gọi (fix_mathtype_parens_in_docx())
    phân biệt 2 trường hợp đó qua lý_do_lỗi. lý_do_lỗi ĐÃ được định dạng
    sẵn để hiển thị trực tiếp (1 câu tiếng Việt chính, có thể kèm thêm 1
    dòng phụ thụt lề "      chi tiết: ..." khi lý do gốc là 1 exception
    ngoài dự kiến) -- nơi gọi chỉ cần in nguyên văn, không cần xử lý gì
    thêm."""
    try:
        ole = olefile.OleFileIO(str(bin_path))
        other_streams = []
        for n in ['\x01CompObj', '\x01Ole', '\x03ObjInfo']:
            other_streams.append((n, ole.openstream(n).read()))
        full = ole.openstream('Equation Native').read()
        root_clsid_str = ole.root.clsid
    except Exception as e:
        return None, 0, (
            "không đọc được dữ liệu bên trong công thức này\n"
            f"      chi tiết: {e}"
        )

    try:
        new_full, n_repl = transform_equation_native(full)
    except MTEFError as e:
        # Đã tiếng Việt sẵn (xem mtef_parser.py/mtef_transform.py) -- dùng
        # thẳng làm dòng chính, KHÔNG bọc thêm tiền tố, KHÔNG thêm dòng phụ.
        return None, 0, str(e)
    except Exception as e:
        return None, 0, (
            "không đọc đúng công thức này, giữ nguyên để an toàn\n"
            f"      chi tiết: {e}"
        )

    if n_repl == 0:
        return None, 0, None

    try:
        root_clsid = uuid.UUID(root_clsid_str).bytes_le if root_clsid_str else b'\x00' * 16
        raw = build_cfb(other_streams + [('Equation Native', new_full)], root_clsid=root_clsid)
        check = olefile.OleFileIO(io.BytesIO(raw))
        got = check.openstream('Equation Native').read()
        got_clsid = check.root.clsid
        check.close()
        if got != new_full:
            return None, 0, (
                "lỗi nội bộ khi đóng gói lại, đã huỷ để an toàn (giữ nguyên bản gốc)\n"
                "      chi tiết: internal error: CFB round-trip mismatch"
            )
        if got_clsid != root_clsid_str:
            return None, 0, (
                "lỗi nội bộ khi đóng gói lại, đã huỷ để an toàn (giữ nguyên bản gốc)\n"
                f"      chi tiết: internal error: CLSID round-trip mismatch "
                f"({got_clsid!r} != {root_clsid_str!r})"
            )
    except Exception as e:
        return None, 0, (
            "không đóng gói lại được công thức sau khi sửa\n"
            f"      chi tiết: {e}"
        )

    return raw, n_repl, None


class FixReport:
    """Kết quả 1 lượt sửa ngoặc trên 1 file .docx -- dùng chung cho cả CLI
    (main()) và các nơi gọi khác như GUI (gui.py), để không phải in ra rồi
    parse lại chuỗi text."""

    def __init__(
        self,
        total: int,
        fixed: list[tuple[str, int]],
        unchanged: list[str],
        skipped: list[tuple[str, str]],
        out_path: Path,
    ) -> None:
        """total: tổng số công thức MathType tìm thấy trong file.
        fixed: list (tên_file, số_cặp_ngoặc_đã_sửa) của từng công thức đã sửa.
        unchanged: list tên_file không có ngoặc cứng cần sửa.
        skipped: list (tên_file, lý_do_lỗi) bị bỏ qua, giữ nguyên không sửa.
        out_path: Path file .docx kết quả đã ghi ra."""
        self.total = total
        self.fixed = fixed
        self.unchanged = unchanged
        self.skipped = skipped
        self.out_path = out_path

    @property
    def total_pairs(self) -> int:
        """Tổng số cặp ngoặc đã sửa, cộng dồn từ mọi công thức trong fixed."""
        return sum(n for _, n in self.fixed)

    def summary_lines(self) -> list[str]:
        """Báo cáo dạng danh sách dòng text -- CLI in thẳng ra console, GUI
        nối lại rồi hiển thị trong ô log, không phải viết 2 lần 2 nơi."""
        lines = [
            f"Tổng số công thức MathType: {self.total}",
            f"Đã sửa: {len(self.fixed)} công thức ({self.total_pairs} cặp ngoặc)",
            f"Không có ngoặc cứng cần sửa: {len(self.unchanged)}",
            f"Bỏ qua (không tự tin đọc đúng cấu trúc, GIỮ NGUYÊN): {len(self.skipped)}",
        ]
        if self.skipped:
            lines.append("")
            lines.append("Danh sách công thức bị bỏ qua:")
            for name, err in self.skipped:
                lines.append(f"[BỎ QUA] {name}: {err}")
        lines.append("")
        lines.append(f"Đã lưu: {self.out_path}")
        return lines


def fix_mathtype_parens_in_docx(
    in_path: str | Path,
    out_path: str | Path,
    log: Callable[[str], None] | None = None,
) -> FixReport:
    """Hàm lõi: sửa các ngoặc tự co giãn phù hợp trong công thức MathType
    của file .docx tại in_path, ghi kết quả ra out_path, trả về 1 FixReport.
    Dùng thư mục tạm riêng cho mỗi lần gọi (thay vì đường dẫn cố định như
    bản CLI ban đầu) để gọi lại nhiều lần trong 1 tiến trình dài (vd từ
    GUI, người dùng bấm sửa nhiều file liên tiếp) không giẫm chân nhau.

    log(text): callback ghi 1 dòng log NGAY tại thời điểm xảy ra (giải
    nén xong, đang xử lý công thức thứ mấy/tổng số mấy, kết quả từng
    công thức, đang nén file kết quả...) -- mặc định no-op (không log
    gì) nếu không truyền, để nơi gọi cũ (không cần log) không phải sửa
    gì cả."""
    log = log or (lambda _msg: None)
    in_path = Path(in_path)
    out_path = Path(out_path)

    work = Path(tempfile.mkdtemp(prefix="fix_mathtype_"))
    try:
        log("Đang mở file .docx")
        with zipfile.ZipFile(in_path) as z:
            z.extractall(work)

        embeddings_dir = work / "word" / "embeddings"
        bin_files = sorted(embeddings_dir.glob("oleObject*.bin")) if embeddings_dir.exists() else []
        log(f"Tìm thấy {len(bin_files)} công thức MathType trong file.")

        fixed: list[tuple[str, int]] = []
        unchanged: list[str] = []
        skipped: list[tuple[str, str]] = []
        for i, bin_path in enumerate(bin_files, start=1):
            log(f"[{i}/{len(bin_files)}] Đang xử lý công thức {bin_path.name}")
            new_bytes, n_repl, error = process_one_equation(bin_path)
            if error:
                log(f"    -> [BỎ QUA] {error}")
                skipped.append((bin_path.name, error))
            elif new_bytes is None:
                log("    -> [OK] Không có ngoặc cần sửa.")
                unchanged.append(bin_path.name)
            else:
                bin_path.write_bytes(new_bytes)
                log(f"    -> [OK] Đã sửa {n_repl} cặp ngoặc.")
                fixed.append((bin_path.name, n_repl))

        log("Đang lưu file kết quả")
        if out_path.exists():
            out_path.unlink()
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in sorted(work.rglob("*")):
                if f.is_file():
                    zf.write(f, f.relative_to(work))
        log(f"Đã lưu: {out_path}")

        return FixReport(len(bin_files), fixed, unchanged, skipped, out_path)
    finally:
        shutil.rmtree(work, ignore_errors=True)
