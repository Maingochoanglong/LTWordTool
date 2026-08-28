# LTWordTool, AGPL-3.0 license
"""
mtef_transform.py
==================
Duyệt cây Rec (do mtef_parser.py đọc ra) để tìm và chuyển các fence
(ngoặc, dấu ngoặc vuông...) TỰ CO GIÃN bao quanh văn bản/số/phép tính
ĐƠN GIẢN thành ngoặc CỨNG -- các template có cấu trúc toán học thật sự
(phân số, căn, luỹ thừa lồng bên trong ngoặc...) được giữ nguyên, không
đụng tới.

Là 1 trong 3 module con của hệ thống "sửa ngoặc MathType" -- xem
fix_mathtype_parens.py (module orchestration) để biết bối cảnh tổng thể
và lý do an toàn (chỉ sửa khi tự tin đọc đúng 100% cấu trúc, còn lại giữ
nguyên). Chỉ phụ thuộc mtef_parser.py (đọc nhị phân) -- KHÔNG biết gì về
OLE/CFB hay .docx.

Ý TƯỞNG THUẬT TOÁN: mỗi lần gọi transform_mtef() chỉ tìm và sửa ĐÚNG 1
cặp ngoặc phù hợp đầu tiên (_find_first_simple_expanding_fence()), sau
đó PARSE LẠI TỪ ĐẦU dữ liệu đã sửa (để không phải tự dịch offset của mọi
Rec còn lại bằng tay), lặp lại tới khi không còn cặp nào phù hợp -- nhờ
vậy 1 ngoặc lồng bên trong 1 ngoặc khác (vd '[x-2(2a+3)]') được sửa từ
TRONG ra NGOÀI, đúng 1 bước 1 lần, an toàn hơn cố gắng sửa nhiều chỗ
cùng lúc trên cùng 1 mảng byte.
"""
from __future__ import annotations

import struct

from mtef_parser import MTEFError, Rec, parse_mtef

LITERAL_FENCE_TYPEFACE = 2
EXPANDING_FENCE_TYPEFACE = 22
FENCE_TEMPLATE_SELECTORS = {1, 3, 9}
SIMPLE_TEXT_SYMBOLS = frozenset(" +-*/=.,:;_−")


def _child_list(rec: Rec, name: str) -> list[Rec] | None:
    """Danh sách Rec con của rec ứng với tên nhóm name (vd "content",
    "subobjects", "cellN") -- xem Parser.parse_record() trong
    mtef_parser.py để biết mỗi rtype gắn tên nhóm nào vào rec.children.
    None nếu rec không có nhóm con tên đó."""
    for child_name, items in rec.children:
        if child_name == name:
            return items
    return None


def _is_simple_text_char(char: str) -> bool:
    """True nếu char là chữ/số, khoảng trắng, hoặc 1 trong các ký hiệu
    phép tính đơn giản cho phép trong SIMPLE_TEXT_SYMBOLS -- dùng để
    quyết định nội dung bên trong 1 fence có được coi là "văn bản/số đơn
    giản" (an toàn để đổi ngoặc) hay là biểu thức toán học phức tạp hơn
    (giữ nguyên, không đụng tới)."""
    return char.isalnum() or char.isspace() or char in SIMPLE_TEXT_SYMBOLS


def _is_literal_fence_char(item: Rec) -> bool:
    """True nếu item là 1 CHAR record đã LÀ ngoặc cứng (typeface ==
    LITERAL_FENCE_TYPEFACE, glyph trong '()[]') -- tức 1 cặp ngoặc lồng
    bên trong đã được transform_mtef() sửa ở vòng lặp TRƯỚC đó (xử lý
    trong cùng trước, xem docstring transform_mtef()). Nhận diện theo
    typeface (nguồn gốc thật của ký tự), KHÔNG theo bản thân ký tự '('/
    ')'/'['/']' -- để không nhận nhầm ngoặc tự co giãn (typeface
    EXPANDING_FENCE_TYPEFACE) hay ngoặc gõ tay bình thường của người
    dùng (typeface Text khác) là "đã chuyển"."""
    return (
        item.rtype == 2
        and item.typeface == LITERAL_FENCE_TYPEFACE
        and item.glyph in "()[]"
    )


def _simple_line_content(line: Rec, source: bytes) -> bytes | None:
    """Nội dung 1 LINE (rtype 1) có được coi là 'văn bản/số đơn giản' hay
    không -- CHO PHÉP bên trong đã chứa 1 hay nhiều cặp ngoặc cứng do
    chính transform_mtef() vừa chuyển ở (các) vòng lặp trước (vd
    '(2a+3)' đã hoá cứng nằm trong '[x-2(2a+3)]'): những CHAR đó được
    nhận diện qua _is_literal_fence_char() và bỏ qua kiểm tra
    SIMPLE_TEXT_SYMBOLS -- nhờ vậy ngoặc BAO NGOÀI 1 ngoặc lồng đã sửa
    vẫn được sửa tiếp ở vòng lặp kế tiếp, thay vì bị coi là 'phức tạp'
    chỉ vì có ký tự '('/')' xuất hiện trong nội dung."""
    items = _child_list(line, "content")
    if not items:
        return None
    content = [item for item in items if item.rtype != 0]
    if not content:
        return None

    text: list[tuple[str, bool]] = []
    for item in content:
        if item.rtype == 2:
            if item.glyph is None or item.children:
                return None
            text.append((item.glyph, _is_literal_fence_char(item)))
        elif item.rtype not in range(8, 20):
            return None

    value = "".join(char for char, _is_fence in text)
    if not value or not any(char.isalnum() for char in value):
        return None
    if not all(is_fence or _is_simple_text_char(char) for char, is_fence in text):
        return None
    return source[content[0].start:items[-1].start]


def _literal_fence(glyph: str) -> bytes:
    """Bytes MTEF của 1 bản ghi CHAR ngoặc CỨNG mang glyph (vd '(', ')',
    '[', ']') -- dùng để thay thế 2 đầu 1 fence tự co giãn khi
    transform_mtef() xác nhận nội dung bên trong đủ đơn giản."""
    return bytes([2, 0, LITERAL_FENCE_TYPEFACE + 128]) + struct.pack("<H", ord(glyph))


def _find_first_simple_expanding_fence(
    top_records: list[Rec], source: bytes
) -> tuple[Rec, str, str, bytes] | None:
    """Tìm TMPL (rtype 3) ĐẦU TIÊN, duyệt theo thứ tự xuất hiện (kể cả
    lồng bên trong record khác, xem vòng lặp mở rộng records bên dưới),
    thoả TẤT CẢ điều kiện:
      - selector thuộc FENCE_TEMPLATE_SELECTORS (kiểu template "ngoặc bao").
      - subobject đầu tiên là 1 LINE (rtype 1) chứa nội dung ở giữa.
      - đúng 2 CHAR mang typeface EXPANDING_FENCE_TYPEFACE (ngoặc tự co
        giãn) làm 2 "cánh" bao quanh, đúng cặp mở/đóng hợp lệ.
      - nội dung LINE ở giữa được _simple_line_content() công nhận là
        "văn bản/số đơn giản".
    Trả về (bản ghi TMPL, glyph ngoặc trái, glyph ngoặc phải, bytes nội
    dung ở giữa) của lần khớp ĐẦU TIÊN, hoặc None nếu không còn fence
    nào phù hợp."""
    records = list(top_records)
    for rec in records:
        for _name, children in rec.children:
            records.extend(children)

        if rec.rtype != 3 or rec.selector not in FENCE_TEMPLATE_SELECTORS:
            continue
        subobjects = _child_list(rec, "subobjects")
        if not subobjects or subobjects[0].rtype != 1:
            continue
        fences = [
            item for item in subobjects[1:]
            if item.rtype == 2
            and item.typeface == EXPANDING_FENCE_TYPEFACE
            and item.glyph in "()[]"
        ]
        if len(fences) != 2:
            continue
        left, right = fences
        if left.glyph not in "([" or right.glyph not in ")]":
            continue
        inner = _simple_line_content(subobjects[0], source)
        if inner is not None:
            return rec, left.glyph, right.glyph, inner
    return None


def transform_mtef(mtef_bytes: bytes, max_iterations: int = 200) -> tuple[bytes, int]:
    """Chuyển các fence tự co giãn bao quanh văn bản/số đơn giản về ngoặc
    cứng. Các template có cấu trúc toán học được giữ nguyên."""
    current = bytearray(mtef_bytes)
    n_replacements = 0

    for _ in range(max_iterations):
        hdr, top, endpos, parser = parse_mtef(bytes(current))
        if endpos != len(current):
            raise MTEFError(
                f"đọc xong nhưng còn dư {len(current) - endpos} byte chưa xử lý, "
                f"cấu trúc công thức không khớp như mong đợi"
            )

        fence = _find_first_simple_expanding_fence(top, bytes(current))
        if fence is None:
            break

        template, left, right, inner = fence
        replacement = _literal_fence(left) + inner + _literal_fence(right)
        current[template.start:template.end] = replacement
        n_replacements += 1
    else:
        raise RuntimeError("too many paren-rewrite iterations; aborting for safety")

    return bytes(current), n_replacements


def transform_equation_native(full_stream: bytes) -> tuple[bytes, int]:
    """full_stream = raw stream 'Equation Native' (header nhỏ + MTEF).
    Trả về (new_full_stream, n_replacements). Ném lỗi nếu parse thất bại --
    nơi gọi cần bắt lỗi và bỏ qua công thức đó, giữ nguyên không sửa.
    Offset 8:10 của header nhỏ này là độ dài payload MTEF (đã đối chiếu thực
    nghiệm khớp đúng len(mtef) trên nhiều mẫu thật) -- phải cập nhật lại nếu
    có sửa, nếu không stream sẽ không nhất quán nội bộ."""
    hdr_len = full_stream[0] | (full_stream[1] << 8)
    header = bytearray(full_stream[:hdr_len])
    mtef = full_stream[hdr_len:]

    new_mtef, n_replacements = transform_mtef(mtef)

    if n_replacements:
        if len(new_mtef) > 0xFFFF:
            raise MTEFError(
                f"công thức sau khi sửa dài {len(new_mtef)} byte, vượt quá giới hạn "
                f"lưu độ dài của định dạng (65535 byte) -- không thể ghi lại an toàn"
            )
        struct.pack_into('<H', header, 8, len(new_mtef) & 0xFFFF)

    return bytes(header) + new_mtef, n_replacements
