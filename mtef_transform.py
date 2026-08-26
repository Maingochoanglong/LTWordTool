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

QUY TẮC THEO VẾ PHƯƠNG TRÌNH (dấu '='): nếu công thức có ÍT NHẤT 1 dấu
'=' (1 CHAR glyph '=', tìm Ở BẤT KỲ ĐÂU trong toàn bộ công thức -- bất kể
lồng sâu bao nhiêu, bất kể bao nhiêu dấu '=', xem _equals_spans()), công
thức được coi là gồm NHIỀU "vế" (phân đoạn giữa 2 dấu '=' liên tiếp, hoặc
từ đầu/cuối công thức tới dấu '=' gần nhất, xem _side_range()). 1 ngoặc
CHỈ được đổi cứng nếu ĐỒNG THỜI:
  (a) nội dung NGAY BÊN TRONG nó đơn giản (như trước, xem
      _simple_line_content()), VÀ
  (b) CẢ VẾ chứa nó cũng chỉ toàn "nguyên tử đơn giản" -- chữ/số/ký hiệu
      đơn giản, ngoặc cứng đã sửa từ trước, bản ghi định dạng thuần, hoặc
      1 ngoặc tự co giãn KHÁC mà chính nội dung bên trong nó CŨNG đơn
      giản (đệ quy -- nhiều ngoặc đơn giản nằm CẠNH NHAU trong cùng vế
      không chặn lẫn nhau, xem _side_is_simple()).
  Hễ vế có BẤT KỲ cấu trúc nào khác (luỹ thừa, phân số, căn, ma trận...),
  NGOẶC ĐÓ GIỮ NGUYÊN dù nội dung riêng nó có đơn giản tới đâu. Vd
  "f(x) = 2^3(x - 2)": vế "f(x)" toàn ký tự đơn giản nên ngoặc đó đổi
  cứng được; vế "2^3(x - 2)" có luỹ thừa (2^3, không phải ngoặc) nên
  ngoặc "(x - 2)" GIỮ NGUYÊN dù bản thân "x - 2" đơn giản.
  KHÔNG có dấu '=' nào trong công thức: giữ nguyên hành vi CŨ (chỉ xét
  nội dung riêng từng ngoặc, không xét gì khác xung quanh).
"""
from __future__ import annotations

import struct

from mtef_parser import MTEFError, Rec, parse_mtef

LITERAL_FENCE_TYPEFACE = 2
EXPANDING_FENCE_TYPEFACE = 22
FENCE_TEMPLATE_SELECTORS = {1, 3, 9}
SIMPLE_TEXT_SYMBOLS = frozenset(" +-*/=.,:;_−")

# rtype các bản ghi THUẦN ĐỊNH DẠNG (ruler, font/size/color def, eqn
# prefs...) -- không mang cấu trúc toán học thật, luôn coi là "đơn giản"/
# trong suốt khi xét độ phức tạp 1 vế trong _side_is_simple(). Rộng hơn
# range(8, 20) dùng trong _simple_line_content() 1 đơn vị (thêm rtype 7 =
# RULER) vì _side_is_simple() phải duyệt CẢ nhóm con "ruler" của
# LINE/PILE -- thứ mà _simple_line_content() (chỉ đọc nhóm "content")
# không bao giờ gặp phải.
_FORMAT_ONLY_RTYPES = range(7, 20)


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


def _fence_content_line(subobjects: list[Rec]) -> tuple[Rec, list[Rec]] | None:
    """Tìm LINE nội dung trong subobjects của 1 TMPL ngoặc, BỎ QUA các
    bản ghi ĐỊNH DẠNG THUẦN (rtype 8-19, vd COLOR) có thể chen NGAY
    TRƯỚC nó -- xác nhận bằng dữ liệu thật (MathType 7.4/DSMT7): 1 TMPL
    ngoặc có thể có subobjects dạng [COLOR, LINE, '(', ')', END] thay vì
    [LINE, '(', ')', END], giống hệt cách COLOR có thể chen trước LINE
    trong 1 phân số 2 slot (xem docstring parse_template_body() trong
    mtef_parser.py) -- KHÔNG chỉ riêng phân số, TMPL ngoặc cũng gặp.
    Trả về (LINE, các phần tử CÒN LẠI sau LINE đó -- thường là 2 CHAR
    "cánh" ngoặc + END), hoặc None nếu không tìm thấy LINE đúng vị trí
    mong đợi. Dùng CHUNG cho _parse_fence() và nhánh "tổ tiên" của
    _side_is_simple(), tránh viết lặp logic bỏ-qua-COLOR ở 2 nơi."""
    idx = 0
    while idx < len(subobjects) and subobjects[idx].rtype in range(8, 20):
        idx += 1
    if idx >= len(subobjects) or subobjects[idx].rtype != 1:
        return None
    return subobjects[idx], subobjects[idx + 1:]


def _parse_fence(
    rec: Rec, source: bytes
) -> tuple[str, str, bytes | None] | None:
    """Nếu rec là 1 TMPL ngoặc (selector thuộc FENCE_TEMPLATE_SELECTORS)
    ĐÚNG cấu trúc 1 LINE nội dung (có thể có vài bản ghi định dạng thuần
    chen trước, xem _fence_content_line()) + 2 CHAR ngoặc tự co giãn hợp
    lệ bao quanh: trả về (glyph ngoặc trái, glyph ngoặc phải, bytes nội
    dung ở giữa -- None nếu nội dung đó KHÔNG đơn giản, xem
    _simple_line_content()). Trả về None (không phải tuple) nếu rec
    không đúng cấu trúc 1 ngoặc hợp lệ (sai selector, thiếu subobjects,
    không đúng 2 "cánh" ngoặc mở/đóng...).

    Dùng CHUNG cho 2 việc: (1) tìm ứng viên ngoặc cần sửa trong
    _find_first_simple_expanding_fence(), (2) đánh giá 1 ngoặc SIBLING
    (chưa sửa, nằm cùng vế với ứng viên đang xét) có được coi là "đơn
    giản" hay không trong _side_is_simple() -- tách hàm ra để 2 nơi dùng
    CHUNG đúng 1 định nghĩa "thế nào là 1 ngoặc hợp lệ", tránh viết 2 lần
    2 chỗ dễ lệch nhau."""
    if rec.rtype != 3 or rec.selector not in FENCE_TEMPLATE_SELECTORS:
        return None
    subobjects = _child_list(rec, "subobjects")
    if not subobjects:
        return None
    found = _fence_content_line(subobjects)
    if found is None:
        return None
    content_line, rest = found
    fences = [
        item for item in rest
        if item.rtype == 2
        and item.typeface == EXPANDING_FENCE_TYPEFACE
        and item.glyph in "()[]"
    ]
    if len(fences) != 2:
        return None
    left, right = fences
    if left.glyph not in "([" or right.glyph not in ")]":
        return None
    return left.glyph, right.glyph, _simple_line_content(content_line, source)


def _literal_fence(glyph: str) -> bytes:
    """Bytes MTEF của 1 bản ghi CHAR ngoặc CỨNG mang glyph (vd '(', ')',
    '[', ']') -- dùng để thay thế 2 đầu 1 fence tự co giãn khi
    transform_mtef() xác nhận nội dung bên trong đủ đơn giản."""
    return bytes([2, 0, LITERAL_FENCE_TYPEFACE + 128]) + struct.pack("<H", ord(glyph))


def _iter_all_records(records: list[Rec]):
    """Sinh lần lượt MỌI Rec trong records VÀ toàn bộ record con của
    chúng, đệ quy qua MỌI nhóm con (bất kể tên nhóm, bất kể lồng sâu bao
    nhiêu) -- dùng để tìm dấu '=' ở BẤT KỲ ĐÂU trong công thức, xem
    _equals_spans()."""
    for rec in records:
        yield rec
        for _name, children in rec.children:
            yield from _iter_all_records(children)


def _equals_spans(top_records: list[Rec]) -> list[tuple[int, int]]:
    """Khoảng byte [start, end) của MỌI dấu '=' (1 CHAR glyph '=', không
    kèm embellishment) xuất hiện Ở BẤT KỲ ĐÂU trong toàn bộ công thức --
    bất kể lồng sâu bao nhiêu, bất kể bao nhiêu dấu '=' -- SẮP XẾP theo
    vị trí byte tăng dần. Dùng làm ranh giới tách "vế" phương trình, xem
    _side_range(). [] nếu công thức không có dấu '=' nào -- khi đó
    _find_first_simple_expanding_fence() giữ nguyên hành vi CŨ, không xét
    vế gì cả."""
    return sorted(
        (rec.start, rec.end) for rec in _iter_all_records(top_records)
        if rec.rtype == 2 and not rec.children and rec.glyph == '='
    )


def _side_range(
    rec_start: int,
    rec_end: int,
    equals_spans: list[tuple[int, int]],
    formula_start: int,
    formula_end: int,
) -> tuple[int, int]:
    """Khoảng byte [side_start, side_end) của "vế" chứa bản ghi có khoảng
    [rec_start, rec_end) -- phân đoạn giữa 2 dấu '=' liên tiếp GẦN NHẤT
    (hoặc đầu/cuối công thức nếu không có dấu '=' nào ở phía đó). LOẠI
    TRỪ mọi dấu '=' nằm LỒNG BÊN TRONG chính bản ghi đang xét
    (equals_spans có thể chứa dấu '=' lồng sâu trong bất kỳ ngoặc/
    template nào, kể cả chính bản ghi này -- 1 dấu '=' nằm trong nội
    dung CỦA CHÍNH bản ghi không phải là ranh giới vế của bản ghi đó)."""
    relevant = [(s, e) for s, e in equals_spans if e <= rec_start or s >= rec_end]

    side_start = formula_start
    for eq_start, eq_end in relevant:
        if eq_start < rec_start:
            side_start = eq_end
        else:
            break

    side_end = formula_end
    for eq_start, eq_end in reversed(relevant):
        if eq_start > rec_start:
            side_end = eq_start
        else:
            break

    return side_start, side_end


def _side_is_simple(
    records: list[Rec],
    side_start: int,
    side_end: int,
    source: bytes,
    skip: Rec,
) -> bool:
    """True nếu MỌI bản ghi (đệ quy qua records) có giao với khoảng byte
    [side_start, side_end) của 1 "vế" phương trình -- TRỪ skip (chính
    ngoặc ứng viên đang xét, đã tự kiểm tra nội dung riêng ở
    _parse_fence()) -- đều là "nguyên tử đơn giản":
      - CHAR chữ/số/ký hiệu đơn giản (_is_simple_text_char), hoặc CHAR
        ngoặc cứng đã sửa từ trước (_is_literal_fence_char).
      - bản ghi thuần định dạng (_FORMAT_ONLY_RTYPES: ruler, font/size/
        color def, eqn prefs...) -- không mang cấu trúc toán học.
      - LINE/PILE (rtype 1/4) -- chỉ là khung chứa trong suốt, KHÔNG tự
        quyết, tiếp tục đệ quy vào từng nhóm con của nó.
      - 1 TMPL ngoặc tự co giãn KHÁC, LÀ SIBLING thật (không chứa skip
        bên trong) mà nội dung bên trong nó CŨNG đơn giản (nhận diện qua
        _parse_fence()) -- coi cả cặp ngoặc đó là 1 khối đã đánh giá
        xong, KHÔNG đệ quy tiếp vào bên trong nó (nhờ vậy nhiều ngoặc
        đơn giản NẰM CẠNH NHAU trong cùng vế không chặn lẫn nhau).
      - 1 TMPL ngoặc tự co giãn là TỔ TIÊN của skip (BAO TRỌN candidate
        đang xét bên trong nó, vd đang xét "(x-2)" lồng trong "((x-2)^2)")
        -- KHÔNG tự chấm điểm nguyên khối (nội dung nó còn dở dang vì
        skip chưa quyết định xong), mà coi "trong suốt", đệ quy tiếp vào
        ĐÚNG phần LINE nội dung của nó.
    Bất kỳ bản ghi nào khác nằm trong khoảng này (luỹ thừa, phân số, căn,
    ma trận, embellishment thật, hoặc 1 ngoặc SIBLING mà nội dung bên
    trong nó KHÔNG đơn giản, hoặc 1 ngoặc TỔ TIÊN không phải ngoặc tự co
    giãn hợp lệ...) đều coi là phức tạp -> trả về False NGAY (chặn cả vế,
    không cần xét tiếp)."""
    for rec in records:
        if rec is skip:
            continue
        if rec.end <= side_start or rec.start >= side_end:
            continue
        if rec.rtype == 0:
            continue
        if rec.rtype == 2:
            if rec.children:
                return False
            if _is_literal_fence_char(rec):
                continue
            if rec.glyph is not None and _is_simple_text_char(rec.glyph):
                continue
            return False
        if rec.rtype in _FORMAT_ONLY_RTYPES:
            continue
        if rec.rtype in (1, 4):
            for _name, children in rec.children:
                if not _side_is_simple(children, side_start, side_end, source, skip):
                    return False
            continue
        if rec.rtype == 3 and rec.start <= skip.start and rec.end >= skip.end:
            # rec BAO TRỌN skip (là TỔ TIÊN thật của candidate đang xét,
            # KHÔNG PHẢI 1 sibling độc lập -- xem MS-CFB/MTEF: record con
            # luôn nằm HOÀN TOÀN trong khoảng byte của cha, nên khoảng
            # [rec.start, rec.end) bao trọn [skip.start, skip.end) CHỈ có
            # thể do quan hệ cha/con, không phải trùng hợp).
            #
            # KHÔNG được tự chấm điểm rec qua _parse_fence() ở đây: nội
            # dung của rec ĐANG CÒN candidate (skip) chưa quyết định xong
            # bên trong, nên _simple_line_content() sẽ LUÔN thấy rec
            # "chưa đơn giản" 1 cách giả tạo (còn 1 TMPL con chưa hoá
            # cứng) -- dù rec có thể là 1 ngoặc tự co giãn HOÀN TOÀN hợp
            # lệ, sẽ tự đủ điều kiện đổi cứng ở 1 VÒNG LẶP SAU (sau khi
            # skip đã hoá cứng). Coi rec "trong suốt": nếu nó không phải
            # ngoặc (luỹ thừa/phân số/căn...) thì BẢN THÂN loại cấu trúc
            # đó đã đủ phức tạp rồi, chặn cả vế; nếu nó LÀ 1 ngoặc, chỉ
            # đệ quy tiếp vào ĐÚNG phần LINE nội dung của nó (subobjects[0])
            # -- KHÔNG xét 2 CHAR "cánh" ngoặc của chính rec (subobjects[1:],
            # không phải "nội dung" mà chỉ là ký hiệu ngoặc của chính rec,
            # sẽ tự được thay khi rec đủ điều kiện đổi cứng, không phải
            # thứ cần chấm "đơn giản hay không" ở đây).
            if rec.selector not in FENCE_TEMPLATE_SELECTORS:
                return False
            subobjects = _child_list(rec, "subobjects")
            if not subobjects:
                return False
            found = _fence_content_line(subobjects)
            if found is None:
                return False
            content_line, _rest = found
            content = _child_list(content_line, "content") or []
            if not _side_is_simple(content, side_start, side_end, source, skip):
                return False
            continue
        parsed = _parse_fence(rec, source)
        if parsed is not None and parsed[2] is not None:
            continue
        return False
    return True


def _find_first_simple_expanding_fence(
    top_records: list[Rec], source: bytes
) -> tuple[Rec, str, str, bytes] | None:
    """Tìm TMPL (rtype 3) ĐẦU TIÊN, duyệt theo thứ tự xuất hiện (kể cả
    lồng bên trong record khác, xem vòng lặp mở rộng records bên dưới),
    thoả TẤT CẢ điều kiện:
      - là 1 ngoặc tự co giãn hợp lệ (xem _parse_fence()): selector thuộc
        FENCE_TEMPLATE_SELECTORS, đúng 1 LINE nội dung + 2 "cánh" ngoặc
        mở/đóng hợp lệ.
      - nội dung ở giữa được _simple_line_content() công nhận là "văn
        bản/số đơn giản".
      - NẾU công thức có dấu '=' (bất kỳ đâu, xem _equals_spans()): CẢ
        VẾ (xem _side_range()) chứa ngoặc này cũng phải "đơn giản" (xem
        _side_is_simple()) -- không có dấu '=' nào cả thì bỏ qua điều
        kiện này, giữ nguyên hành vi cũ (chỉ xét nội dung riêng ngoặc).
    Trả về (bản ghi TMPL, glyph ngoặc trái, glyph ngoặc phải, bytes nội
    dung ở giữa) của lần khớp ĐẦU TIÊN thoả hết các điều kiện trên, hoặc
    None nếu không còn ngoặc nào phù hợp."""
    equals_spans = _equals_spans(top_records)
    formula_end = len(source)

    records = list(top_records)
    for rec in records:
        for _name, children in rec.children:
            records.extend(children)

        parsed = _parse_fence(rec, source)
        if parsed is None:
            continue
        left, right, inner = parsed
        if inner is None:
            continue

        if equals_spans:
            side_start, side_end = _side_range(
                rec.start, rec.end, equals_spans, 0, formula_end
            )
            if not _side_is_simple(top_records, side_start, side_end, source, skip=rec):
                continue

        return rec, left, right, inner
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
        raise RuntimeError("quá nhiều lần lặp sửa ngoặc liên tiếp -- dừng lại để an toàn")

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
