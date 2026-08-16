

"""
Chuyển ngoặc tự co giãn trong công thức MathType thành ngoặc cứng khi phần
nội dung bên trong chỉ là chữ, số hoặc phép tính đơn giản.

Module thuần (không có CLI/main() riêng) -- chỉ để gui.py import và gọi
hàm fix_mathtype_parens_in_docx(in_path, out_path) bên dưới, trả về 1
FixReport (tổng số công thức, danh sách đã sửa/giữ nguyên/bỏ qua).

Yêu cầu: pip install olefile

Cách hoạt động: .docx là 1 file zip; mỗi công thức MathType là 1 "OLE
object" nhị phân trong word/embeddings/oleObjectN.bin, đọc/ghi bằng
olefile (thư viện Python thuần, không cần cài đặt hệ thống). Bên trong
là cấu trúc nhị phân MTEF (định dạng riêng của MathType, xem spec chính
thức tại docs.wiris.com/.../mathtype-mtef-v5-mathtype-40-and-later) --
không có thư viện Python/pip nào đọc VÀ ghi lại được định dạng này (đã
tìm), nên phần parser (lớp Parser) và phần ghi container OLE (build_cfb)
trong file này là tự viết. Script chỉ tháo các template ngoặc tự co giãn
an toàn thành ký tự ngoặc cứng, rồi ghi lại toàn bộ OLE container.

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
MathType 3.5 (cũ hơn 4.0) dùng định dạng khác hẳn (MTEF v3/v4) -- script
tự nhận diện qua byte phiên bản ghi sẵn trong header của từng công thức
và BỎ QUA an toàn (liệt kê rõ lý do trong báo cáo), không cố đọc sai.
"""
import shutil
import struct
import tempfile
import uuid
import zipfile
import io
from pathlib import Path

import olefile

class Rec:
    """1 bản ghi (record) MTEF, kèm khoảng byte [start, end) mà nó chiếm
    (end không tính children lồng bên trong), để có thể chèn/thay đúng
    đoạn byte của 1 cặp ngoặc mà không đụng phần còn lại."""
    __slots__ = ("rtype","start","end","opts","selector","variation","glyph","mtcode",
                 "typeface","children")
    def __init__(self, rtype, start):
        """Khởi tạo bản ghi loại rtype, bắt đầu tại byte start."""
        self.rtype = rtype
        self.start = start
        self.end = None
        self.opts = None
        self.selector = None
        self.variation = None
        self.glyph = None
        self.mtcode = None
        self.typeface = None
        self.children = []

    def __repr__(self):
        """Chuỗi mô tả ngắn để debug."""
        return f"<{self.rtype} {self.start}:{self.end} glyph={self.glyph!r}>"

class MTEFError(Exception):
    """Lỗi phát sinh khi parse cấu trúc MTEF thất bại."""

class Parser:
    """Đọc tuần tự các trường nhị phân trong 1 chuỗi MTEF, trả về vị trí kế tiếp sau mỗi lần đọc."""

    def __init__(self, data: bytes):
        """Nhận vào toàn bộ dữ liệu MTEF cần đọc."""
        self.d = data
        self.n = len(data)

    def u8(self, p):
        """Đọc 1 byte không dấu tại vị trí p."""
        if p >= self.n:
            raise MTEFError(f"EOF reading byte at {p}")
        return self.d[p]

    def unsigned(self, p):
        """Đọc số nguyên không dấu độ dài thay đổi (1 hoặc 3 byte) tại p."""
        b = self.u8(p)
        if b < 255:
            return b, p + 1
        lo, hi = self.u8(p+1), self.u8(p+2)
        return lo | (hi << 8), p + 3

    def signed(self, p):
        """Đọc số nguyên có dấu độ dài thay đổi (1 hoặc 3 byte) tại p."""
        b = self.u8(p)
        if b != 255:
            return b - 128, p + 1
        lo, hi = self.u8(p+1), self.u8(p+2)
        v = lo | (hi << 8)
        return v - 32768, p + 3

    def u16(self, p):
        """Đọc số nguyên 16-bit little-endian tại p."""
        return (self.u8(p) | (self.u8(p+1) << 8)), p + 2

    def cstr(self, p):
        """Đọc chuỗi kết thúc bằng byte 0 (C-string) tại p."""
        start = p
        while self.u8(p) != 0:
            p += 1
        return self.d[start:p].decode("latin-1"), p + 1

    def nudge_if(self, p, opts):
        """Bỏ qua vùng 'nudge' (offset thủ công) nếu cờ tương ứng trong opts được bật."""
        if opts & 0x08:
            b1, b2 = self.u8(p), self.u8(p+1)
            if b1 == 128 and b2 == 128:
                p += 6
            else:
                p += 2
        return p

    def dim_array(self, p):
        """Đọc mảng kích thước mã hoá dạng nibble (nửa byte) của EQN_PREFS, trả về vị trí kế tiếp."""
        count = self.u8(p); p += 1
        nibble_pos = [p, True]
        def next_nibble():
            """Lấy nibble (4 bit) kế tiếp trong dòng byte hiện tại."""
            bp, high = nibble_pos
            b = self.u8(bp)
            if high:
                nibble_pos[1] = False
                return (b >> 4) & 0xF
            else:
                nibble_pos[0] = bp + 1
                nibble_pos[1] = True
                return b & 0xF
        total_nibbles = 0
        for _ in range(count):
            next_nibble(); total_nibbles += 1
            while True:
                v = next_nibble(); total_nibbles += 1
                if v == 0xF:
                    break
        if total_nibbles % 2 == 1:
            next_nibble()
        end_p = nibble_pos[0] if nibble_pos[1] else nibble_pos[0] + 1
        return end_p

    def style_array(self, p):
        """Đọc mảng kiểu chữ (style) của EQN_PREFS, trả về vị trí kế tiếp."""
        count = self.u8(p); p += 1
        for _ in range(count):
            idx, p = self.unsigned(p)
            if idx:
                p += 1
        return p

    def parse_record(self, p):
        """Phân giải 1 bản ghi MTEF bắt đầu tại p, trả về (Rec, vị_trí_kế_tiếp).
        Các rtype chính: 0=END, 1=LINE, 2=CHAR, 3=TMPL, 4=PILE, 5=MATRIX,
        6=EMBELL, 7=RULER, 8=FONT_STYLE_DEF, 9=SIZE, 10-14=TYPESIZE, 15=COLOR,
        16=COLOR_DEF, 17=FONT_DEF, 18=EQN_PREFS, 19=ENCODING_DEF; >=100 là
        dữ liệu mở rộng tự khai báo độ dài."""
        start = p
        rtype = self.u8(p); p += 1
        rec = Rec(rtype, start)

        if rtype >= 100:
            length, p = self.unsigned(p)
            p += length
            rec.end = p
            return rec, p

        if rtype == 0:
            rec.end = p
            return rec, p

        if rtype in (1, 2, 3, 4, 5, 6):
            opts = self.u8(p); p += 1
            rec.opts = opts
            p = self.nudge_if(p, opts)

            if rtype == 1:
                if opts & 0x04:
                    p += 2
                if opts & 0x02:
                    ruler, p = self.parse_record(p)
                    rec.children.append(("ruler", [ruler]))
                if not (opts & 0x01):
                    lst, p = self.parse_list(p)
                    rec.children.append(("content", lst))
                rec.end = p
                return rec, p

            if rtype == 2:
                typeface, p = self.signed(p)
                rec.typeface = typeface
                mtcode = None
                if not (opts & 0x20):
                    mtcode, p = self.u16(p)
                if opts & 0x04:
                    p += 1
                if opts & 0x10:
                    p += 2
                rec.mtcode = mtcode
                if mtcode is not None:
                    try:
                        rec.glyph = chr(mtcode)
                    except ValueError:
                        rec.glyph = None
                if opts & 0x01:
                    lst, p = self.parse_list(p)
                    rec.children.append(("embell", lst))
                rec.end = p
                return rec, p

            if rtype == 3:
                selector = self.u8(p); p += 1
                v1 = self.u8(p); p += 1
                if v1 & 0x80:
                    v2 = self.u8(p); p += 1
                    variation = (v1 & 0x7F) | (v2 << 8)
                else:
                    variation = v1
                p += 1
                rec.selector = selector
                rec.variation = variation
                p = self.parse_template_body(rec, p)
                rec.end = p
                return rec, p

            if rtype == 4:
                p += 2
                if opts & 0x02:
                    ruler, p = self.parse_record(p)
                    rec.children.append(("ruler", [ruler]))
                lst, p = self.parse_list(p)
                rec.children.append(("lines", lst))
                rec.end = p
                return rec, p

            if rtype == 5:
                p += 3
                rows = self.u8(p); p += 1
                cols = self.u8(p); p += 1
                p += (2 * (rows + 1) + 7) // 8
                p += (2 * (cols + 1) + 7) // 8

                for i in range(cols):
                    lst, p = self.parse_list(p)
                    rec.children.append((f"cell{i}", lst))
                rec.end = p
                return rec, p

            if rtype == 6:
                p += 1
                rec.end = p
                return rec, p

        if rtype == 7:
            n, p = self.unsigned(p)
            p += n * 3
            rec.end = p
            return rec, p

        if rtype == 8:
            _, p = self.unsigned(p)
            p += 1
            rec.end = p
            return rec, p

        if rtype == 9:
            b1 = self.u8(p); p += 1
            if b1 == 101:
                p += 2
            elif b1 == 100:
                p += 1 + 2
            else:
                p += 1
            rec.end = p
            return rec, p

        if 10 <= rtype <= 14:
            rec.end = p
            return rec, p

        if rtype == 15:
            _, p = self.unsigned(p)
            rec.end = p
            return rec, p

        if rtype == 16:
            opts = self.u8(p); p += 1
            n = 4 if opts & 0x01 else 3
            p += n * 2
            if opts & 0x04:
                _, p = self.cstr(p)
            rec.end = p
            return rec, p

        if rtype == 17:
            _, p = self.unsigned(p)
            _, p = self.cstr(p)
            rec.end = p
            return rec, p

        if rtype == 18:
            p += 1
            p = self.dim_array(p)
            p = self.dim_array(p)
            p = self.style_array(p)
            rec.end = p
            return rec, p

        if rtype == 19:
            _, p = self.cstr(p)
            rec.end = p
            return rec, p

        raise MTEFError(f"Unknown record type {rtype} at byte {start}")

    def parse_template_body(self, rec, p):
        """A TMPL's subobjects (however many slots/characters its class has) are
        ALL just ONE flat object list, terminated by a single END -- exactly
        like LINE/PILE/MATRIX-cell content. Confirmed empirically: a 2-slot
        fraction is [COLOR, LINE, COLOR, LINE, END], one trailing END total,
        not one per slot."""
        items, p = self.parse_list(p)
        rec.children.append(("subobjects", items))
        return p

    def parse_list(self, p):
        """Đọc liên tiếp các bản ghi tại p cho đến khi gặp END (rtype 0)."""
        items = []
        while True:
            rec, p2 = self.parse_record(p)
            items.append(rec)
            p = p2
            if rec.rtype == 0:
                break
        return items, p

def parse_header(data):
    """Đọc phần header đầu MTEF (phiên bản, nền tảng, ứng dụng nguồn...),
    trả về (dict thông tin header, vị trí kế tiếp)."""
    p = 0
    ver = data[p]; p += 1
    plat = data[p]; p += 1
    prod = data[p]; p += 1
    prodver = data[p]; p += 1
    prodsub = data[p]; p += 1
    start = p
    while data[p] != 0:
        p += 1
    appkey = data[start:p].decode("latin-1")
    p += 1
    opts = data[p]; p += 1
    return dict(version=ver, platform=plat, product=prod, prodver=prodver,
                prodsub=prodsub, appkey=appkey, inline=bool(opts & 1)), p

def parse_mtef(mtef_bytes):
    """Top level = zero or more def/pref records, then exactly one PILE or LINE
    record (which is fully self-contained, including its own content + END).
    There is no additional outer terminator beyond that final record.
    Chỉ hỗ trợ MTEF v5 (dùng bởi MathType 4.0 trở lên -- toàn bộ dòng
    MathType 4/5/6.x/7.x/365 hiện có). MTEF v3 (Equation Editor 3.x) và v4
    (MathType 3.5 riêng) có cấu trúc khác hẳn nên bị chặn ở đây, báo lỗi rõ
    ràng thay vì đọc sai."""
    hdr, p = parse_header(mtef_bytes)
    if hdr['version'] != 5:
        raise MTEFError(
            f"MTEF version {hdr['version']} không được hỗ trợ (công cụ này chỉ đọc "
            f"MTEF v5, dùng bởi MathType 4.0 trở lên). Công thức này nhiều khả năng "
            f"đến từ Equation Editor 3.x hoặc MathType 3.5 (MTEF v3/v4)."
        )
    parser = Parser(mtef_bytes)
    top = []
    while True:
        rec, p = parser.parse_record(p)
        top.append(rec)
        if rec.rtype in (1, 4):
            if p < len(mtef_bytes):
                end_rec, p = parser.parse_record(p)
                top.append(end_rec)
            break
        if p >= len(mtef_bytes):
            break
    return hdr, top, p, parser

LITERAL_FENCE_TYPEFACE = 2
EXPANDING_FENCE_TYPEFACE = 22
FENCE_TEMPLATE_SELECTORS = {1, 3, 9}
SIMPLE_TEXT_SYMBOLS = frozenset(" +-*/=.,:;_−")


def _child_list(rec, name):
    for child_name, items in rec.children:
        if child_name == name:
            return items
    return None


def _is_simple_text_char(char):
    return char.isalnum() or char.isspace() or char in SIMPLE_TEXT_SYMBOLS


def _simple_line_content(line, source):
    items = _child_list(line, "content")
    if not items:
        return None
    content = [item for item in items if item.rtype != 0]
    if not content:
        return None

    text = []
    for item in content:
        if item.rtype == 2:
            if item.glyph is None or item.children:
                return None
            text.append(item.glyph)
        elif item.rtype not in range(8, 20):
            return None

    value = "".join(text)
    if not value or not any(char.isalnum() for char in value):
        return None
    if not all(_is_simple_text_char(char) for char in value):
        return None
    return source[content[0].start:items[-1].start]


def _literal_fence(glyph):
    return bytes([2, 0, LITERAL_FENCE_TYPEFACE + 128]) + struct.pack("<H", ord(glyph))


def _find_first_simple_expanding_fence(top_records, source):
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

def transform_mtef(mtef_bytes: bytes, max_iterations=200):
    """Chuyển các fence tự co giãn bao quanh văn bản/số đơn giản về ngoặc
    cứng. Các template có cấu trúc toán học được giữ nguyên."""
    current = bytearray(mtef_bytes)
    n_replacements = 0

    for _ in range(max_iterations):
        hdr, top, endpos, parser = parse_mtef(bytes(current))
        if endpos != len(current):
            raise MTEFError(f"leftover {len(current) - endpos} bytes after parse")

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

def transform_equation_native(full_stream: bytes):
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
        struct.pack_into('<H', header, 8, len(new_mtef) & 0xFFFF)

    return bytes(header) + new_mtef, n_replacements

FREESECT     = 0xFFFFFFFF
ENDOFCHAIN   = 0xFFFFFFFE
FATSECT      = 0xFFFFFFFD
DIFSECT      = 0xFFFFFFFC
NOSTREAM     = 0xFFFFFFFF

SECTOR       = 512
MINI_SECTOR  = 64
MINI_CUTOFF  = 4096

def _pad(b, boundary, fill=b'\x00'):
    """Đệm thêm byte fill vào cuối b cho đủ bội số của boundary."""
    rem = (-len(b)) % boundary
    return b + fill * rem

def _cfb_name_key(name):
    """Khóa so sánh tên đúng theo "Compound File Directory Sorting Order"
    của spec MS-CFB: so độ dài (số ký tự UTF-16) trước, dài bằng nhau thì so
    từng ký tự không phân biệt hoa/thường. Đây là thứ tự BẮT BUỘC giữa các
    anh/em (sibling) trong cây thư mục CFB -- xem ghi chú trong build_cfb()."""
    return (len(name), name.upper())

def build_cfb(streams, root_clsid=b'\x00' * 16):
    """streams: list [(tên, bytes), ...] theo đúng thứ tự cần có.
    root_clsid: CLSID gốc (16 byte) của storage -- BẮT BUỘC phải giữ nguyên
    giá trị thật (vd MathType là 0002CE03-0000-0000-C000-000000000046), nếu
    không Word sẽ không biết mở công thức bằng app nào nữa (mất luôn
    right-click -> Edit / double-click để sửa), dù nội dung công thức vẫn
    đọc/hiển thị bình thường -- lỗi này KHÔNG gây crash hay sai công thức
    nên rất dễ bị bỏ sót nếu chỉ test bằng cách mở/xem, phải test bằng
    double-click từng công thức mới thấy.
    Trả về bytes của 1 file CFB/OLE2 hợp lệ chứa đúng các stream đó, đặt
    phẳng làm con trực tiếp của root storage (không có sub-storage).
    Layout (header, directory entry 128 byte, FAT sentinel...) đối chiếu
    theo đúng cấu trúc MS-CFB đã công bố."""

    streams = sorted(streams, key=lambda item: _cfb_name_key(item[0]))
    names = [n for n, _ in streams]
    datas = [d for _, d in streams]

    mini_idx = [i for i, d in enumerate(datas) if len(d) < MINI_CUTOFF]
    big_idx = [i for i, d in enumerate(datas) if len(d) >= MINI_CUTOFF]

    mini_blob = bytearray()
    mini_start_sector = {}
    for i in mini_idx:
        d = datas[i]
        if len(d) == 0:
            mini_start_sector[i] = ENDOFCHAIN
            continue
        mini_start_sector[i] = len(mini_blob) // MINI_SECTOR
        mini_blob += d
        mini_blob = bytearray(_pad(bytes(mini_blob), MINI_SECTOR))
    mini_blob = bytes(mini_blob)

    regular_sectors = []

    def add_regular(data_bytes):
        """Cấp phát sector 512-byte cho data_bytes, trả về (sector bắt đầu, số sector)."""
        start = len(regular_sectors)
        if len(data_bytes) == 0:
            return ENDOFCHAIN, 0
        padded = _pad(data_bytes, SECTOR)
        n = len(padded) // SECTOR
        for i in range(n):
            regular_sectors.append(padded[i * SECTOR:(i + 1) * SECTOR])
        return start, n

    ministream_start, ministream_nsec = (
        add_regular(mini_blob) if mini_blob else (ENDOFCHAIN, 0)
    )

    big_start_sector = {}
    for i in big_idx:
        s, _n = add_regular(datas[i])
        big_start_sector[i] = s

    minifat_vals = []
    for i in mini_idx:
        d = datas[i]
        count = (len(d) + MINI_SECTOR - 1) // MINI_SECTOR
        base = len(minifat_vals)
        for k in range(count):
            minifat_vals.append(base + k + 1 if k < count - 1 else ENDOFCHAIN)
    minifat_bytes = b''.join(struct.pack('<L', v) for v in minifat_vals)
    minifat_bytes = _pad(minifat_bytes, SECTOR, fill=struct.pack('<L', FREESECT))
    minifat_start, minifat_nsec = (
        add_regular(minifat_bytes) if minifat_vals else (ENDOFCHAIN, 0)
    )

    n_streams = len(names)
    dir_count = 1 + n_streams
    entries_bytes = []

    def pack_entry(name, etype, color, left, right, child, sector_start, size, clsid=b'\x00' * 16):
        """Đóng gói 1 mục thư mục (directory entry) CFB dài 128 byte.
        clsid: chỉ Root Entry mới cần giá trị thật (định danh COM cho MathType/
        Equation Editor, để Word biết mở bằng app nào khi double-click/right-
        click -> Edit) -- các stream con vẫn để toàn số 0 như bình thường."""
        raw_name = name.encode('utf-16-le')
        name_len = len(raw_name) + 2
        raw_name = _pad(raw_name + b'\x00\x00', 64)[:64]
        return struct.pack(
            '<64sHBBLLL16sLQQLQ',
            raw_name, name_len, etype, color,
            left, right, child,
            clsid, 0, 0, 0,
            sector_start, size,
        )

    root_child = 1 if n_streams else NOSTREAM
    entries_bytes.append(pack_entry(
        "Root Entry", 5, 1, NOSTREAM, NOSTREAM, root_child,
        ministream_start, len(mini_blob), clsid=root_clsid,
    ))

    for idx, name in enumerate(names):
        eid = idx + 1
        right = eid + 1 if eid < n_streams else NOSTREAM
        i = idx
        if i in mini_idx:
            sector_start = mini_start_sector[i]
            size = len(datas[i])
        else:
            sector_start = big_start_sector[i]
            size = len(datas[i])
        entries_bytes.append(pack_entry(
            name, 2, 1, NOSTREAM, right, NOSTREAM,
            sector_start, size,
        ))

    dir_bytes = b''.join(entries_bytes)
    entries_per_sector = SECTOR // 128
    pad_entries = (-dir_count) % entries_per_sector
    empty_entry = pack_entry('', 0, 1, NOSTREAM, NOSTREAM, NOSTREAM, 0, 0)
    dir_bytes += empty_entry * pad_entries
    dir_start, dir_nsec = add_regular(dir_bytes)

    total_data_sectors = len(regular_sectors)

    fat_nsec = 1
    while True:
        total_sectors = total_data_sectors + fat_nsec
        needed = (total_sectors + 127) // 128
        if needed <= fat_nsec:
            break
        fat_nsec += 1
    fat_sector_start = total_data_sectors
    total_sectors = total_data_sectors + fat_nsec

    fat = [FREESECT] * total_sectors

    def chain(start_sector, nsec):
        """Nối nsec sector liên tiếp bắt đầu từ start_sector thành 1 chuỗi trong bảng FAT."""
        for k in range(nsec):
            sector = start_sector + k
            fat[sector] = start_sector + k + 1 if k < nsec - 1 else ENDOFCHAIN

    if ministream_nsec:
        chain(ministream_start, ministream_nsec)
    for i in big_idx:
        d = datas[i]
        nsec = (len(d) + SECTOR - 1) // SECTOR
        chain(big_start_sector[i], nsec)
    if minifat_nsec:
        chain(minifat_start, minifat_nsec)
    chain(dir_start, dir_nsec)
    for k in range(fat_nsec):
        fat[fat_sector_start + k] = FATSECT

    fat_bytes = b''.join(struct.pack('<L', v) for v in fat)
    fat_bytes = _pad(fat_bytes, SECTOR, fill=struct.pack('<L', FREESECT))
    for k in range(fat_nsec):
        regular_sectors.append(fat_bytes[k * SECTOR:(k + 1) * SECTOR])

    header = b''
    header += struct.pack('>Q', 0xD0CF11E0A1B11AE1)
    header += b'\x00' * 16
    header += struct.pack('<HHHHH', 0x003E, 0x0003, 0xFFFE, 0x0009, 0x0006)
    header += b'\x00' * 6
    header += struct.pack(
        '<LLLLLLLLL',
        0,
        fat_nsec,
        dir_start,
        0,
        MINI_CUTOFF,
        minifat_start,
        minifat_nsec,
        ENDOFCHAIN,
        0,
    )
    difat = [FREESECT] * 109
    for k in range(fat_nsec):
        difat[k] = fat_sector_start + k
    header += b''.join(struct.pack('<L', v) for v in difat)
    assert len(header) == 512, len(header)

    body = b''.join(regular_sectors)
    return header + body

def process_one_equation(bin_path: Path):
    """Returns (new_bytes_or_None, n_replacements, error_or_None)."""
    try:
        ole = olefile.OleFileIO(str(bin_path))
        other_streams = []
        for n in ['\x01CompObj', '\x01Ole', '\x03ObjInfo']:
            other_streams.append((n, ole.openstream(n).read()))
        full = ole.openstream('Equation Native').read()
        root_clsid_str = ole.root.clsid
    except Exception as e:
        return None, 0, f"could not read OLE streams: {e}"

    try:
        new_full, n_repl = transform_equation_native(full)
    except Exception as e:
        return None, 0, f"MTEF parse/transform failed: {e}"

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
            return None, 0, "internal error: CFB round-trip mismatch"
        if got_clsid != root_clsid_str:
            return None, 0, f"internal error: CLSID round-trip mismatch ({got_clsid!r} != {root_clsid_str!r})"
    except Exception as e:
        return None, 0, f"CFB rebuild failed: {e}"

    return raw, n_repl, None

class FixReport:
    """Kết quả 1 lượt sửa ngoặc trên 1 file .docx -- dùng chung cho cả CLI
    (main()) và các nơi gọi khác như GUI (gui.py), để không phải in ra rồi
    parse lại chuỗi text."""

    def __init__(self, total, fixed, unchanged, skipped, out_path):
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
    def total_pairs(self):
        """Tổng số cặp ngoặc đã sửa, cộng dồn từ mọi công thức trong fixed."""
        return sum(n for _, n in self.fixed)

    def summary_lines(self):
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
                lines.append(f"  - {name}: {err}")
        lines.append("")
        lines.append(f"Đã lưu: {self.out_path}")
        return lines

def fix_mathtype_parens_in_docx(in_path, out_path):
    """Hàm lõi: sửa các ngoặc tự co giãn phù hợp trong công thức MathType
    của file .docx tại in_path, ghi kết quả ra out_path, trả về 1 FixReport.
    Dùng thư mục tạm riêng cho mỗi lần gọi (thay vì đường dẫn cố định như
    bản CLI ban đầu) để gọi lại nhiều lần trong 1 tiến trình dài (vd từ
    GUI, người dùng bấm sửa nhiều file liên tiếp) không giẫm chân nhau."""
    in_path = Path(in_path)
    out_path = Path(out_path)

    work = Path(tempfile.mkdtemp(prefix="fix_mathtype_"))
    try:
        with zipfile.ZipFile(in_path) as z:
            z.extractall(work)

        embeddings_dir = work / "word" / "embeddings"
        bin_files = sorted(embeddings_dir.glob("oleObject*.bin")) if embeddings_dir.exists() else []

        fixed, unchanged, skipped = [], [], []
        for bin_path in bin_files:
            new_bytes, n_repl, error = process_one_equation(bin_path)
            if error:
                skipped.append((bin_path.name, error))
            elif new_bytes is None:
                unchanged.append(bin_path.name)
            else:
                bin_path.write_bytes(new_bytes)
                fixed.append((bin_path.name, n_repl))

        if out_path.exists():
            out_path.unlink()
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in sorted(work.rglob("*")):
                if f.is_file():
                    zf.write(f, f.relative_to(work))

        return FixReport(len(bin_files), fixed, unchanged, skipped, out_path)
    finally:
        shutil.rmtree(work, ignore_errors=True)
