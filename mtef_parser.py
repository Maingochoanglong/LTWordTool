"""
mtef_parser.py
==============
Đọc (KHÔNG sửa, KHÔNG ghi) cấu trúc nhị phân MTEF v5 -- định dạng công
thức của "MathType 4.0 trở lên" (toàn bộ dòng MathType 4.x/5.x/6.x/
7.x/365 hiện có), theo spec chính thức tại
docs.wiris.com/.../mathtype-mtef-v5-mathtype-40-and-later.

Đây là 1 trong 3 module con của hệ thống "sửa ngoặc MathType" -- xem
fix_mathtype_parens.py (module orchestration, nơi 3 module con này được
ghép lại) để có bối cảnh tổng thể, lý do phải tự viết parser (không có
thư viện Python nào đọc VÀ ghi lại được MTEF), và các lưu ý an toàn khi
sử dụng. 2 module con còn lại:
  - mtef_transform.py : duyệt cây Rec do module này trả về, tìm & chuyển
                        ngoặc tự co giãn "đơn giản" thành ngoặc cứng.
  - cfb_builder.py     : đóng gói lại thành container OLE/CFB hợp lệ.

Tương thích phiên bản: CHỈ đọc MTEF v5 (dùng bởi MathType 4.0 trở lên).
Công thức từ Equation Editor 3.x hoặc MathType 3.5 (MTEF v3/v4, cấu trúc
khác hẳn) bị chặn ở parse_mtef() -- ném MTEFError rõ ràng thay vì cố đọc
sai, xem docstring parse_mtef().
"""
from __future__ import annotations

from typing import TypedDict


class Rec:
    """1 bản ghi (record) MTEF, kèm khoảng byte [start, end) mà nó chiếm
    (end không tính children lồng bên trong), để có thể chèn/thay đúng
    đoạn byte của 1 cặp ngoặc mà không đụng phần còn lại."""

    __slots__ = (
        "rtype", "start", "end", "opts", "selector", "variation", "glyph",
        "mtcode", "typeface", "embell_type", "children",
    )

    rtype: int
    start: int
    end: int | None
    opts: int | None
    selector: int | None
    variation: int | None
    glyph: str | None
    mtcode: int | None
    typeface: int | None
    embell_type: int | None
    children: list[tuple[str, list["Rec"]]]

    def __init__(self, rtype: int, start: int) -> None:
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
        self.embell_type = None
        self.children = []

    def __repr__(self) -> str:
        """Chuỗi mô tả ngắn để debug."""
        return f"<{self.rtype} {self.start}:{self.end} glyph={self.glyph!r}>"


class MTEFError(Exception):
    """Lỗi phát sinh khi parse cấu trúc MTEF thất bại."""


class MTEFHeader(TypedDict):
    """Thông tin header đầu 1 chuỗi MTEF -- xem parse_header()."""

    version: int
    platform: int
    product: int
    prodver: int
    prodsub: int
    appkey: str
    inline: bool


class Parser:
    """Đọc tuần tự các trường nhị phân trong 1 chuỗi MTEF, trả về vị trí kế tiếp sau mỗi lần đọc."""

    def __init__(self, data: bytes) -> None:
        """Nhận vào toàn bộ dữ liệu MTEF cần đọc."""
        self.d = data
        self.n = len(data)

    def u8(self, p: int) -> int:
        """Đọc 1 byte không dấu tại vị trí p."""
        if p >= self.n:
            raise MTEFError(f"dữ liệu công thức kết thúc bất thường (thiếu byte) tại vị trí {p}")
        return self.d[p]

    def unsigned(self, p: int) -> tuple[int, int]:
        """Đọc số nguyên không dấu độ dài thay đổi (1 hoặc 3 byte) tại p."""
        b = self.u8(p)
        if b < 255:
            return b, p + 1
        lo, hi = self.u8(p + 1), self.u8(p + 2)
        return lo | (hi << 8), p + 3

    def signed(self, p: int) -> tuple[int, int]:
        """Đọc số nguyên có dấu độ dài thay đổi (1 hoặc 3 byte) tại p."""
        b = self.u8(p)
        if b != 255:
            return b - 128, p + 1
        lo, hi = self.u8(p + 1), self.u8(p + 2)
        v = lo | (hi << 8)
        return v - 32768, p + 3

    def u16(self, p: int) -> tuple[int, int]:
        """Đọc số nguyên 16-bit little-endian tại p."""
        return (self.u8(p) | (self.u8(p + 1) << 8)), p + 2

    def cstr(self, p: int) -> tuple[str, int]:
        """Đọc chuỗi kết thúc bằng byte 0 (C-string) tại p."""
        start = p
        while self.u8(p) != 0:
            p += 1
        return self.d[start:p].decode("latin-1"), p + 1

    def nudge_if(self, p: int, opts: int) -> int:
        """Bỏ qua vùng 'nudge' (offset thủ công) nếu cờ tương ứng trong opts được bật."""
        if opts & 0x08:
            b1, b2 = self.u8(p), self.u8(p + 1)
            if b1 == 128 and b2 == 128:
                p += 6
            else:
                p += 2
        return p

    def dim_array(self, p: int) -> int:
        """Đọc mảng kích thước mã hoá dạng nibble (nửa byte) của EQN_PREFS, trả về vị trí kế tiếp."""
        count = self.u8(p); p += 1
        nibble_pos = [p, True]

        def next_nibble() -> int:
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

    def style_array(self, p: int) -> int:
        """Đọc mảng kiểu chữ (style) của EQN_PREFS, trả về vị trí kế tiếp."""
        count = self.u8(p); p += 1
        for _ in range(count):
            idx, p = self.unsigned(p)
            if idx:
                p += 1
        return p

    def parse_record(self, p: int) -> tuple[Rec, int]:
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
                rec.embell_type = self.u8(p)
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

        raise MTEFError(
            f"gặp loại bản ghi không xác định (mã {rtype}) tại byte {start}, "
            f"có thể định dạng công thức không được hỗ trợ"
        )

    def parse_template_body(self, rec: Rec, p: int) -> int:
        """Các subobject của 1 TMPL (dù class đó có bao nhiêu slot/ký tự)
        đều nằm chung trong ĐÚNG 1 danh sách object phẳng duy nhất, kết
        thúc bằng 1 END duy nhất -- giống hệt nội dung LINE/PILE/1 ô
        MATRIX. Đã xác nhận bằng thực nghiệm: 1 phân số 2 slot có dạng
        [COLOR, LINE, COLOR, LINE, END], chỉ 1 END ở cuối cùng, không
        phải 1 END/slot."""
        items, p = self.parse_list(p)
        rec.children.append(("subobjects", items))
        return p

    def parse_list(self, p: int) -> tuple[list[Rec], int]:
        """Đọc liên tiếp các bản ghi tại p cho đến khi gặp END (rtype 0)."""
        items: list[Rec] = []
        while True:
            rec, p2 = self.parse_record(p)
            items.append(rec)
            p = p2
            if rec.rtype == 0:
                break
        return items, p


def parse_header(data: bytes) -> tuple[MTEFHeader, int]:
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
    header: MTEFHeader = dict(
        version=ver, platform=plat, product=prod, prodver=prodver,
        prodsub=prodsub, appkey=appkey, inline=bool(opts & 1),
    )
    return header, p


def parse_mtef(mtef_bytes: bytes) -> tuple[MTEFHeader, list[Rec], int, Parser]:
    """Cấp cao nhất = 0 hoặc nhiều bản ghi def/pref, sau đó đúng 1 bản ghi
    PILE hoặc LINE (tự nó đã đầy đủ, gồm cả nội dung + END riêng). Không
    có ký tự kết thúc nào khác ngoài chính bản ghi cuối cùng đó.
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
    top: list[Rec] = []
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
