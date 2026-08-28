# LTWordTool, AGPL-3.0 license
"""
cfb_builder.py
==============
Đóng gói 1 tập stream nhị phân (vd 'Equation Native' đã sửa cùng các
stream phụ '\\x01CompObj', '\\x01Ole', '\\x03ObjInfo') thành 1 file
CFB/OLE2 (Compound File Binary) hợp lệ, để ghi ĐÈ lên
word/embeddings/oleObjectN.bin -- đúng định dạng Word/MathType mong đợi
khi double-click mở lại công thức.

Là 1 trong 3 module con của hệ thống "sửa ngoặc MathType" -- xem
fix_mathtype_parens.py (module orchestration) để biết bối cảnh tổng
thể. Module này KHÔNG biết gì về MTEF/MathType -- chỉ nhận vào
(tên, bytes) của từng stream và ghi ra đúng layout nhị phân CFB theo
spec MS-CFB đã công bố (header, FAT, mini-FAT, directory entries...).
"""
from __future__ import annotations

import struct

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF

SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096


def _pad(b: bytes, boundary: int, fill: bytes = b'\x00') -> bytes:
    """Đệm thêm byte fill vào cuối b cho đủ bội số của boundary."""
    rem = (-len(b)) % boundary
    return b + fill * rem


def _cfb_name_key(name: str) -> tuple[int, str]:
    """Khóa so sánh tên đúng theo "Compound File Directory Sorting Order"
    của spec MS-CFB: so độ dài (số ký tự UTF-16) trước, dài bằng nhau thì so
    từng ký tự không phân biệt hoa/thường. Đây là thứ tự BẮT BUỘC giữa các
    anh/em (sibling) trong cây thư mục CFB -- xem ghi chú trong build_cfb()."""
    return (len(name), name.upper())


def build_cfb(
    streams: list[tuple[str, bytes]], root_clsid: bytes = b'\x00' * 16
) -> bytes:
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
    mini_start_sector: dict[int, int] = {}
    for i in mini_idx:
        d = datas[i]
        if len(d) == 0:
            mini_start_sector[i] = ENDOFCHAIN
            continue
        mini_start_sector[i] = len(mini_blob) // MINI_SECTOR
        mini_blob += d
        mini_blob = bytearray(_pad(bytes(mini_blob), MINI_SECTOR))
    mini_blob_bytes = bytes(mini_blob)

    regular_sectors: list[bytes] = []

    def add_regular(data_bytes: bytes) -> tuple[int, int]:
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
        add_regular(mini_blob_bytes) if mini_blob_bytes else (ENDOFCHAIN, 0)
    )

    big_start_sector: dict[int, int] = {}
    for i in big_idx:
        s, _n = add_regular(datas[i])
        big_start_sector[i] = s

    minifat_vals: list[int] = []
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
    entries_bytes: list[bytes] = []

    def pack_entry(
        name: str, etype: int, color: int, left: int, right: int, child: int,
        sector_start: int, size: int, clsid: bytes = b'\x00' * 16,
    ) -> bytes:
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
        ministream_start, len(mini_blob_bytes), clsid=root_clsid,
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

    def chain(start_sector: int, nsec: int) -> None:
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
    if fat_nsec > 109:
        raise ValueError(
            f"cần {fat_nsec} sector FAT, vượt quá 109 mục DIFAT mà header CFB "
            f"hỗ trợ trực tiếp (chưa cài DIFAT sector mở rộng) -- dữ liệu quá "
            f"lớn để đóng gói bằng cách hiện tại"
        )
    difat = [FREESECT] * 109
    for k in range(fat_nsec):
        difat[k] = fat_sector_start + k
    header += b''.join(struct.pack('<L', v) for v in difat)
    assert len(header) == 512, len(header)

    body = b''.join(regular_sectors)
    return header + body
