#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
replace_docx.py
================
Thay thế nội dung (kèm định dạng) trong file Word thật (.docx), theo 1
hay NHIỀU cặp (tìm, thay), chạy TUẦN TỰ trên cùng 1 file gốc:

    file gốc (source)  = file cần sửa
    file tìm            = đoạn nội dung cần TÌM trong file gốc
    file thay thế       = đoạn nội dung dùng để THAY vào chỗ đó
    file kết quả (out)  = file ghi ra sau khi thay xong TOÀN BỘ các cặp

Module thuần (không có CLI/main() riêng) -- chỉ để gui.py (qua
pipeline.py) import và gọi hàm replace_docx() bên dưới. Xem docstring
của hàm đó để biết chi tiết nguyên tắc xử lý.

Yêu cầu: pip install lxml
"""

import copy
import os
import zipfile

from lxml import etree

W_URI = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DOCUMENT_PART = 'word/document.xml'  # đường dẫn cố định của phần nội dung chính trong mọi file .docx


def w(tag):
    """Tên thẻ kèm namespace Word. vd: w('p') -> '{...}p'."""
    return f'{{{W_URI}}}{tag}'


def check_is_real_docx(path):
    """Bắt lỗi sớm, rõ ràng nếu file không phải Word .docx thật (zip OOXML)."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Không tìm thấy file: {path}")
    if not zipfile.is_zipfile(path):
        raise ValueError(
            f"'{path}' không phải file Word .docx thật (không mở được như file zip). "
            f"Có thể đây là file text thuần được đặt tên đuôi .docx, hoặc file bị hỏng."
        )


def read_document_xml_root(docx_path):
    """Đọc thẳng word/document.xml từ trong file .docx bằng zipfile (không giải nén ra đĩa)."""
    with zipfile.ZipFile(docx_path) as z:
        return etree.fromstring(z.read(DOCUMENT_PART))


def paragraph_text(p):
    """Nối toàn bộ chữ (mọi thẻ w:t) trong 1 đoạn văn, bất kể Word có tách
    thành bao nhiêu run nhỏ bên trong."""
    return ''.join(t.text or '' for t in p.findall('.//' + w('t')))


def paragraph_is_boilerplate_empty(p):
    """Có phải đoạn trắng cuối file - dấu kết thúc mặc định của Word - hay
    không: không có chữ, và không chứa ảnh/hình vẽ/object nào."""
    if paragraph_text(p).strip():
        return False
    return not any(p.find('.//' + w(tag)) is not None for tag in ('drawing', 'pict', 'object'))


def get_content_paragraphs(root):
    """Danh sách đoạn văn 'nội dung thật' trực tiếp trong <w:body>, bỏ đoạn
    trắng cuối cùng nếu có (không tính là nội dung)."""
    paras = [c for c in root.find(w('body')) if c.tag == w('p')]
    if len(paras) > 1 and paragraph_is_boilerplate_empty(paras[-1]):
        paras = paras[:-1]
    return paras


def find_match_positions(source_paras, pattern_texts):
    """Mọi vị trí bắt đầu (không chồng lấn) trong source_paras mà một dãy
    đoạn văn liên tiếp có nội dung chữ khớp CHÍNH XÁC với pattern_texts."""
    n = len(pattern_texts)
    positions = []
    i = 0
    while n and i <= len(source_paras) - n:
        if [paragraph_text(source_paras[i + k]) for k in range(n)] == pattern_texts:
            positions.append(i)
            i += n  # nhảy qua đoạn vừa khớp để không chồng lấn
        else:
            i += 1
    return positions


def _apply_one_pair(source_root, find_path, replacement_path):
    """Lõi xử lý 1 cặp: tìm nội dung find_path trong source_root (lxml
    Element <w:document> đã đọc sẵn trong bộ nhớ), thay bằng nội dung
    replacement_path, SỬA TRỰC TIẾP lên source_root. Không đọc/ghi file
    zip nào ở đây -- tách riêng để replace_docx() có thể gọi lại hàm này
    liên tiếp cho nhiều cặp trên CÙNG 1 source_root đang có trong bộ nhớ
    (cặp sau thấy được kết quả cặp trước đã sửa), mà không phải ghi ra
    rồi đọc lại file tạm giữa mỗi cặp.

    Trả về số chỗ đã thay của riêng cặp này."""
    for p in (find_path, replacement_path):
        check_is_real_docx(p)

    find_paras = get_content_paragraphs(read_document_xml_root(find_path))
    replacement_paras = get_content_paragraphs(read_document_xml_root(replacement_path))
    if not find_paras:
        raise ValueError("File tìm không có nội dung nào để tìm.")
    pattern_texts = [paragraph_text(p) for p in find_paras]

    source_paras = [c for c in source_root.find(w('body')) if c.tag == w('p')]

    positions = find_match_positions(source_paras, pattern_texts)
    n = len(pattern_texts)
    for pos in positions:
        anchor = source_paras[pos]  # đoạn đầu tiên của dãy khớp - dùng làm mốc chèn
        for rp in replacement_paras:
            anchor.addprevious(copy.deepcopy(rp))
        for k in range(n):
            source_paras[pos + k].getparent().remove(source_paras[pos + k])

    return len(positions)


def _write_docx_with_new_document_xml(source_path, source_root, out_path):
    """Ghi ra out_path: copy nguyên file .docx tại source_path (zip), chỉ
    thay đúng phần word/document.xml bằng nội dung hiện tại của
    source_root (mọi phần khác -- ảnh, style, theme, header/footer...
    -- copy y nguyên byte-for-byte)."""
    new_document_xml = etree.tostring(source_root, xml_declaration=True, encoding='UTF-8')
    with zipfile.ZipFile(source_path) as zin, \
         zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_document_xml if item.filename == DOCUMENT_PART else zin.read(item.filename)
            zout.writestr(item, data)


def replace_docx(source_path, pairs, out_path):
    """Thay nội dung theo TỪNG CẶP (find_path, replacement_path) trong
    pairs, chạy TUẦN TỰ trên source_path (cặp sau thấy được kết quả cặp
    trước đã sửa), xuất kết quả CUỐI CÙNG ra out_path. Trả về list số chỗ
    đã thay, đúng theo thứ tự pairs.

    Danh sách chỉ có 1 cặp vẫn gọi hàm này (pairs=[(find, thay)]) -- không
    có hàm riêng cho trường hợp 1 cặp, để tránh 2 hàm gần giống nhau cho
    cùng 1 việc.

    Nguyên tắc mỗi cặp:
      1. Lấy đoạn văn "nội dung thật" của file tìm và file thay thế (bỏ
         đoạn trắng cuối file).
      2. Tìm trong file gốc (đã qua các cặp trước đó, nếu có) mọi dãy
         đoạn văn liên tiếp khớp CHÍNH XÁC văn bản của file tìm.
      3. Thay bằng đoạn của file thay thế, GIỮ NGUYÊN định dạng gốc của
         file đó (font, cỡ chữ, màu, đậm/nghiêng/gạch chân, căn lề, thụt
         lề, style đoạn văn...) - không lấy định dạng tại chỗ bị thay.
      4. Chỉ word/document.xml thay đổi; mọi phần khác của file .docx
         (ảnh, style, theme, header/footer...) được copy y nguyên
         byte-for-byte từ file gốc.

    Nếu 1 cặp lỗi giữa chừng (vd file không đọc được, không có nội dung để
    tìm...): DỪNG NGAY, không chạy cặp sau, KHÔNG ghi ra out_path -- file
    kết quả chỉ được ghi khi TOÀN BỘ danh sách chạy xong không lỗi, tránh
    để lại file dở dang. Lỗi ném lại kèm số thứ tự cặp và tên 2 file, để
    biết chính xác cặp nào cần sửa.

    PHẠM VI (cố ý như vậy, không phải thiếu sót):
      - CHỈ xử lý văn bản trong THÂN tài liệu (w:body). Header, footer,
        viền trang, đánh số trang nằm ở phần khác của file .docx (header/
        footer là XML riêng; viền trang nằm trong w:sectPr) -- không bao
        giờ bị đụng tới, dù nội dung cần tìm khớp trúng đoạn nào cạnh đó.
      - KHÔNG hỗ trợ ảnh/OLE object (kể cả công thức MathType) nhúng BÊN
        TRONG đoạn văn của file thay thế: file media/quan hệ liên quan
        không được copy theo (chỉ document.xml đổi), tham chiếu sẽ vỡ,
        nội dung nhúng biến mất im lặng trong kết quả. File thay thế chỉ
        nên là văn bản có định dạng thuần.
      - Style đoạn văn (w:pStyle) được giữ dưới dạng THAM CHIẾU tên/ID,
        chỉ hiển thị đúng nếu file gốc có định nghĩa style cùng tên trong
        styles.xml của nó.
    """
    if not pairs:
        raise ValueError("Danh sách cặp tìm/thay đang trống.")

    check_is_real_docx(source_path)
    source_root = read_document_xml_root(source_path)

    counts = []
    for i, (find_path, replacement_path) in enumerate(pairs, start=1):
        try:
            counts.append(_apply_one_pair(source_root, find_path, replacement_path))
        except Exception as e:
            find_name = os.path.basename(find_path)
            replacement_name = os.path.basename(replacement_path)
            raise ValueError(
                f"Lỗi ở cặp thứ {i}/{len(pairs)} "
                f"(tìm: {find_name}, thay: {replacement_name}): {e}"
            ) from e

    _write_docx_with_new_document_xml(source_path, source_root, out_path)
    return counts
