#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
replace_docx.py
================
Thay thế nội dung (kèm định dạng) trong file Word thật (.docx):

    A.docx  = file gốc
    B.docx  = đoạn nội dung cần TÌM trong A (để thay thế)
    C.docx  = đoạn nội dung dùng để THAY vào chỗ đó
    D.docx  = file kết quả

Module thuần (không có CLI/main() riêng) -- chỉ để gui.py import và gọi
hàm replace_docx() bên dưới. Xem docstring của hàm đó để biết chi tiết
nguyên tắc xử lý.

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


def find_match_positions(a_paras, pattern_texts):
    """Mọi vị trí bắt đầu (không chồng lấn) trong a_paras mà một dãy đoạn
    văn liên tiếp có nội dung chữ khớp CHÍNH XÁC với pattern_texts."""
    n = len(pattern_texts)
    positions = []
    i = 0
    while n and i <= len(a_paras) - n:
        if [paragraph_text(a_paras[i + k]) for k in range(n)] == pattern_texts:
            positions.append(i)
            i += n  # nhảy qua đoạn vừa khớp để không chồng lấn
        else:
            i += 1
    return positions


def replace_docx(a_path, b_path, c_path, d_path):
    """Thay nội dung B -> C trong A, xuất kết quả ra d_path. Trả về số chỗ đã thay.

    Nguyên tắc:
      1. Lấy đoạn văn "nội dung thật" của B và C (bỏ đoạn trắng cuối file).
      2. Tìm trong A mọi dãy đoạn văn liên tiếp khớp CHÍNH XÁC văn bản của B.
      3. Thay bằng đoạn của C, GIỮ NGUYÊN định dạng gốc của C (font, cỡ chữ,
         màu, đậm/nghiêng/gạch chân, căn lề, thụt lề...) - không lấy định
         dạng của A tại chỗ đó.
      4. Chỉ word/document.xml thay đổi; mọi phần khác của file .docx (ảnh,
         style, theme, header/footer...) được copy y nguyên byte-for-byte.
    """
    for p in (a_path, b_path, c_path):
        check_is_real_docx(p)

    b_paras = get_content_paragraphs(read_document_xml_root(b_path))
    c_paras = get_content_paragraphs(read_document_xml_root(c_path))
    if not b_paras:
        raise ValueError("B.docx không có nội dung nào để tìm.")
    pattern_texts = [paragraph_text(p) for p in b_paras]

    a_root = read_document_xml_root(a_path)
    a_paras = [c for c in a_root.find(w('body')) if c.tag == w('p')]

    positions = find_match_positions(a_paras, pattern_texts)
    n = len(pattern_texts)
    for pos in positions:
        anchor = a_paras[pos]  # đoạn đầu tiên của dãy khớp - dùng làm mốc chèn
        for cp in c_paras:
            anchor.addprevious(copy.deepcopy(cp))
        for k in range(n):
            a_paras[pos + k].getparent().remove(a_paras[pos + k])

    new_document_xml = etree.tostring(a_root, xml_declaration=True, encoding='UTF-8')

    # Copy nguyên file .docx (zip), chỉ thay đúng phần word/document.xml
    with zipfile.ZipFile(a_path) as zin, \
         zipfile.ZipFile(d_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_document_xml if item.filename == DOCUMENT_PART else zin.read(item.filename)
            zout.writestr(item, data)

    return len(positions)
