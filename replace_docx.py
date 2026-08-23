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

NGUYÊN TẮC TÌM/THAY (1 CƠ CHẾ DUY NHẤT, không còn phân biệt "nguyên
khối" hay "chuỗi con" như bản trước): coi nội dung file tìm là 1 CHUỖI
KÝ TỰ CÓ ĐỊNH DẠNG (mỗi ký tự kèm định dạng hiển thị của run, kể cả tab;
định dạng đoạn không tham gia so khớp; nếu file tìm có TỪ 2 đoạn văn trở
lên thì ranh giới đoạn văn cũng được coi là 1 "ký tự" phải khớp), tìm chuỗi
này ở BẤT KỲ đâu trong file gốc (có thể nằm gọn trong 1 đoạn, hoặc vắt qua
nhiều đoạn văn liên tiếp),
giống hệt Tìm & Thay của Word ở chỗ "cứ đúng chuỗi là tìm ra, không cần
biết nó nằm trong 1 hay nhiều đoạn" -- khác Word ở chỗ đây so khớp CẢ
ĐỊNH DẠNG, không chỉ thuần văn bản.

BẢNG (w:tbl): nội dung trong ô bảng CŨNG được tìm/thay, kể cả vắt qua
nhiều đoạn văn nếu 1 ô có nhiều đoạn. Nhưng mỗi ô là 1 "luồng" văn bản
RIÊNG -- chỗ khớp KHÔNG BAO GIỜ vắt từ đoạn văn trước bảng sang đoạn văn
sau bảng, hay từ ô này sang ô khác (kể cả 2 ô liền kề cùng hàng), đúng
như cảm giác trực quan rằng mỗi ô là 1 "hộp" nội dung tách biệt. Bảng
lồng trong ô bảng cũng được xử lý đúng (đệ quy), xem
_iter_paragraph_groups().

Yêu cầu: pip install lxml
"""
from __future__ import annotations

import copy
import os
import zipfile
from collections.abc import Callable, Iterator

from lxml import etree

W_URI = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_FORMAT_RELEVANT_TAGS = {
    f'{{{W_URI}}}{tag}'
    for tag in (
        'b', 'bCs', 'i', 'iCs', 'u', 'strike', 'dstrike', 'color', 'sz',
        'szCs', 'rFonts', 'vertAlign', 'highlight', 'shd', 'position',
        'spacing', 'caps', 'smallCaps', 'vanish', 'webHidden', 'bdr',
        'emboss', 'imprint', 'outline', 'shadow',
    )
}
DOCUMENT_PART = 'word/document.xml'

# 1 cặp (find_path, replacement_path, backward_stop_text, forward_stop_text);
# dạng 3 phần tử (không có forward_stop_text riêng) vẫn được chấp nhận ở
# replace_docx() cho dữ liệu cũ -- xem docstring replace_docx().
FindReplacePair = tuple[str, str, str] | tuple[str, str, str, str]


def w(tag: str) -> str:
    """Tên thẻ kèm namespace Word. vd: w('p') -> '{...}p'."""
    return f'{{{W_URI}}}{tag}'


def check_is_real_docx(path: str) -> None:
    """Bắt lỗi sớm, rõ ràng nếu file không phải Word .docx thật (zip OOXML)."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Không tìm thấy file: {path}")
    if not zipfile.is_zipfile(path):
        raise ValueError(
            f"'{path}' không phải file Word .docx thật (không mở được như file zip). "
            f"Có thể đây là file text thuần được đặt tên đuôi .docx, hoặc file bị hỏng."
        )


def read_document_xml_root(docx_path: str) -> etree._Element:
    """Đọc thẳng word/document.xml từ trong file .docx bằng zipfile (không giải nén ra đĩa)."""
    with zipfile.ZipFile(docx_path) as z:
        return etree.fromstring(z.read(DOCUMENT_PART))


def paragraph_text(p: etree._Element) -> str:
    """Nối toàn bộ chữ (mọi thẻ w:t) trong 1 đoạn văn, bất kể Word có tách
    thành bao nhiêu run nhỏ bên trong. CHỈ dùng để nhận diện đoạn trắng
    boilerplate (paragraph_is_boilerplate_empty) -- việc so khớp tìm/thay
    thật sự dùng _paragraph_atoms()/_document_atoms() bên dưới (có tính
    cả định dạng lẫn tab)."""
    return ''.join(t.text or '' for t in p.findall('.//' + w('t')))


def paragraph_is_boilerplate_empty(p: etree._Element) -> bool:
    """Có phải đoạn trắng cuối file - dấu kết thúc mặc định của Word - hay
    không: không có chữ, và không chứa ảnh/hình vẽ/object nào."""
    if paragraph_text(p).strip():
        return False
    return not any(p.find('.//' + w(tag)) is not None for tag in ('drawing', 'pict', 'object'))


def get_content_paragraphs(root: etree._Element) -> list[etree._Element]:
    """Danh sách đoạn văn 'nội dung thật' trực tiếp trong <w:body>, bỏ đoạn
    trắng cuối cùng nếu có (không tính là nội dung)."""
    paras = [c for c in root.find(w('body')) if c.tag == w('p')]
    if len(paras) > 1 and paragraph_is_boilerplate_empty(paras[-1]):
        paras = paras[:-1]
    return paras


class _Atom:
    """1 đơn vị trong chuỗi so khớp: 1 ký tự chữ (kind='t'), 1 tab
    (kind='tab'), hoặc 1 ranh giới đoạn văn (kind='pbreak', giữa 2 đoạn
    liên tiếp -- coi như 1 "ký tự" đặc biệt phải khớp CHÍNH XÁC như Enter
    thật). para = đoạn văn <w:p> "gắn với" atom này (để biết atom thuộc
    đoạn nào khi cần tách/ghép); run/el/offset chỉ có ý nghĩa với
    kind='t'/'tab' (None với 'pbreak').

    Dùng __slots__ vì đây là đối tượng tạo ra RẤT NHIỀU lần (1 atom / 1
    ký tự trong toàn bộ tài liệu) -- tránh overhead __dict__ mặc định của
    Python trên số lượng lớn instance ngắn hạn như vậy."""

    __slots__ = ("ch", "fmt", "run", "el", "offset", "kind", "para")

    ch: str
    fmt: str | None
    run: etree._Element | None
    el: etree._Element | None
    offset: int
    kind: str
    para: etree._Element

    def __init__(
        self,
        ch: str,
        fmt: str | None,
        run: etree._Element | None,
        el: etree._Element | None,
        offset: int,
        kind: str,
        para: etree._Element,
    ) -> None:
        self.ch = ch
        self.fmt = fmt
        self.run = run
        self.el = el
        self.offset = offset
        self.kind = kind
        self.para = para


def _canonical_xml(el: etree._Element) -> str:
    """Chuỗi hoá 1 phần tử XML ỔN ĐỊNH bất kể thứ tự thuộc tính/thẻ con --
    để 2 <w:rPr> cùng 1 TẬP thuộc tính định dạng nhưng khai báo khác thứ tự
    vẫn được coi là CÙNG định dạng, trong khi bất kỳ khác biệt THẬT SỰ nào
    vẫn làm 2 rPr bị coi là khác nhau."""
    def norm(e: etree._Element) -> tuple:
        attrs = tuple(sorted(e.attrib.items()))
        children = tuple(sorted(norm(c) for c in e))
        return (e.tag, attrs, children)
    return repr(norm(el))


def _run_format_key(run: etree._Element) -> str | None:
    """Khoá so sánh định dạng liên quan của 1 <w:r>.

    Chỉ các tag trong _FORMAT_RELEVANT_TAGS được dùng để so khớp; các tag
    khác (vd lang hoặc định dạng RTL) vẫn được giữ nguyên khi sao chép run.
    """
    rpr = run.find(w('rPr'))
    if rpr is None:
        return None

    relevant_rpr = copy.deepcopy(rpr)
    for child in relevant_rpr:
        if child.tag not in _FORMAT_RELEVANT_TAGS:
            relevant_rpr.remove(child)
    return _canonical_xml(relevant_rpr) if len(relevant_rpr) else None


def _paragraph_atoms(p: etree._Element) -> list[_Atom]:
    """Chuỗi phẳng các _Atom trong đoạn văn p, ĐÚNG THEO THỨ TỰ xuất hiện.
    CHỈ đọc <w:t> (chữ) và <w:tab/> (tab) bên trong mỗi run (p.iter(w('r')),
    kể cả run trong w:hyperlink) -- phần tử nội dung KHÁC (vd <w:br/>)
    hiện KHÔNG được tính vào chuỗi ký tự (xem giới hạn trong docstring
    replace_docx())."""
    atoms: list[_Atom] = []
    for r in p.iter(w('r')):

        fmt = _run_format_key(r)
        for child in r:
            if child.tag == w('t'):
                text = child.text or ''
                for i, ch in enumerate(text):
                    atoms.append(_Atom(ch, fmt, r, child, i, 't', p))
            elif child.tag == w('tab'):
                atoms.append(_Atom('\t', fmt, r, child, 0, 'tab', p))
    return atoms


def _document_atoms(paragraphs: list[etree._Element]) -> list[_Atom]:
    """Nối atoms của NHIỀU đoạn văn LIÊN TIẾP thành 1 chuỗi phẳng DUY
    NHẤT, chèn 1 _Atom kind='pbreak' giữa 2 đoạn kế nhau (đại diện ranh
    giới đoạn văn, như 1 lần Enter) -- nhờ vậy so khớp CHUỖI CON có thể
    vắt qua nhiều đoạn văn bằng ĐÚNG cơ chế so khớp atom hiện có (coi ranh
    giới đoạn văn cũng là 1 "ký tự" phải khớp CHÍNH XÁC). ch='\\n' cho
    pbreak -- an toàn, không trùng ký tự thật nào (Word không cho phép ký
    tự xuống dòng trần nằm trong <w:t>, luôn dùng <w:br/> riêng, hiện
    không được đọc vào atoms)."""
    atoms: list[_Atom] = []
    for i, p in enumerate(paragraphs):
        if i > 0:

            atoms.append(_Atom('\n', None, None, None, 0, 'pbreak', p))
        atoms.extend(_paragraph_atoms(p))
    return atoms


def _find_first_atom_match(
    atoms: list[_Atom], pattern_keys: list[tuple[str, str | None]]
) -> tuple[int, int] | None:
    """Vị trí (start, end) của chỗ khớp ĐẦU TIÊN mà pattern_keys (list
    (ch, fmt)) khớp CHÍNH XÁC, liên tục, trong atoms. None nếu không có."""
    n = len(pattern_keys)
    limit = len(atoms) - n
    for i in range(limit + 1):
        if all((atoms[i + k].ch, atoms[i + k].fmt) == pattern_keys[k] for k in range(n)):
            return i, i + n
    return None


def _build_run_from_pieces(
    source_run: etree._Element | None, pieces: list[list[str]]
) -> etree._Element:
    """Dựng 1 <w:r> MỚI mang định dạng (rPr) của source_run (deep-copy),
    chứa các mảnh pieces (['t', text] hoặc ['tab']) theo đúng thứ tự.
    xml:space="preserve" luôn đặt trên <w:t> mới -- an toàn cho mảnh có
    khoảng trắng ở đầu/cuối."""
    r = etree.Element(w('r'))
    if source_run is not None:
        rpr = source_run.find(w('rPr'))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
    for piece in pieces:
        if piece[0] == 't':
            t = etree.SubElement(r, w('t'))
            t.text = piece[1]
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        else:
            etree.SubElement(r, w('tab'))
    return r


def _atoms_to_run_pieces(atom_list: list[_Atom]) -> list[etree._Element]:
    """Chuyển 1 danh sách _Atom liên tục thành list <w:r> MỚI -- mỗi lần
    run gốc đổi lại tách thành 1 <w:r> riêng (deep-copy đúng rPr của run
    gốc đó). Viết tổng quát (không giả định chỉ 1 run gốc) dù trong thực
    tế atom_list truyền vào đây luôn chỉ thuộc ĐÚNG 1 run biên (xem
    boundary_start_run/boundary_end_run trong _splice_match)."""
    runs: list[etree._Element] = []
    cur_run: etree._Element | None = None
    cur_pieces: list[list[str]] = []
    for atom in atom_list:
        if atom.run is not cur_run:
            if cur_pieces:
                runs.append(_build_run_from_pieces(cur_run, cur_pieces))
            cur_run, cur_pieces = atom.run, []
        if atom.kind == 't':
            if cur_pieces and cur_pieces[-1][0] == 't':
                cur_pieces[-1][1] += atom.ch
            else:
                cur_pieces.append(['t', atom.ch])
        else:
            cur_pieces.append(['tab'])
    if cur_pieces:
        runs.append(_build_run_from_pieces(cur_run, cur_pieces))
    return runs


def _replace_runs_in_place(
    para: etree._Element,
    old_runs: list[etree._Element],
    new_runs: list[etree._Element],
) -> None:
    """Xoá đúng old_runs (list <w:r> hiện có, đúng thứ tự, trong para)
    khỏi para, chèn new_runs vào ĐÚNG vị trí old_runs[0] từng đứng. Nếu
    old_runs rỗng (para vốn không có run nào bị động tới ở phần này),
    chèn new_runs ngay sau <w:pPr> nếu có, đầu đoạn nếu không. pPr và mọi
    phần tử KHÁC của para (kể cả run không góp ký tự nào, vd công thức
    MathType/OLE) không bị đụng tới."""
    if old_runs:
        anchor = old_runs[0]
        for r in new_runs:
            anchor.addprevious(r)
        for r in old_runs:
            r.getparent().remove(r)
    else:
        all_children = list(para)
        ppr_idx = next((i for i, c in enumerate(all_children) if c.tag == w('pPr')), -1)
        insert_at = ppr_idx + 1
        for offset, r in enumerate(new_runs):
            para.insert(insert_at + offset, r)


def _set_paragraph_format(para: etree._Element, format_ppr: etree._Element | None) -> None:
    """Đặt <w:pPr> của para thành một deep-copy của format_ppr (hoặc xoá
    hẳn nếu format_ppr là None). pPr luôn là con đầu tiên của <w:p>."""
    old_ppr = para.find(w('pPr'))
    if old_ppr is not None:
        para.remove(old_ppr)
    if format_ppr is not None:
        para.insert(0, copy.deepcopy(format_ppr))


def _copy_paragraph_format_from(para: etree._Element, template_para: etree._Element) -> None:
    """Sao chép nguyên định dạng đoạn từ template_para sang para."""
    _set_paragraph_format(para, template_para.find(w('pPr')))


def _new_paragraph_from_template(
    template_para: etree._Element, runs: list[etree._Element]
) -> etree._Element:
    """Dựng 1 <w:p> MỚI mang deep-copy <w:pPr> của đoạn tương ứng trong
    file thay thế, rồi chèn runs vào đó. Không kế thừa pPr từ đoạn nguồn:
    đoạn mới là một phần của replacement nên phải giữ đúng căn lề/thụt
    lề/style của replacement."""
    new_p = etree.Element(w('p'))
    ppr = template_para.find(w('pPr'))
    if ppr is not None:
        new_p.append(copy.deepcopy(ppr))
    for r in runs:
        new_p.append(r)
    return new_p


def _has_unmatched_text_in_para(
    atoms: list[_Atom], para: etree._Element, start: int, end: int
) -> bool:
    """True nếu para vẫn còn chữ/tab nằm ngoài vùng [start, end). Khi còn
    nội dung nguồn này, pPr của para phải được giữ để không làm thay đổi bố
    cục của tiền tố/hậu tố không được thay."""
    return any(
        atom.para is para and atom.kind != 'pbreak' and not (start <= i < end)
        for i, atom in enumerate(atoms)
    )


def _splice_match(
    atoms: list[_Atom],
    start: int,
    end: int,
    replacement_paras: list[etree._Element],
) -> tuple[etree._Element, etree._Element | None, etree._Element, etree._Element | None]:
    """Xử lý 1 chỗ khớp [start, end) trong atoms = _document_atoms(...)
    của TOÀN BỘ đoạn văn nguồn HIỆN TẠI. SỬA TRỰC TIẾP lên cây XML -- có
    thể động tới NHIỀU đoạn văn nếu chỗ khớp vắt qua 1 hay nhiều ranh
    giới đoạn (atom kind='pbreak' nằm trong [start, end)).

    CHỈ động tới ĐÚNG những <w:r> có ký tự nằm trong [start, end) và đúng
    những <w:p> nằm HOÀN TOÀN trong khoảng đó (bị xoá nguyên phần tử) --
    mọi run/đoạn văn KHÁC (kể cả công thức MathType/OLE không góp ký tự
    nào) KHÔNG BAO GIỜ bị đụng tới, dù nằm ở vị trí nào.

    Số đoạn văn KẾT QUẢ tại chỗ này = max(1, số đoạn văn của file thay
    thế): mọi đoạn replacement có pPr được deep-copy từ CHÍNH file thay;
    pPr đoạn nguồn chỉ được giữ khi nó còn chữ/tab ngoài vùng khớp. Nếu
    file thay thế có NHIỀU đoạn hơn số đoạn nguồn bị khớp, đoạn nguồn cuối
    cùng sẽ được TÁCH thêm; nếu ÍT đoạn hơn (kể cả 0, tức xoá hẳn không
    thay gì), các đoạn nguồn thừa sẽ được gộp/xoá.

    Trả về (start_para, first_repl_run, end_para, last_repl_run) -- đoạn
    văn + <w:r> ĐẦU TIÊN của phần vừa chèn (dùng để áp định dạng LÙI VỀ
    đầu đoạn start_para), và đoạn văn + <w:r> CUỐI CÙNG của phần vừa chèn
    (dùng để áp định dạng TIẾN TỚI stop-text/cuối đoạn end_para) -- 2 cặp
    này CÓ THỂ là cùng 1 đoạn văn (chỗ khớp không vắt qua đoạn nào) hoặc
    khác nhau. Cả 2 run = None nếu file thay thế không có run nào (xoá
    hẳn, không có định dạng nào để lan)."""
    first_para = atoms[start].para
    last_para = atoms[end - 1].para
    same_para = last_para is first_para
    n_repl = len(replacement_paras)
    first_has_unmatched_text = _has_unmatched_text_in_para(
        atoms, first_para, start, end
    )
    last_has_unmatched_text = (
        first_has_unmatched_text if same_para else
        _has_unmatched_text_in_para(atoms, last_para, start, end)
    )

    if not same_para:
        p = first_para.getnext()
        while p is not None and p is not last_para:
            nxt = p.getnext()
            if p.tag == w('p'):
                p.getparent().remove(p)
            p = nxt

    boundary_start_run = atoms[start].run if atoms[start].kind != 'pbreak' else None
    boundary_end_run = atoms[end - 1].run if atoms[end - 1].kind != 'pbreak' else None

    prefix_atoms = (
        [atoms[i] for i in range(start) if atoms[i].run is boundary_start_run]
        if boundary_start_run is not None else []
    )
    prefix_runs = _atoms_to_run_pieces(prefix_atoms)

    suffix_needs_move = (same_para and n_repl >= 2) or (not same_para and n_repl <= 1)
    if suffix_needs_move:
        suffix_atoms = [atoms[i] for i in range(end, len(atoms)) if atoms[i].para is last_para]
    else:
        suffix_atoms = (
            [atoms[i] for i in range(end, len(atoms)) if atoms[i].run is boundary_end_run]
            if boundary_end_run is not None else []
        )
    suffix_runs = _atoms_to_run_pieces(suffix_atoms)

    first_para_touched: list[etree._Element] = []
    last_para_touched: list[etree._Element] = []
    seen_ids = set()
    for i in range(start, end):
        a = atoms[i]
        if a.kind == 'pbreak' or id(a.run) in seen_ids:
            continue
        seen_ids.add(id(a.run))
        if a.para is first_para:
            first_para_touched.append(a.run)
        elif a.para is last_para:
            last_para_touched.append(a.run)

    if suffix_needs_move and same_para:

        seen2 = {id(r) for r in first_para_touched}
        for a in suffix_atoms:
            if id(a.run) not in seen2:
                seen2.add(id(a.run))
                first_para_touched.append(a.run)

    repl_runs_per_para = [
        [copy.deepcopy(r) for r in rp.iter(w('r'))] for rp in replacement_paras
    ]

    if n_repl <= 1:
        middle = repl_runs_per_para[0] if n_repl == 1 else []
        _replace_runs_in_place(first_para, first_para_touched, prefix_runs + middle + suffix_runs)
        if n_repl == 1 and not first_has_unmatched_text and not last_has_unmatched_text:
            _copy_paragraph_format_from(first_para, replacement_paras[0])
        first_repl_run = middle[0] if middle else None
        last_repl_run = middle[-1] if middle else None
        if not same_para:
            last_para.getparent().remove(last_para)
        return first_para, first_repl_run, first_para, last_repl_run

    _replace_runs_in_place(first_para, first_para_touched, prefix_runs + repl_runs_per_para[0])
    if not first_has_unmatched_text:
        _copy_paragraph_format_from(first_para, replacement_paras[0])
    first_repl_run = repl_runs_per_para[0][0] if repl_runs_per_para[0] else None

    insert_after = first_para
    for i, mid_runs in enumerate(repl_runs_per_para[1:-1], start=1):
        new_p = _new_paragraph_from_template(replacement_paras[i], mid_runs)
        insert_after.addnext(new_p)
        insert_after = new_p

    last_runs = repl_runs_per_para[-1] + suffix_runs
    if same_para:

        new_last = _new_paragraph_from_template(replacement_paras[-1], last_runs)
        insert_after.addnext(new_last)
        end_para = new_last
    else:
        _replace_runs_in_place(last_para, last_para_touched, last_runs)
        if not last_has_unmatched_text:
            _copy_paragraph_format_from(last_para, replacement_paras[-1])
        end_para = last_para
    last_repl_run = repl_runs_per_para[-1][-1] if repl_runs_per_para[-1] else None
    return first_para, first_repl_run, end_para, last_repl_run


def _set_run_format(run: etree._Element, format_rpr: etree._Element | None) -> None:
    """Đặt <w:rPr> của run thành 1 bản deep-copy của format_rpr (hoặc XOÁ
    hẳn rPr nếu format_rpr là None) -- LUÔN THAY THẾ (không gộp) rPr cũ."""
    old_rpr = run.find(w('rPr'))
    if old_rpr is not None:
        run.remove(old_rpr)
    if format_rpr is not None:
        run.insert(0, copy.deepcopy(format_rpr))


def _apply_format_backward(
    start_para: etree._Element,
    first_repl_run: etree._Element | None,
    target_format: etree._Element | None,
    stop_text: str,
) -> None:
    """Áp target_format (rPr hoặc None) cho chữ trong start_para, TÍNH
    LÙI TỪ NGAY TRƯỚC first_repl_run (phần vừa chèn từ file thay thế) VỀ
    ĐẦU ĐOẠN VĂN. Nếu stop_text có giá trị, dừng NGAY SAU lần xuất hiện
    GẦN NHẤT của nó khi đi lùi từ phần thay (không áp định dạng lên chính
    chuỗi dừng); nếu rỗng hoặc không tìm thấy, áp tới đầu đoạn. Nếu điểm
    dừng rơi giữa 1 run, run được tách để phần trước chuỗi dừng không bị
    đổi định dạng. KHÔNG BAO GIỜ vượt sang đoạn văn trước đó.
    Run KHÔNG có chữ (công thức MathType/OLE, hình vẽ) không bao giờ bị
    áp định dạng này."""
    if first_repl_run is None:
        return
    atoms = _paragraph_atoms(start_para)
    positions = [i for i, a in enumerate(atoms) if a.run is first_repl_run]
    end_idx = positions[0] if positions else len(atoms)

    if not stop_text:
        start_idx = 0
    else:
        head_text = ''.join(a.ch for a in atoms[:end_idx])
        pos = head_text.rfind(stop_text)
        start_idx = pos + len(stop_text) if pos != -1 else 0

    if start_idx >= end_idx:
        return

    boundary_run = atoms[start_idx].run
    if boundary_run is not None and any(
        atoms[i].run is boundary_run for i in range(0, start_idx)
    ):
        _split_run_at_atom(atoms, boundary_run, start_idx)
        atoms = _paragraph_atoms(start_para)

    touched_ids = {id(atoms[i].run) for i in range(start_idx, end_idx)}
    if not touched_ids:
        return
    for r in start_para.iter(w('r')):
        if id(r) in touched_ids and any(c.tag in (w('t'), w('tab')) for c in r):
            _set_run_format(r, target_format)


def _split_run_at_atom(atoms: list[_Atom], run: etree._Element, split_atom_idx: int) -> None:
    """Tách VẬT LÝ 1 run thành 2 <w:r> riêng (CÙNG rPr với run gốc) ngay
    tại vị trí split_atom_idx (chỉ số trong atoms) -- dùng để áp định
    dạng khác nhau cho 2 nửa (vd điểm dừng format_until_text rơi vào
    giữa 1 run). Không làm gì nếu split_atom_idx không thực sự rơi vào
    GIỮA run (run không có atom nào cả trước lẫn từ split_atom_idx)."""
    run_indices = [i for i, a in enumerate(atoms) if a.run is run]
    if not run_indices:
        return
    before = [i for i in run_indices if i < split_atom_idx]
    after = [i for i in run_indices if i >= split_atom_idx]
    if not before or not after:
        return

    before_pieces = _atoms_to_run_pieces([atoms[i] for i in before])
    after_pieces = _atoms_to_run_pieces([atoms[i] for i in after])
    for r in before_pieces + after_pieces:
        run.addprevious(r)
    run.getparent().remove(run)


def _apply_format_forward(
    end_para: etree._Element,
    last_repl_run: etree._Element | None,
    target_format: etree._Element | None,
    stop_text: str,
) -> None:
    """Áp target_format (rPr hoặc None) cho chữ trong end_para, bắt đầu
    NGAY SAU last_repl_run (phần vừa chèn từ file thay thế), đi TỚI
    TRƯỚC, dừng lại khi:
      - stop_text rỗng ("" hoặc None): hết đoạn văn end_para (gặp Enter),
        KHÔNG dừng ở chỗ Word tự xuống hàng do hết chiều rộng.
      - stop_text có giá trị: gặp ĐÚNG chuỗi đó lần đầu tiên tính từ chỗ
        vừa thay (dừng NGAY TRƯỚC chuỗi đó) -- so khớp THUẦN VĂN BẢN,
        KHÔNG xét định dạng (khác cách so khớp file tìm/thay ở trên). Nếu
        điểm dừng rơi vào GIỮA 1 run, run đó được TÁCH VẬT LÝ tại đúng
        ranh giới (xem _split_run_at_atom()) để chỉ đúng phần trước điểm
        dừng đổi định dạng, phần từ điểm dừng trở đi (cùng run gốc) giữ
        NGUYÊN định dạng cũ. Không tìm thấy trong phần còn lại của đoạn
        văn: áp tới hết đoạn văn (coi như không có stop_text). KHÔNG BAO
        GIỜ vượt sang đoạn văn sau, kể cả khi không tìm thấy stop_text.
    Run KHÔNG có chữ (công thức MathType/OLE, hình vẽ) không bao giờ bị
    áp định dạng này."""
    if last_repl_run is None:
        return
    atoms = _paragraph_atoms(end_para)
    positions = [i for i, a in enumerate(atoms) if a.run is last_repl_run]
    start_idx = (positions[-1] + 1) if positions else len(atoms)

    if not stop_text:
        stop_idx = len(atoms)
    else:
        tail_text = ''.join(a.ch for a in atoms[start_idx:])
        pos = tail_text.find(stop_text)
        stop_idx = start_idx + pos if pos != -1 else len(atoms)

    if start_idx >= stop_idx:
        return

    if stop_idx < len(atoms):
        boundary_run = atoms[stop_idx].run
        if boundary_run is not None and any(
            atoms[i].run is boundary_run for i in range(start_idx, stop_idx)
        ):
            _split_run_at_atom(atoms, boundary_run, stop_idx)
            atoms = _paragraph_atoms(end_para)

    touched_ids = {id(atoms[i].run) for i in range(start_idx, min(stop_idx, len(atoms)))}
    if not touched_ids:
        return
    for r in end_para.iter(w('r')):
        if id(r) in touched_ids and any(c.tag in (w('t'), w('tab')) for c in r):
            _set_run_format(r, target_format)


MAX_DOCUMENT_ITERATIONS = 5000


def _iter_paragraph_groups(container: etree._Element) -> Iterator[list[etree._Element]]:
    """Sinh lần lượt từng NHÓM (list) các <w:p> LIÊN TIẾP là con trực
    tiếp của container (w:body của toàn tài liệu, hoặc 1 ô bảng w:tc) --
    1 nhóm = 1 "luồng" văn bản ĐỘC LẬP, nơi chỗ khớp CÓ THỂ vắt qua ranh
    giới đoạn văn bên trong nhóm đó (xem _document_atoms()).

    Gặp 1 bảng (w:tbl) xen giữa: nhóm HIỆN TẠI kết thúc ngay tại đó (bảng
    LUÔN ngắt luồng văn bản -- chỗ khớp không bao giờ vắt từ đoạn văn
    trước bảng sang đoạn văn sau bảng, hay từ ô bảng này sang ô bảng
    khác, kể cả 2 ô liền kề cùng hàng, giống cảm giác trực quan mỗi ô là
    1 "hộp" nội dung tách biệt), rồi ĐỆ QUY vào TỪNG ô của bảng đó (theo
    đúng thứ tự hàng rồi cột), coi mỗi ô là 1 container riêng -- 1 ô lại
    có thể sinh ra NHIỀU nhóm nữa nếu bên trong nó có bảng lồng. Nhờ đệ
    quy này, bảng lồng trong ô bảng được xử lý đúng, không giới hạn độ
    sâu.

    Mọi thẻ KHÁC không phải <w:p>/<w:tbl> (vd <w:sectPr>, bookmark, dấu
    vết track-changes...) trong SUỐT với việc gom nhóm -- bị bỏ qua HOÀN
    TOÀN như thể không tồn tại, không ngắt nhóm -- giữ đúng hành vi lọc
    vốn có từ trước (khi chưa hỗ trợ bảng, code cũ cũng chỉ lọc lấy
    <w:p>, bỏ qua mọi thẻ khác).

    Thứ tự sinh ra = ĐÚNG thứ tự xuất hiện trong tài liệu (nhóm văn bản
    trước 1 bảng, rồi lần lượt từng ô của bảng đó, rồi nhóm văn bản sau
    bảng...) -- dùng làm thứ tự "khớp trước" khi tìm chỗ khớp đầu tiên
    trong toàn tài liệu, xem _find_match_in_document()."""
    current_group: list[etree._Element] = []
    for child in container:
        if child.tag == w('p'):
            current_group.append(child)
        elif child.tag == w('tbl'):
            if current_group:
                yield current_group
                current_group = []
            for tr in child:
                if tr.tag != w('tr'):
                    continue
                for tc in tr:
                    if tc.tag != w('tc'):
                        continue
                    yield from _iter_paragraph_groups(tc)
    if current_group:
        yield current_group


def _find_match_in_document(
    source_root: etree._Element, pattern_keys: list[tuple[str, str | None]]
) -> tuple[list[_Atom], int, int] | None:
    """Tìm chỗ khớp ĐẦU TIÊN của pattern_keys, xét theo ĐÚNG thứ tự xuất
    hiện trong toàn tài liệu (xem _iter_paragraph_groups()). Trả về
    (atoms, start, end) của NHÓM đoạn văn (1 luồng -- thân trang hoặc 1 ô
    bảng) chứa chỗ khớp đó, hoặc None nếu không tìm thấy ở bất kỳ đâu.
    Duyệt từng nhóm theo thứ tự, dừng ngay ở nhóm ĐẦU TIÊN có khớp (khỏi
    cần gộp atoms của cả tài liệu vào 1 mảng khổng lồ) -- ĐÚNG vì các
    nhóm vốn đã theo thứ tự tài liệu."""
    body = source_root.find(w('body'))
    for group in _iter_paragraph_groups(body):
        atoms = _document_atoms(group)
        match = _find_first_atom_match(atoms, pattern_keys)
        if match is not None:
            start, end = match
            return atoms, start, end
    return None


def _replace_all_matches(
    source_root: etree._Element,
    pattern_keys: list[tuple[str, str | None]],
    replacement_paras: list[etree._Element],
    backward_stop_text: str,
    forward_stop_text: str,
) -> int:
    """Tìm & thay HẾT mọi chỗ khớp trong TOÀN BỘ nội dung file (thân
    trang LẪN mọi ô bảng, kể cả bảng lồng) -- 1 CƠ CHẾ DUY NHẤT bất kể
    chỗ khớp nằm gọn trong 1 đoạn hay vắt qua nhiều đoạn văn của CÙNG 1
    luồng (xem _iter_paragraph_groups()). Mỗi lần TÍNH LẠI toàn bộ nhóm
    đoạn văn + atoms từ đầu (theo cấu trúc HIỆN TẠI của source_root,
    khỏi phải tự dịch chỉ số thủ công sau mỗi lần sửa -- cùng cách làm
    với transform_mtef() trong mtef_transform.py), thay 1 chỗ khớp
    ĐẦU TIÊN, lặp lại.

    CHỈ SAU KHI thay hết mọi chỗ khớp mới áp định dạng LÙI và TIẾN (tới
    chuỗi dừng riêng của từng chiều, hoặc đầu/cuối đoạn) cho TỪNG chỗ vừa thay (theo đúng
    thứ tự đã thay) -- hoãn lại để lần tìm KẾ TIẾP trong lúc đang thay
    không bị ảnh hưởng bởi định dạng vừa áp (nếu áp ngay sau mỗi lần
    thay, 1 số chỗ khớp còn lại có thể không còn đúng định dạng file tìm
    nữa nếu định dạng file thay khác định dạng file tìm, và dừng sớm hơn
    dự kiến).

    Trả về số chỗ đã thay."""
    count = 0
    anchors: list[tuple[etree._Element, etree._Element | None, etree._Element, etree._Element | None]] = []
    for _ in range(MAX_DOCUMENT_ITERATIONS):
        found = _find_match_in_document(source_root, pattern_keys)
        if found is None:
            break
        atoms, start, end = found
        anchors.append(_splice_match(atoms, start, end, replacement_paras))
        count += 1
    else:
        raise RuntimeError(
            "Quá nhiều lần thay thế liên tiếp -- dừng lại để an toàn (có "
            "thể chuỗi thay thế trùng ngay chuỗi cần tìm, gây lặp vô hạn)."
        )

    if replacement_paras:
        first_repl_run = next(replacement_paras[0].iter(w('r')), None)
        target_format = first_repl_run.find(w('rPr')) if first_repl_run is not None else None
        for start_para, first_repl_run_in_para, end_para, last_repl_run in anchors:
            _apply_format_backward(
                start_para, first_repl_run_in_para, target_format, backward_stop_text
            )
            _apply_format_forward(end_para, last_repl_run, target_format, forward_stop_text)

    return count


def _apply_one_pair(
    source_root: etree._Element,
    find_path: str,
    replacement_path: str,
    backward_stop_text: str,
    forward_stop_text: str,
) -> int:
    """Lõi xử lý 1 cặp: tìm nội dung find_path trong source_root (lxml
    Element <w:document> đã đọc sẵn trong bộ nhớ), thay bằng nội dung
    replacement_path, SỬA TRỰC TIẾP lên source_root. Không đọc/ghi file
    zip nào ở đây -- tách riêng để replace_docx() có thể gọi lại hàm này
    liên tiếp cho nhiều cặp trên CÙNG 1 source_root đang có trong bộ nhớ.
    Trả về số chỗ đã thay của riêng cặp này."""
    for p in (find_path, replacement_path):
        check_is_real_docx(p)

    find_paras = get_content_paragraphs(read_document_xml_root(find_path))
    replacement_paras = get_content_paragraphs(read_document_xml_root(replacement_path))
    if not find_paras:
        raise ValueError("File tìm không có nội dung nào để tìm.")

    pattern_atoms = _document_atoms(find_paras)
    if not pattern_atoms:
        raise ValueError("File tìm không có nội dung nào để tìm.")
    pattern_keys = [(a.ch, a.fmt) for a in pattern_atoms]

    return _replace_all_matches(
        source_root,
        pattern_keys,
        replacement_paras,
        backward_stop_text,
        forward_stop_text,
    )


def _write_docx_with_new_document_xml(
    source_path: str, source_root: etree._Element, out_path: str
) -> None:
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


def replace_docx(
    source_path: str,
    pairs: list[FindReplacePair],
    out_path: str,
    log: Callable[[str], None] | None = None,
) -> list[int]:
    """Thay nội dung theo TỪNG BỘ (find_path, replacement_path,
    backward_stop_text, forward_stop_text) trong pairs, chạy TUẦN TỰ trên source_path (bộ
    sau thấy được kết quả bộ trước đã sửa VÀ đã áp định dạng xong), xuất
    kết quả CUỐI CÙNG ra out_path. Trả về list số chỗ đã thay, đúng theo
    thứ tự pairs.

    log(text): callback ghi 1 dòng log NGAY tại thời điểm xảy ra (đang xử
    lý cặp thứ mấy/tổng số mấy, kết quả từng cặp -- đã thay bao nhiêu chỗ,
    hay không khớp chỗ nào...) -- mặc định no-op (không log gì) nếu không
    truyền, cùng convention với fix_mathtype_parens_in_docx() trong
    fix_mathtype_parens.py.

    Hai chuỗi dừng lan định dạng đi RIÊNG theo từng bộ: backward_stop_text
    chặn chiều lùi và forward_stop_text chặn chiều tiến. Để trống một ô
    thì chiều đó lan tới đầu/cuối đoạn tương ứng.

    Danh sách chỉ có 1 bộ vẫn gọi hàm này -- không có hàm riêng cho trường hợp 1 cặp, để
    tránh 2 hàm gần giống nhau cho cùng 1 việc.

    NGUYÊN TẮC MỖI CẶP (1 cơ chế duy nhất, áp dụng cho pairs):
      1. Coi nội dung file tìm là 1 CHUỖI KÝ TỰ CÓ ĐỊNH DẠNG: mỗi ký tự
          (kể cả tab, w:tab) kèm định dạng run có ảnh hưởng hiển thị
          (rPr: đậm/nghiêng/gạch chân/font/màu...). Định dạng đoạn (pPr:
          căn lề, thụt lề, giãn dòng, tab, style...) không tham gia so
          khớp, để chuỗi tìm được ở mọi vị trí. Nếu file tìm có TỪ 2 đoạn
          văn trở lên, ranh giới đoạn văn vẫn được coi là 1 "ký tự" phải
          khớp.
      2. Tìm chuỗi này ở BẤT KỲ vị trí nào trong file gốc -- thân trang
         lẫn trong từng ô bảng (kể cả bảng lồng) -- có thể nằm gọn trong
         1 đoạn, hoặc vắt qua nhiều đoạn văn liên tiếp CỦA CÙNG 1 LUỒNG
         (thân trang, hoặc 1 ô bảng cụ thể), KHÔNG cần biết/khai báo
         trước là trường hợp nào. KHÔNG BAO GIỜ vắt từ đoạn văn trước 1
         bảng sang đoạn văn sau bảng, hay từ ô bảng này sang ô bảng khác
         (kể cả 2 ô liền kề cùng hàng) -- xem _iter_paragraph_groups().
          Chỉ khớp nếu CẢ chữ LẪN định dạng run (và ranh giới đoạn văn,
          nếu pattern có) giống hệt, liên tục -- khác Tìm & Thay của Word
          ở chỗ Word mặc định chỉ cần chứa đúng CHỮ, không quan tâm định
          dạng. Có thể khớp NHIỀU lần (thay hết, không chồng lấn lên nhau).
      3. Thay đúng phần khớp bằng nội dung file thay thế, GIỮ NGUYÊN
          định dạng riêng của replacement ở CẢ run lẫn đoạn văn. Mỗi đoạn
          replacement được deep-copy w:pPr riêng (căn lề, thụt lề, style)
          vào đoạn kết quả tương ứng. Duy nhất đoạn nguồn còn tiền tố/hậu
          tố chữ/tab nằm ngoài vùng khớp thì giữ pPr gốc để bố cục của phần
          không thay không bị đổi. Số đoạn văn tại chỗ này sau khi thay =
          số đoạn văn của file thay thế: nếu file thay thế có NHIỀU đoạn
          hơn số đoạn nguồn bị khớp, đoạn nguồn sẽ được TÁCH thêm ra; nếu
          ÍT đoạn hơn (kể cả 0 -- xoá hẳn không thay gì), các đoạn nguồn
          thừa sẽ được gộp/xoá.
      4. backward_stop_text và forward_stop_text CỦA BỘ ĐÓ (mặc định rỗng
         nếu để trống): SAU
         KHI thay xong TOÀN BỘ chỗ khớp của RIÊNG bộ này (chưa chạy sang
         bộ kế tiếp), định dạng của RUN ĐẦU TIÊN trong file thay thế được
         áp cho chữ liền kề CẢ HAI PHÍA trong đoạn văn: lùi trước phần thay
         và tiến sau phần thay. Dừng lại khi:
            - rỗng: đầu/cuối đoạn văn (gặp Enter), không phải chỗ Word tự
              xuống hàng do hết chiều rộng.
           - có giá trị: gặp ĐÚNG chuỗi đó gần nhất khi đi lùi, hoặc lần
             đầu khi đi tiến (so THUẦN VĂN BẢN, không xét định dạng). Bản
             thân chuỗi dừng giữ nguyên định dạng cũ. Không tìm thấy thì
             áp tới đầu/cuối đoạn tương ứng; không bao giờ vượt qua ranh
             giới đoạn văn.
         Công thức MathType/OLE object hay hình vẽ nhúng không bao giờ bị
         áp định dạng này.

    Nếu 1 bộ lỗi giữa chừng (vd file không đọc được, không có nội dung để
    tìm...): DỪNG NGAY, không chạy bộ sau, KHÔNG ghi ra out_path -- file
    kết quả chỉ được ghi khi TOÀN BỘ danh sách chạy xong không lỗi, tránh
    để lại file dở dang. Lỗi ném lại kèm số thứ tự bộ và tên 2 file, để
    biết chính xác bộ nào cần sửa.

    PHẠM VI (cố ý như vậy, không phải thiếu sót):
      - CHỈ xử lý văn bản trong THÂN tài liệu (w:body), kể cả trong bảng
        (xem điểm 2 ở trên). Header, footer, viền trang, đánh số trang
        nằm ở phần khác của file .docx -- không bao giờ bị đụng tới.
      - KHÔNG hỗ trợ ảnh/OLE object (kể cả công thức MathType) nhúng BÊN
        TRONG đoạn văn của file thay thế: file media/quan hệ liên quan
        không được copy theo (chỉ document.xml đổi), tham chiếu sẽ vỡ,
        nội dung nhúng biến mất im lặng trong kết quả. File thay thế chỉ
        nên là văn bản có định dạng thuần.
      - Style đoạn văn (w:pStyle) được giữ dưới dạng THAM CHIẾU tên/ID,
        chỉ hiển thị đúng nếu file gốc có định nghĩa style cùng tên trong
        styles.xml của nó.
      - Nội dung run KHÁC w:t/w:tab (vd <w:br/> xuống dòng cứng) nằm
        trong 1 run bị chỗ khớp cắt ngang giữa chừng hiện KHÔNG được bảo
        toàn khi tách run.
    """
    log = log or (lambda _msg: None)
    if not pairs:
        raise ValueError("Danh sách cặp tìm/thay đang trống.")

    check_is_real_docx(source_path)
    source_root = read_document_xml_root(source_path)

    counts = []
    for i, pair in enumerate(pairs, start=1):
        if len(pair) == 3:

            find_path, replacement_path, stop_text = pair
            backward_stop_text = forward_stop_text = stop_text
        elif len(pair) == 4:
            find_path, replacement_path, backward_stop_text, forward_stop_text = pair
        else:
            raise ValueError(
                f"Cặp thứ {i} phải có 3 hoặc 4 giá trị (nhận được {len(pair)})."
            )
        find_name = os.path.basename(find_path)
        replacement_name = os.path.basename(replacement_path)
        log(f'Cặp {i}/{len(pairs)}: đang tìm/thay ("{find_name}" → "{replacement_name}")...')
        try:
            count = _apply_one_pair(
                source_root,
                find_path,
                replacement_path,
                backward_stop_text,
                forward_stop_text,
            )
        except Exception as e:
            raise ValueError(
                f"[LỖI] Cặp {i}/{len(pairs)} (tìm: {find_name}, thay: {replacement_name}).\n"
                f"      chi tiết: {e}"
            ) from e
        counts.append(count)
        if count == 0:
            log(
                f"[BỎ QUA] Cặp {i}/{len(pairs)}: không tìm thấy chỗ nào khớp "
                f"(kiểm tra lại nội dung/định dạng file tìm)."
            )
        else:
            log(f"[OK] Cặp {i}/{len(pairs)}: đã thay {count} chỗ.")

    _write_docx_with_new_document_xml(source_path, source_root, out_path)
    return counts
