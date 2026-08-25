"""
mathtype_refresh.py
====================
Mở 1 file .docx bằng Word (COM, CHẠY ẨN) để buộc MathType vẽ lại ảnh xem
trước (WMF/EMF cache) cho các công thức mà fix_mathtype_parens.py vừa sửa
ngoặc. fix_mathtype_parens.py chỉ sửa dữ liệu nhị phân MTEF -- không tự vẽ
lại ảnh xem trước hiển thị khi KHÔNG double-click vào công thức; cách duy
nhất để buộc vẽ lại là Activate() từng công thức trong MathType, bấm Update
rồi Close and Return.

KHỚP CÔNG THỨC CẦN REFRESH VỚI ĐÚNG InlineShape TRONG WORD: đầu vào chỉ có
TÊN FILE (oleObjectN.bin), còn Word/COM chỉ đưa ra 1 danh sách InlineShape
theo THỨ TỰ, không kèm tên file. Đọc word/document.xml +
word/_rels/document.xml.rels để lấy thứ tự tên file, ghép 1-1 với danh
sách InlineShape lọc cùng tiêu chí ProgID. Số lượng 2 bên lệch nhau: DỪNG
LẠI, không refresh công thức nào -- an toàn hơn đoán bừa sai công thức.

Bấm menu MathType (Update / Close and Return) qua PostMessage theo VỊ TRÍ
cố định trên menu (KHÔNG so tên nút, tránh lệch ngôn ngữ hiển thị).

Yêu cầu: pip install pywin32. Chỉ chạy được trên Windows, máy phải có Word
+ MathType cài sẵn. Thiếu pywin32 hoặc COM lỗi: KHÔNG raise, trả về
RefreshReport ghi rõ lý do bỏ qua -- vì file .docx kết quả (MTEF đã sửa
đúng) đã được ghi ra XONG từ bước trước, refresh chỉ là bước "làm đẹp
thêm" ảnh xem trước.

LƯU Ý: tính năng tự đóng hộp thoại "License Information" của MathType
(nhắc dùng thử, thỉnh thoảng bật lên khi Activate()) đã được TẠM BỎ khỏi
bản này -- sẽ làm lại sau. Nếu hộp thoại đó xuất hiện trong lúc chạy, công
thức đang xử lý sẽ time-out chờ cửa sổ MathType
(MATHTYPE_APPEAR_TIMEOUT_SECONDS) và bị đưa vào skipped, cho tới khi người
dùng tự đóng hộp thoại bằng tay.
"""
from __future__ import annotations

import os
import time
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from lxml import etree

try:
    import win32com.client
    import win32con
    import win32gui
except ImportError:
    # Cả 3 tên cùng về None -- tránh NameError khó hiểu nếu lỡ gọi nhầm 1
    # trong 3 tên mà quên kiểm tra win32com is None trước.
    win32com = None
    win32con = None
    win32gui = None

O_URI = 'urn:schemas-microsoft-com:office:office'
R_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_OLE_OBJECT_TAG = f'{{{O_URI}}}OLEObject'
_R_ID_ATTR = f'{{{R_URI}}}id'
DOCUMENT_PART = 'word/document.xml'
RELS_PART = 'word/_rels/document.xml.rels'

WD_INLINE_SHAPE_OLE = 1
EQUATION_PROGID_PREFIXES = ("Equation.DSMT", "Equation.3")

MATHTYPE_TITLE_SUBSTRING = "mathtype"
MATHTYPE_APPEAR_TIMEOUT_SECONDS = 6.0
MATHTYPE_CLOSE_TIMEOUT_SECONDS = 3.0
POLL_SECONDS = 0.3
POST_UPDATE_WAIT_SECONDS = 1.0
POST_CLOSE_WAIT_SECONDS = 0.5

# Vị trí (index) CỐ ĐỊNH trên menu MathType -- không so tên nút, để không
# lệch khi giao diện MathType hiển thị Anh/Việt khác nhau tuỳ máy.
MENU_TAB_INDEX = 0
MENU_ITEM_UPDATE = 3
MENU_ITEM_CLOSE_RETURN = 2


def _equation_bin_names_in_document_order(docx_path: str) -> list[str]:
    """Tên file (vd "oleObject3.bin") của các công thức MathType/Equation
    Editor, THEO ĐÚNG THỨ TỰ xuất hiện trong word/document.xml. Đọc thẻ
    <o:OLEObject Type="Embed" ProgID="..." r:id="..."/> có ProgID khớp
    EQUATION_PROGID_PREFIXES, giải r:id qua word/_rels/document.xml.rels
    để lấy tên file cuối cùng của Target quan hệ đó."""
    with zipfile.ZipFile(docx_path) as z:
        doc_root = etree.fromstring(z.read(DOCUMENT_PART))
        rels_root = etree.fromstring(z.read(RELS_PART))

    rel_targets = {rel.get('Id'): rel.get('Target') for rel in rels_root}

    names: list[str] = []
    for ole in doc_root.iter(_OLE_OBJECT_TAG):
        prog_id = ole.get('ProgID', '')
        if not any(prog_id.startswith(p) for p in EQUATION_PROGID_PREFIXES):
            continue
        target = rel_targets.get(ole.get(_R_ID_ATTR))
        if target:
            names.append(os.path.basename(target))
    return names


def _is_equation_shape(inline_shape: Any) -> bool:
    """True nếu 1 InlineShape (từ doc.InlineShapes qua COM) là công thức
    MathType/Equation Editor -- CÙNG tiêu chí lọc với
    _equation_bin_names_in_document_order(), để 2 danh sách khớp thứ tự."""
    if inline_shape.Type != WD_INLINE_SHAPE_OLE:
        return False
    try:
        prog_id = inline_shape.OLEFormat.ProgID
    except Exception:
        return False
    return any(prog_id.startswith(p) for p in EQUATION_PROGID_PREFIXES)


def _equation_shapes_in_order(doc: Any) -> list[Any]:
    """Danh sách InlineShape công thức trong doc (COM Document đang mở),
    theo đúng thứ tự doc.InlineShapes trả về."""
    return [s for s in doc.InlineShapes if _is_equation_shape(s)]


def _find_mathtype_window() -> int | None:
    """hwnd cửa sổ đầu tiên có tiêu đề chứa MATHTYPE_TITLE_SUBSTRING
    (không phân biệt hoa/thường), None nếu chưa thấy."""
    found: list[int] = []

    def callback(hwnd: int, _extra: object) -> bool:
        if MATHTYPE_TITLE_SUBSTRING in win32gui.GetWindowText(hwnd).lower():
            found.append(hwnd)
        return True

    win32gui.EnumWindows(callback, found)
    return found[0] if found else None


def _wait_for_mathtype_window(timeout: float = MATHTYPE_APPEAR_TIMEOUT_SECONDS) -> int | None:
    """Chờ tới khi cửa sổ MathType xuất hiện, thăm dò mỗi POLL_SECONDS,
    tối đa timeout giây. Trả về hwnd, hoặc None nếu hết giờ."""
    waited = 0.0
    while waited <= timeout:
        hwnd = _find_mathtype_window()
        if hwnd:
            return hwnd
        time.sleep(POLL_SECONDS)
        waited += POLL_SECONDS
    return None


def _wait_for_window_closed(hwnd: int, timeout: float = MATHTYPE_CLOSE_TIMEOUT_SECONDS) -> None:
    """Chờ tới khi hwnd không còn là cửa sổ hợp lệ (đã đóng), tối đa
    timeout giây -- không raise nếu hết giờ mà vẫn còn (nơi gọi tự xử lý
    tiếp theo bình thường)."""
    waited = 0.0
    while win32gui.IsWindow(hwnd) and waited <= timeout:
        time.sleep(POLL_SECONDS)
        waited += POLL_SECONDS


def _click_menu_by_index(hwnd: int, tab_index: int, item_index: int) -> bool:
    """Gửi WM_COMMAND tới hwnd theo ĐÚNG vị trí (tab_index, item_index)
    trên thanh menu, qua PostMessage (không cần hwnd đang focus/hiện). Trả
    về True nếu GỬI được lệnh (không đảm bảo MathType đã xử lý xong)."""
    h_menu = win32gui.GetMenu(hwnd)
    if not h_menu:
        return False
    h_submenu = win32gui.GetSubMenu(h_menu, tab_index)
    if not h_submenu:
        return False
    item_id = win32gui.GetMenuItemID(h_submenu, item_index)
    if item_id <= 0:
        return False
    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, item_id, 0)
    return True


def _refresh_one_shape(shape: Any, log: Callable[[str], None]) -> tuple[bool, str | None]:
    """Activate 1 InlineShape công thức, đợi cửa sổ MathType, bấm Update
    rồi Close and Return theo index cố định, đợi đóng hẳn. Trả về (ok,
    lý_do_hoặc_None) -- KHÔNG raise, để 1 công thức lỗi không chặn các
    công thức còn lại."""
    try:
        shape.OLEFormat.Activate()
    except Exception as e:
        log(f"    -> Không kích hoạt (Activate) được công thức: {e}")
        return False, f"không mở được công thức trong MathType\n      chi tiết: {e}"

    log("    -> Đang chờ cửa sổ MathType mở ra...")
    hwnd = _wait_for_mathtype_window()
    if hwnd is None:
        log("    -> Không thấy cửa sổ MathType xuất hiện, bỏ qua.")
        return False, "cửa sổ MathType không xuất hiện"

    log("    -> Đang bấm Update (đẩy dữ liệu mới + vẽ lại ảnh xem trước)...")
    if not _click_menu_by_index(hwnd, MENU_TAB_INDEX, MENU_ITEM_UPDATE):
        log("    -> Không gửi được lệnh Update.")
        return False, "không gửi được lệnh Update tới MathType"
    time.sleep(POST_UPDATE_WAIT_SECONDS)

    log("    -> Đang bấm Close and Return...")
    if not _click_menu_by_index(hwnd, MENU_TAB_INDEX, MENU_ITEM_CLOSE_RETURN):
        log("    -> Không gửi được lệnh Close and Return.")
        return False, "không gửi được lệnh Close and Return tới MathType"

    _wait_for_window_closed(hwnd)
    time.sleep(POST_CLOSE_WAIT_SECONDS)
    log("    -> Xong.")
    return True, None


class RefreshReport:
    """Kết quả 1 lượt refresh_equation_previews() trên 1 file .docx."""

    def __init__(self, total_targets: int, refreshed: list[str], skipped: list[tuple[str, str]]) -> None:
        """total_targets: số công thức cần vẽ lại ảnh. refreshed: tên file
        đã vẽ lại thành công. skipped: (tên_file, lý_do) chưa vẽ lại được
        -- KHÔNG ảnh hưởng dữ liệu MTEF đã sửa đúng từ trước, chỉ ảnh xem
        trước của công thức đó có thể còn hiện dạng cũ."""
        self.total_targets = total_targets
        self.refreshed = refreshed
        self.skipped = skipped

    def summary_lines(self) -> list[str]:
        """Danh sách dòng text theo nhãn [OK]/[BỎ QUA], cùng định dạng với
        FixReport.summary_lines()."""
        if self.total_targets == 0:
            return ["[OK] Vẽ lại ảnh xem trước MathType: không có công thức nào cần vẽ lại."]
        header = f"Vẽ lại ảnh xem trước MathType: {len(self.refreshed)}/{self.total_targets} công thức"
        lines = [f"[OK] {header}" if not self.skipped else header]
        for name, reason in self.skipped:
            lines.append(f"[BỎ QUA] {name}: {reason}")
        return lines


def refresh_equation_previews(
    docx_path: str,
    equation_bin_names: list[str],
    log: Callable[[str], None] | None = None,
) -> RefreshReport:
    """Mở docx_path bằng Word (COM, CHẠY ẨN), buộc MathType vẽ lại ảnh xem
    trước cho các công thức có tên file trong equation_bin_names; công
    thức khác giữ nguyên. SỬA TRỰC TIẾP lên docx_path -- gọi hàm này SAU
    KHI đã có file kết quả CUỐI CÙNG (xem pipeline.py).

    equation_bin_names: tên file (vd "oleObject3.bin"), thường lấy từ
    FixReport.fixed của fix_mathtype_parens_in_docx() -- chỉ công thức
    THỰC SỰ đã đổi ngoặc mới cần vẽ lại ảnh.

    log(text): callback ghi 1 dòng log ngay khi xảy ra; mặc định no-op.

    Trả về RefreshReport. KHÔNG raise cho lỗi có thể lường trước (thiếu
    pywin32, Word/MathType lỗi COM...) -- xem docstring module."""
    log = log or (lambda _msg: None)
    target_names = set(equation_bin_names)
    total = len(target_names)
    refreshed: list[str] = []
    skipped: list[tuple[str, str]] = []

    if total == 0:
        log("Không có công thức nào cần vẽ lại ảnh xem trước.")
        return RefreshReport(0, refreshed, skipped)

    if win32com is None:
        reason = "chưa cài pywin32 (pip install pywin32)"
        log(f"Bỏ qua vẽ lại ảnh xem trước: {reason}.")
        skipped.extend((n, reason) for n in target_names)
        return RefreshReport(total, refreshed, skipped)

    log("Đang đọc thứ tự công thức trong file (document.xml)...")
    try:
        ordered_names = _equation_bin_names_in_document_order(docx_path)
    except Exception as e:
        reason = f"không đọc được cấu trúc file để xác định thứ tự công thức\n      chi tiết: {e}"
        log(f"Bỏ qua vẽ lại ảnh xem trước: {reason}")
        skipped.extend((n, reason) for n in target_names)
        return RefreshReport(total, refreshed, skipped)

    word = None
    doc = None
    try:
        log("Đang mở Word (chạy ẩn, Visible=False)...")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        resolved_path = str(Path(docx_path).resolve())
        log(f"Đang mở file: {resolved_path}")
        doc = word.Documents.Open(resolved_path)

        shapes = _equation_shapes_in_order(doc)
        if len(shapes) != len(ordered_names):
            reason = (
                f"số công thức Word thấy ({len(shapes)}) khác số đọc được từ "
                f"document.xml ({len(ordered_names)}), không khớp thứ tự an toàn"
            )
            log(f"CẢNH BÁO: {reason} -- dừng lại, không vẽ lại ảnh nào.")
            skipped.extend((n, reason) for n in target_names)
            return RefreshReport(total, refreshed, skipped)

        name_to_shape = dict(zip(ordered_names, shapes))

        log(f"Bắt đầu vẽ lại ảnh cho {total} công thức đã sửa ngoặc...")
        for i, name in enumerate(sorted(target_names), start=1):
            log(f"[{i}/{total}] Công thức {name}:")
            shape = name_to_shape.get(name)
            if shape is None:
                log("    -> Không tìm thấy công thức này khi mở bằng Word.")
                skipped.append((name, "không tìm thấy trong file khi mở bằng Word"))
                continue
            ok, reason = _refresh_one_shape(shape, log)
            if ok:
                refreshed.append(name)
            else:
                skipped.append((name, reason))

        log("Đang lưu file...")
        doc.Save()
        log("Đã lưu.")

    except Exception as e:
        reason = f"gặp lỗi ngoài dự kiến khi điều khiển Word/MathType tự động\n      chi tiết: {e}"
        log(f"Dừng lại: {reason}")
        done_names = set(refreshed) | {n for n, _ in skipped}
        skipped.extend((n, reason) for n in target_names if n not in done_names)

    finally:
        log("Đang đóng Word...")
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass

    return RefreshReport(total, refreshed, skipped)
