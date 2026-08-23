"""
mathtype_refresh.py
====================
Mở 1 file .docx bằng Word (COM, CHẠY ẨN) để buộc MathType vẽ lại ảnh
xem trước (WMF/EMF cache) cho ĐÚNG các công thức mà fix_mathtype_parens.py
vừa sửa ngoặc -- công thức không đổi gì thì KHÔNG đụng tới (ảnh cũ của nó
vẫn đúng, khỏi cần vẽ lại, đỡ tốn thời gian mở Word cho từng công thức).

VÌ SAO CẦN BƯỚC NÀY: fix_mathtype_parens.py chỉ sửa dữ liệu nhị phân MTEF
bên trong oleObjectN.bin -- không tự vẽ lại ảnh xem trước mà Word hiển
thị khi KHÔNG double-click vào công thức. Ảnh cũ (ngoặc tự co giãn) vẫn
còn nguyên cho tới khi MathType tự vẽ lại ảnh đó. Không có API COM nào
yêu cầu "vẽ lại ảnh" mà không cần mở thật công thức lên -- cách duy nhất
là Activate() để MathType mở, "làm bẩn" (dirty) rồi lưu lại.

CƠ CHẾ (đã kiểm chứng thực nghiệm bằng test_word_wizard_index_click.py,
xem file đó để có bản gốc đơn giản hơn, chạy độc lập ngoài GUI):
  1. Mở Word qua COM, Visible=False -- ẩn cửa sổ WORD. Cửa sổ MathType
     (1 ứng dụng OLE server RIÊNG, không phải 1 phần của Word) vẫn tự
     bung ra bình thường khi Activate() 1 công thức -- ẩn Word không ẩn
     được MathType.
  2. Với TỪNG công thức cần refresh: OLEFormat.Activate() -> đợi cửa sổ
     MathType xuất hiện (dò theo tiêu đề chứa "mathtype") -> gửi
     WM_COMMAND THẲNG tới hwnd đó qua PostMessage (KHÔNG dùng SendKeys --
     PostMessage không cần cửa sổ đang được focus/hiện, nên an toàn hơn
     hẳn: không chiếm bàn phím/chuột của máy, người dùng có thể làm việc
     khác trong lúc script chạy), theo ĐÚNG VỊ TRÍ (index) cố định trên
     menu -- KHÔNG so tên nút (tránh lệch ngôn ngữ hiển thị Anh/Việt của
     MathType):
       - Tab 0 (File), nút vị trí 3 = Update: đẩy dữ liệu MTEF hiện tại
         lên Word + buộc vẽ lại ảnh xem trước.
       - Tab 0 (File), nút vị trí 2 = Close and Return: đóng cửa sổ
         MathType, quay về Word.
  3. Đợi cửa sổ MathType đóng hẳn rồi mới sang công thức kế tiếp.
  4. Sau khi xong hết, doc.Save() rồi đóng Word.

KHỚP CÔNG THỨC CẦN REFRESH VỚI ĐÚNG InlineShape TRONG WORD: fix_mathtype_
parens.py chỉ biết công thức qua TÊN FILE (oleObjectN.bin), còn Word/COM
chỉ đưa ra được 1 danh sách InlineShape theo thứ tự xuất hiện, không có
tên file. Cầu nối ở đây là đọc THẲNG word/document.xml +
word/_rels/document.xml.rels (xem _equation_bin_names_in_document_order())
để có 1 danh sách tên file công thức THEO ĐÚNG THỨ TỰ xuất hiện trong tài
liệu -- thứ tự này PHẢI khớp 1-1 với thứ tự Word trả về trong
doc.InlineShapes khi lọc cùng tiêu chí ProgID (2 phía đọc CÙNG 1 file,
CÙNG bộ lọc). Nếu số lượng 2 bên lệch nhau (vd có công thức neo kiểu nổi
-- floating -- không nằm trong InlineShapes): DỪNG LẠI, không vẽ lại ảnh
nào cả, an toàn hơn đoán bừa sai công thức.

Yêu cầu: pip install pywin32. CHỈ chạy được trên Windows, máy phải có
Word + MathType (COM) cài sẵn. Nếu thiếu pywin32 hoặc COM lỗi: hàm ở đây
KHÔNG raise -- trả về RefreshReport ghi rõ lý do bỏ qua, vì lúc gọi tới
đây file .docx kết quả (với MTEF đã sửa đúng) đã được ghi ra XONG rồi ở
bước fix_mathtype_parens_in_docx() từ trước -- refresh chỉ là bước "làm
đẹp thêm" ảnh xem trước, thất bại ở đây không được phép làm mất file kết
quả đã có.
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
    # Cả 3 tên đều phải về None cùng nhau -- nếu chỉ win32com=None mà
    # win32con/win32gui vẫn giữ tên module cũ (hoặc chưa từng được gán),
    # lỡ có chỗ nào gọi nhầm 2 tên đó mà quên kiểm tra win32com is None
    # trước sẽ ăn NameError khó hiểu thay vì lỗi rõ ràng ngay từ đầu.
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

# Vị trí (index) CỐ ĐỊNH trên menu MathType -- xem test_word_wizard_index_click.py.
MENU_TAB_INDEX = 0
MENU_ITEM_UPDATE = 3
MENU_ITEM_CLOSE_RETURN = 2

# Hộp thoại "License information" (nhắc dùng thử/kích hoạt MathType) --
# xuất hiện KHÔNG cố định lần nào (theo chu kỳ nhắc riêng của MathType,
# quan sát thực tế: vài ngày mới lặp lại), và CHỈ CÓ THỂ bật lên ở lần
# Activate() ĐẦU TIÊN trong cả lượt refresh_equation_previews() -- xem
# docstring _wait_for_mathtype_window(). Chỉ có thể xuất hiện 1 lần nên
# chỉ cần dismiss ĐÚNG 1 LẦN, không cần đếm số lần lặp lại.
LICENSE_DIALOG_BUTTON_PREFIX = "continue trial"


def _equation_bin_names_in_document_order(docx_path: str) -> list[str]:
    """Danh sách tên file (vd "oleObject3.bin") của các công thức
    MathType/Equation Editor, THEO ĐÚNG THỨ TỰ xuất hiện trong
    word/document.xml (duyệt document order qua lxml .iter()) -- xem
    docstring module để biết vì sao thứ tự này cần khớp với
    _equation_shapes_in_order().

    Nhận diện 1 công thức qua thẻ <o:OLEObject Type="Embed" ProgID="..."
    r:id="..."/>: ProgID phải khớp EQUATION_PROGID_PREFIXES (bỏ qua OLE
    object khác, vd bảng Excel nhúng); r:id trỏ tới 1 <Relationship>
    trong word/_rels/document.xml.rels, lấy tên file cuối cùng của
    Target quan hệ đó (vd Target="embeddings/oleObject3.bin" ->
    "oleObject3.bin")."""
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
    """True nếu 1 InlineShape (lấy từ doc.InlineShapes qua COM) là công
    thức MathType/Equation Editor -- CÙNG tiêu chí lọc với phía XML
    (_equation_bin_names_in_document_order()), để 2 danh sách khớp thứ
    tự 1-1."""
    if inline_shape.Type != WD_INLINE_SHAPE_OLE:
        return False
    try:
        prog_id = inline_shape.OLEFormat.ProgID
    except Exception:
        return False
    return any(prog_id.startswith(p) for p in EQUATION_PROGID_PREFIXES)


def _equation_shapes_in_order(doc: Any) -> list[Any]:
    """Danh sách InlineShape công thức trong doc (COM Document đang mở),
    THEO ĐÚNG THỨ TỰ doc.InlineShapes trả về (= thứ tự xuất hiện trong
    tài liệu)."""
    return [s for s in doc.InlineShapes if _is_equation_shape(s)]


def _find_mathtype_window() -> int | None:
    """hwnd của cửa sổ đầu tiên có tiêu đề chứa MATHTYPE_TITLE_SUBSTRING
    (không phân biệt hoa/thường), hoặc None nếu chưa thấy."""
    found: list[int] = []

    def callback(hwnd: int, _extra: object) -> bool:
        title = win32gui.GetWindowText(hwnd)
        if MATHTYPE_TITLE_SUBSTRING in title.lower():
            found.append(hwnd)
        return True

    win32gui.EnumWindows(callback, found)
    return found[0] if found else None


def _find_dialog_with_button(button_text_prefix: str) -> tuple[int, int] | None:
    """Duyệt MỌI cửa sổ cấp cao nhất đang hiển thị, tìm cửa sổ ĐẦU TIÊN có
    1 control con (nút bấm) mà chữ trên nút bắt đầu bằng button_text_prefix
    (không phân biệt hoa/thường) -- nhận diện theo NÚT BẤM thay vì tiêu đề
    cửa sổ, vì tiêu đề có thể đổi theo bản dựng/ngôn ngữ, còn text nút
    "Continue trial" là đặc trưng ổn định của hộp thoại "License
    information" (nhắc dùng thử/kích hoạt) của MathType.

    Trả về (hwnd cửa sổ cha, hwnd nút bấm) của lần khớp đầu tiên, hoặc
    None nếu không thấy cửa sổ nào có nút như vậy."""
    match: list[tuple[int, int]] = []

    def enum_child(child_hwnd: int, parent_hwnd: int) -> bool:
        text = win32gui.GetWindowText(child_hwnd).strip().lower()
        if text.startswith(button_text_prefix):
            match.append((parent_hwnd, child_hwnd))
            return False
        return True

    def enum_top(hwnd: int, _extra: object) -> bool:
        if match or not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            win32gui.EnumChildWindows(hwnd, enum_child, hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(enum_top, None)
    return match[0] if match else None


def _dismiss_license_dialog_if_present() -> bool:
    """Nếu hộp thoại "License information" (nhắc dùng thử MathType chưa
    kích hoạt) đang mở, TỰ ĐỘNG bấm nút "Continue trial" của nó bằng
    PostMessage(BM_CLICK) (cùng cách làm với _click_menu_by_index() --
    không cần cửa sổ đang được focus/hiện), rồi trả về True. Trả về False
    nếu không thấy hộp thoại này (trường hợp bình thường, đa số các lần
    gọi)."""
    found = _find_dialog_with_button(LICENSE_DIALOG_BUTTON_PREFIX)
    if found is None:
        return False
    _dialog_hwnd, button_hwnd = found
    win32gui.PostMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
    return True


def _wait_for_mathtype_window(
    timeout: float = MATHTYPE_APPEAR_TIMEOUT_SECONDS,
    check_license_dialog: bool = False,
    log: Callable[[str], None] | None = None,
) -> int | None:
    """Chờ tới khi cửa sổ MathType xuất hiện, thăm dò mỗi POLL_SECONDS,
    tối đa timeout giây.

    check_license_dialog: CHỈ nên bật True cho lần activate ĐẦU TIÊN
    trong cả lượt refresh_equation_previews() -- hộp thoại "License
    information" chỉ có thể xuất hiện ở lần mở MathType đầu tiên sau 1
    chu kỳ nhắc (nhiều ngày mới lặp lại, theo quan sát thực tế); mọi
    công thức mở SAU trong CÙNG lượt chạy chắc chắn không còn gặp lại,
    nên bỏ qua hẳn bước dò/dismiss nó cho các công thức đó
    (EnumWindows/EnumChildWindows duyệt TOÀN BỘ cửa sổ trên máy, không
    cần trả giá đó ở mọi công thức chỉ để bắt 1 hộp thoại không thể xuất
    hiện lại).

    Khi check_license_dialog=True: ở MỖI lần thăm dò, kiểm tra THÊM hộp
    thoại đó, nhưng CHỈ bấm "Continue trial" ĐÚNG 1 LẦN cho cả lượt chờ
    này -- hộp thoại chỉ có thể xuất hiện 1 lần (xem hằng số
    LICENSE_DIALOG_BUTTON_PREFIX) nên không cần vòng lặp dò/dismiss/reset
    thời gian chờ nhiều lần như trước; dismiss xong vẫn tiếp tục đúng
    vòng chờ cửa sổ chính bình thường (không reset lại timeout).

    Trả về hwnd, hoặc None nếu hết giờ vẫn chưa thấy."""
    log = log or (lambda _msg: None)
    already_dismissed = False
    waited = 0.0
    while waited <= timeout:
        hwnd = _find_mathtype_window()
        if hwnd:
            return hwnd
        if check_license_dialog and not already_dismissed and _dismiss_license_dialog_if_present():
            already_dismissed = True
            log('    -> Gặp hộp thoại nhắc dùng thử MathType, đã tự bấm "Continue trial".')
        time.sleep(POLL_SECONDS)
        waited += POLL_SECONDS
    return None


def _wait_for_window_closed(hwnd: int, timeout: float = MATHTYPE_CLOSE_TIMEOUT_SECONDS) -> None:
    """Chờ tới khi hwnd không còn là cửa sổ hợp lệ nữa (đã đóng), tối đa
    timeout giây -- không raise nếu hết giờ mà vẫn còn (nơi gọi tự xử lý
    tiếp theo bình thường)."""
    waited = 0.0
    while win32gui.IsWindow(hwnd) and waited <= timeout:
        time.sleep(POLL_SECONDS)
        waited += POLL_SECONDS


def _click_menu_by_index(hwnd: int, tab_index: int, item_index: int) -> bool:
    """Gửi WM_COMMAND THẲNG tới hwnd theo ĐÚNG vị trí (tab_index,
    item_index) trên thanh menu -- KHÔNG so tên nút (an toàn hơn khi
    MathType hiển thị tiếng Anh/Việt khác nhau tuỳ máy). Dùng PostMessage
    (không phải SendKeys): không cần hwnd đang được focus/hiện, không
    chiếm bàn phím/chuột thật của máy.

    Trả về True nếu GỬI được lệnh (không đảm bảo MathType đã XỬ LÝ xong
    -- nơi gọi tự chờ thêm nếu cần, xem _refresh_one_shape())."""
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


def _refresh_one_shape(
    shape: Any, check_license_dialog: bool, log: Callable[[str], None]
) -> tuple[bool, str | None]:
    """Activate 1 InlineShape công thức, đợi cửa sổ MathType, bấm Update
    rồi Close and Return theo index cố định, đợi đóng hẳn. Trả về
    (True, None) nếu xong trọn vẹn, (False, lý_do) nếu phải bỏ qua --
    KHÔNG raise, để 1 công thức lỗi không chặn các công thức còn lại.

    check_license_dialog: truyền thẳng cho _wait_for_mathtype_window() --
    xem docstring hàm đó. Nếu Activate() ở đây tự lỗi (raise) TRƯỚC khi
    kịp gọi tới bước chờ, hộp thoại license coi như CHƯA từng được kiểm
    tra thật sự cho công thức này -- nơi gọi (refresh_equation_
    previews()) cố tình không xử lý riêng trường hợp hiếm này để giữ
    logic đơn giản, chấp nhận rủi ro rất nhỏ là hộp thoại có thể xuất
    hiện muộn hơn 1 công thức so với dự kiến."""
    try:
        shape.OLEFormat.Activate()
    except Exception as e:
        log(f"    -> Không kích hoạt (Activate) được công thức: {e}")
        return False, (
            "không mở được công thức trong MathType\n"
            f"      chi tiết: {e}"
        )

    log("    -> Đang chờ cửa sổ MathType mở ra...")
    hwnd = _wait_for_mathtype_window(check_license_dialog=check_license_dialog, log=log)
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
    """Kết quả 1 lượt refresh_equation_previews() trên 1 file .docx --
    cùng tinh thần với FixReport của fix_mathtype_parens.py."""

    def __init__(self, total_targets: int, refreshed: list[str], skipped: list[tuple[str, str]]) -> None:
        """total_targets: số công thức đáng lẽ cần vẽ lại ảnh (đầu vào).
        refreshed: list tên file đã vẽ lại ảnh xem trước thành công.
        skipped: list (tên_file, lý_do) chưa vẽ lại được -- KHÔNG phải
        lỗi ngăn cản kết quả chính: dữ liệu MTEF đã sửa đúng từ trước
        (ở fix_mathtype_parens_in_docx()) không bị ảnh hưởng gì, chỉ
        ảnh xem trước của công thức đó có thể còn hiện dạng ngoặc cũ cho
        tới khi người dùng tự mở lại bằng tay."""
        self.total_targets = total_targets
        self.refreshed = refreshed
        self.skipped = skipped

    def summary_lines(self) -> list[str]:
        """Danh sách dòng text theo nhãn 3 mức ([OK]/[BỎ QUA]) -- cùng
        định dạng với FixReport.summary_lines(), vì nội dung này chảy
        thẳng vào cùng ô Nhật ký qua CombinedReport.summary_lines()."""
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
    """Mở docx_path bằng Word (COM, CHẠY ẨN), buộc MathType vẽ lại ảnh
    xem trước cho ĐÚNG các công thức có tên file nằm trong
    equation_bin_names -- công thức khác giữ nguyên. SỬA TRỰC TIẾP lên
    docx_path (mở, Activate từng công thức, Save, đóng lại) -- gọi hàm
    này SAU KHI đã có file kết quả CUỐI CÙNG (xem pipeline.py), không
    phải file trung gian.

    equation_bin_names: tên file (vd "oleObject3.bin"), thường lấy từ
    FixReport.fixed của fix_mathtype_parens_in_docx() -- chỉ công thức
    THỰC SỰ đã đổi ngoặc mới cần vẽ lại ảnh.

    log(text): callback ghi 1 dòng log NGAY tại thời điểm xảy ra (không
    đợi xong hết mới trả về 1 cục) -- mặc định no-op nếu không truyền.

    Trả về RefreshReport. KHÔNG raise cho các lỗi có thể lường trước
    (thiếu pywin32, Word/MathType lỗi COM...) -- xem docstring module."""
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
        reason = (
            "không đọc được cấu trúc file để xác định thứ tự công thức\n"
            f"      chi tiết: {e}"
        )
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
        is_first_activation = True
        for i, name in enumerate(sorted(target_names), start=1):
            log(f"[{i}/{total}] Công thức {name}:")
            shape = name_to_shape.get(name)
            if shape is None:
                log("    -> Không tìm thấy công thức này khi mở bằng Word.")
                skipped.append((name, "không tìm thấy trong file khi mở bằng Word"))
                continue
            ok, reason = _refresh_one_shape(shape, is_first_activation, log)
            is_first_activation = False
            if ok:
                refreshed.append(name)
            else:
                skipped.append((name, reason))

        log("Đang lưu file...")
        doc.Save()
        log("Đã lưu.")

    except Exception as e:
        reason = (
            "gặp lỗi ngoài dự kiến khi điều khiển Word/MathType tự động\n"
            f"      chi tiết: {e}"
        )
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
