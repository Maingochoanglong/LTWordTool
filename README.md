# LTWordTool

🇻🇳 Tiếng Việt | [🇬🇧 English](README.en.md)

![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

**Công cụ desktop xử lý file Word (.docx) chứa công thức MathType** — thay thế nội dung giữ nguyên định dạng, và tự động sửa lỗi ngoặc tự co giãn trong công thức MathType.

Dành cho nhân viên văn phòng, giáo viên, biên tập viên xử lý tài liệu tiếng Việt có công thức toán.

---

## Mục lục

- [Tính năng](#tính-năng)
- [Yêu cầu hệ thống](#yêu-cầu-hệ-thống)
- [Cài đặt](#cài-đặt)
- [Sử dụng](#sử-dụng)
- [Cấu trúc dự án](#cấu-trúc-dự-án)
- [Giấy phép](#giấy-phép)

---

## Tính năng

- **Thay thế nội dung theo định dạng** — tìm & thay 1 hay nhiều cặp nội dung trong file `.docx`, giữ nguyên định dạng (đậm/nghiêng/màu/font...), lan định dạng tự động về 2 phía theo chuỗi dừng tùy chọn.
- **Sửa ngoặc MathType** — tự động chuyển ngoặc tự co giãn quanh văn bản/số đơn giản trong công thức MathType thành ngoặc cứng, xử lý đúng cả trường hợp ngoặc lồng nhau nhiều lớp.
- **Tự động vẽ lại ảnh xem trước MathType** (yêu cầu Windows + Word + MathType) — sau khi sửa ngoặc, tool tự mở Word (chạy ẩn) để MathType vẽ lại đúng ảnh xem trước, không cần double-click từng công thức.
- Giao diện đơn giản, không cần dòng lệnh — chọn file, tích chọn thao tác, bấm chạy.

## Yêu cầu hệ thống

| Thành phần   | Bắt buộc | Ghi chú                                                                 |
|--------------|:--------:|--------------------------------------------------------------------------|
| Python 3.10+ | Có       | Chỉ cần nếu chạy từ mã nguồn                                             |
| PySide6      | Có       | Giao diện                                                                 |
| lxml         | Có       | Đọc/ghi XML trong `.docx`                                                |
| olefile      | Có       | Đọc cấu trúc OLE/CFB                                                     |
| pywin32      | Không    | Chỉ cần cho bước vẽ lại ảnh xem trước (Windows + Word + MathType cài sẵn) |
| Windows      | Không    | Bắt buộc riêng cho bước vẽ lại ảnh xem trước; các thao tác khác chạy được trên mọi hệ điều hành |

## Cài đặt

### Cách 1 — Tải bản dựng sẵn (khuyên dùng cho người dùng phổ thông)

Tải file `.exe` mới nhất tại [Releases](https://github.com/Maingochoanglong/LTWordTool/releases/latest) — không cần cài Python.

### Cách 2 — Chạy từ mã nguồn

```bash
git clone https://github.com/Maingochoanglong/LTWordTool.git
cd LTWordTool
pip install -r requirements.txt
python gui.py
```

Trên Windows, cài thêm `pywin32` nếu muốn dùng tính năng vẽ lại ảnh xem trước:

```bash
pip install pywin32
```

## Sử dụng

1. Mở app, chọn **FILE NGUỒN** (`.docx` cần xử lý).
2. Tích chọn 1 hoặc cả 2 thao tác:
   - **Thay thế nội dung** — thêm các cặp (file tìm, file thay) vào danh sách.
   - **Sửa ngoặc MathType** — không cần cấu hình thêm.
3. Chọn nơi lưu **FILE KẾT QUẢ** (hoặc để trống để tool tự đặt tên).
4. Bấm **Thực Hiện**, theo dõi tiến độ trong ô nhật ký.

Nếu tích cả 2 thao tác, thứ tự chạy luôn cố định: thay thế nội dung trước, sửa ngoặc MathType sau.

## Cấu trúc dự án

| File                              | Vai trò                                                        |
|------------------------------------|-----------------------------------------------------------------|
| `gui.py`                          | Giao diện Qt6 (PySide6), lắp ráp các thao tác                   |
| `panel.py`                        | Khung UI dùng chung (luồng nền, nút bấm, nhật ký, nhớ cấu hình)  |
| `pipeline.py`                     | Nối các bước xử lý chạy tuần tự                                 |
| `replace_docx.py`                 | Thay thế nội dung giữ định dạng trong `.docx`                   |
| `fix_mathtype_parens.py`          | Orchestration sửa ngoặc MathType                                 |
| `mtef_parser.py` / `mtef_transform.py` | Đọc & biến đổi cấu trúc nhị phân MTEF                       |
| `cfb_builder.py`                  | Đóng gói lại container OLE/CFB                                   |
| `mathtype_refresh.py`             | Tự động vẽ lại ảnh xem trước qua Word COM                        |

## Giấy phép

Mã nguồn phát hành theo **[AGPL-3.0](LICENSE)** — dùng, sửa, tự chạy hoàn toàn miễn phí.

Muốn dùng theo hướng khác (thương mại, nhúng vào sản phẩm, phân phối lại không theo AGPL-3.0)? Liên hệ: **<email-của-bạn>**

## Báo lỗi / Đóng góp

Gặp lỗi, muốn góp ý, hoặc có câu hỏi khi sử dụng?

1. **Kiểm tra trước** trong [Issues đã có](https://github.com/Maingochoanglong/LTWordTool/issues?q=is%3Aissue) — có thể lỗi đã được báo cáo hoặc trả lời rồi.
2. Chưa thấy → **tạo [Issue mới](https://github.com/Maingochoanglong/LTWordTool/issues/new)**, mô tả rõ: đang làm gì, mong đợi gì, thực tế xảy ra gì (kèm file `.docx` mẫu nếu tiện, đã xoá nội dung nhạy cảm).

Vui lòng dùng Issue thay vì nhắn tin riêng — giúp người dùng khác gặp vấn đề tương tự cũng tìm thấy câu trả lời, và không phụ thuộc vào 1 kênh liên hệ duy nhất.

Pull request cải tiến luôn được hoan nghênh.
