# Text Converter Studio

App desktop chuyên convert văn bản với nhiều kiểu chuyển đổi, tập trung nhóm định dạng tài liệu.

## Tính năng

- Convert nhiều định dạng văn bản: TXT, MD, HTML, DOCX, PDF, CSV, JSON
- Hỗ trợ **2 chiều PDF ↔ DOCX**
- Chọn **nhiều file nguồn cùng lúc** để convert hàng loạt
- Hỗ trợ các cặp phổ biến:
  - TXT ↔ MD
  - MD ↔ HTML
  - TXT ↔ HTML
  - TXT ↔ DOCX
  - DOCX ↔ MD
  - TXT ↔ PDF
  - PDF → TXT
  - HTML → PDF
  - CSV ↔ JSON
  - CSV → TXT
  - JSON ↔ TXT
- Tự tạo thư mục làm việc:
  - `converted_file/workspace/input`
  - `converted_file/workspace/output`

Khi chạy EXE, app vẫn lưu trong thư mục `converted_file/workspace` cạnh file chạy, không lưu ở thư mục temp.

## Cài đặt

Từ thư mục gốc `Random_Essential`, chạy:

```bash
cd Convert_File
pip install -r requirements.txt
```

## Chạy app

```bash
cd Convert_File
python converter_app.py
```

## Ghi chú quan trọng

- `DOCX -> PDF` ưu tiên dùng `docx2pdf` (chất lượng tốt hơn), nhưng phụ thuộc môi trường Word trên Windows.
- Nếu `docx2pdf` không dùng được, app sẽ fallback sang xuất PDF dạng text để vẫn tạo được file.
- `PDF -> DOCX` dùng `pdf2docx` nên độ chính xác phụ thuộc layout PDF nguồn.
