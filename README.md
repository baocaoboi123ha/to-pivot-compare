# TO Pivot Compare (web riêng)

Ứng dụng Streamlit độc lập: pivot sheet **Topos** và **BC**, so sánh **Số lượng** và **Unit**, xuất file Excel mới (không ghi đè file gốc).

## Có cần cài Excel không?

**Không.** Chỉ cần trình duyệt; server xử lý `.xlsx` bằng **openpyxl**.

## Chạy local

```bash
cd to-pivot-compare-web
pip install -r requirements.txt
streamlit run app.py
```

## File kết quả

- Giữ nguyên mọi sheet trong file upload.
- Thêm sheet `Ket qua lech ma SP`: **chỉ các dòng pivot lệch** (không có dòng Khop).
- Cột: TransferFrom/To, Date, Ma SP, Topos SL/Unit, BC SL/Unit, Chenh lech SL, Status, Ma bill Topos lech.

Hậu tố file: ` - PivotCompare v5.xlsx`.

## Cấu trúc file nguồn

Một file `.xlsx` (mẫu `TO 2009.xlsx`): sheet **Topos**, **BC**, và sheet phụ khác.

BC `ItemNo` map sang Ma SP qua `Item Reference No.` (nếu có) hoặc tên sản phẩm trùng Topos.
