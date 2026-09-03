# WFM Pipeline Dashboard — Setup Guide

## Yêu cầu
- Python environment hiện tại đã có đủ thư viện (xem requirements_clean.txt trong repo)
- Cần thêm: `pip install streamlit nbconvert --break-system-packages`

## Cài đặt nhanh (1 lần)

```bash
# 1. Copy app.py vào thư mục CODE/Python_Code/dashboard/
copy app.py "C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\dashboard\app.py"

# 2. Cài streamlit nếu chưa có
pip install streamlit nbconvert --break-system-packages
```

## Chạy dashboard

```bash
cd "C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\dashboard"
streamlit run app.py
```

Hoặc tạo shortcut `.bat`:
```bat
@echo off
cd /d "C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\dashboard"
streamlit run app.py
pause
```

Dashboard sẽ mở tại: http://localhost:8501

---

## Cấu trúc file được tạo tự động

```
dashboard/
├── app.py                  ← File này
├── pipeline.db             ← SQLite: task history + email queue
├── logs/                   ← Log files từ mỗi lần chạy
│   ├── hc_master_run_20260901_0812.log
│   ├── email_atd_preview_20260901_0900.log
│   └── email_atd_preview_20260901_0900_executed.ipynb
└── email_previews/         ← HTML preview của email
    ├── email_atd_preview.html
    ├── email_performance_preview.html
    └── ...
```

---

## Flow sử dụng hàng ngày

### Buổi sáng (ETL)
1. Sidebar → **▶️ Run all ETL** → chờ Layer 1 xong
2. Sidebar → **🤖 Run all Bots** → Teams alerts tự động

### Trước khi gửi email
1. Sidebar → **👁 Preview all Emails**
2. Tab **Email Queue** → review từng email
3. Nhấn **✅ Approve** cho email đã kiểm tra
4. Nhấn **📨 Send** để gửi qua Outlook

### Chạy từng task
- Tab **Pipeline** → từng task card có nút **▶️ Run** / **👁 Preview** / **📋 Queue**

---

## Notes kỹ thuật

### Email Preview hoạt động như thế nào?
1. Dashboard clone notebook sang file tạm
2. Patch `SEND_EMAIL = True` → `SEND_EMAIL = False` trong clone
3. Chạy notebook bằng `jupyter nbconvert --execute`
4. Extract HTML output từ `display(HTML(...))` calls trong notebook
5. Lưu vào `email_previews/{task_id}_preview.html`
6. Hiện trong iframe trong tab Email Queue

### Khi nào nên dùng Send trực tiếp (không Preview)?
- Bot real-time (Layer 2): luôn send trực tiếp → Teams webhook
- ATD Realtime email: time-sensitive, có thể skip preview

### Troubleshooting
- **Task failed**: xem log trong tab **📋 Logs**
- **Preview HTML trống**: notebook không có `display(HTML(...))` output
- **Send không work**: đảm bảo Outlook đang mở trên máy (win32com requirement)
