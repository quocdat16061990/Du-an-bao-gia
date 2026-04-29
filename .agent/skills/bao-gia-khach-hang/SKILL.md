---
name: bao-gia-khach-hang
description: Tự động hóa tạo báo giá phần mềm chuyên nghiệp từ Excel Master, bảo toàn định dạng và công thức.
---

# Skill: Tự động hóa Báo giá Khách hàng (Pro Version)

Bộ skill này vận hành một hệ thống Quotation Engine (Python) cao cấp, giúp tạo báo giá từ file định mức mà không làm hỏng cấu trúc phức tạp của template Excel (merged cells, formulas, styles).

---

## 1. Yêu Cầu Cài Đặt (Setup)

### Bước 1: Môi trường Python
AI phải sử dụng môi trường ảo (`venv`) tại `f:\Bao_Gia`:
```powershell
.\venv\Scripts\Activate.ps1
pip install openpyxl
```

### Bước 2: Cấu trúc File
- `Bao_Gia_Mau.xlsx`: Template báo giá (Chứa logo, footer, merged cells).
- `Dinh_Muc_Phan_Mem_Full.xlsx`: Cơ sở dữ liệu 100+ tính năng và đơn giá.
- `.agent\skills\bao-gia-khach-hang\script\Bao-Gia.py`: Script gốc (không đổi).

> [!IMPORTANT]
> **Lưu ý về đường dẫn:** Tất cả các tệp phải sử dụng **đường dẫn tương đối** (Relative Paths) để đảm bảo script hoạt động ổn định khi di chuyển thư mục dự án sang máy khác.

---

## 2. Quy Tắc Vận Hành "Safe-Workflow" (Execution Rule)

Khi có yêu cầu tạo báo giá mới, AI **BẮT BUỘC** thực hiện theo quy trình 3 bước sau:

### Bước 1: Khởi tạo (Sync)
Copy file script mẫu ra thư mục gốc để làm việc:
```powershell
cp -Force .agent\skills\bao-gia-khach-hang\script\Bao-Gia.py Bao-Gia.py
```

### Bước 2: Cập nhật dữ liệu (Update)
Cập nhật thông tin vào file `Bao-Gia.py` tại thư mục gốc:
- `PARTNER`: Thông tin khách hàng.
- `REQUESTED_FEATURES`: Danh sách tính năng (hỗ trợ tìm kiếm gần đúng).

### Bước 3: Thực thi & Kiểm tra (Run)
Chạy script bằng **venv** tại thư mục gốc:
```powershell
.\venv\Scripts\python.exe Bao-Gia.py
```

---

## 3. Đặc tính Kỹ thuật (Technical Features)

AI cần biết các khả năng của Engine này để tư vấn cho người dùng:
1.  **Safe Insertion:** Tự động bảo toàn cấu trúc `Merged Cells` ở Footer khi chèn thêm hàng (không bị xô lệch bảng).
2.  **Auto-Detect:** Tự động tìm dòng "STT" và dòng "Tổng cộng" trong template (linh hoạt khi thay đổi mẫu Excel).
3.  **Dynamic Formulas:** 
    - Tự động điền công thức `=E*F` cho từng dòng sản phẩm.
    - Tự động cập nhật công thức `=SUM(G12:G...)` ở ô Tổng cộng để bao phủ toàn bộ dữ liệu mới.
4.  **Style Inheritance:** Tự động sao chép toàn bộ Border, Font, Fill từ dòng mẫu xuống các dòng mới tạo.

---

> [!IMPORTANT]
> **Nguyên tắc vàng:** Không bao giờ sửa trực tiếp dữ liệu khách hàng vào file trong thư mục `.agent`. Luôn copy ra ngoài thư mục gốc (`Bao-Gia.py`) trước khi thao tác.

> [!TIP]
> **Lệnh chạy nhanh**: `.\venv\Scripts\python.exe Bao-Gia.py`
