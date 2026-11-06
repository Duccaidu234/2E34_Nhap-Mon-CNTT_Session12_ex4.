# TÀI LIỆU QUẢN LÝ VÀ ĐẶC TẢ HỆ THỐNG BỆNH ÁN ĐIỆN TỬ (CLINIC-CLOUD)

---

## PHẦN 1: PHÂN TÍCH VÀ ĐẶC TẢ SRS (MODULE BỆNH ÁN ĐIỆN TỬ)

### 1.1 Khái niệm & Phân biệt Yêu cầu Chức năng và Phi chức năng
*   **Yêu cầu Chức năng (Functional Requirements - FR):** Định nghĩa các hành vi, tính năng cụ thể mà hệ thống phải thực hiện khi người dùng thao tác. *(Ví dụ: Khám bệnh, kê đơn, truy xuất lịch sử bệnh án).*
*   **Yêu cầu Phi chức năng (Non-Functional Requirements - NFR):** Định nghĩa các tiêu chuẩn về vận hành, chất lượng, hiệu năng và đặc biệt là tính bảo mật dữ liệu y tế tuân thủ theo các quy định an toàn thông tin.

---

### 1.2 Yêu cầu Chức năng (Functional Requirements - FR)

#### FR-01: Quản lý Hồ sơ & Lịch sử Khám bệnh
*   **Mô tả:** Cho phép Bác sĩ tạo mới, cập nhật và xem chi tiết bệnh án của bệnh nhân.
*   **Đầu vào:** Mã bệnh nhân, Triệu chứng lâm sàng, Chẩn đoán, Kết quả xét nghiệm.
*   **Đầu ra:** Bệnh án điện tử hoàn chỉnh được lưu trữ trên hệ thống Clinic-Cloud.

#### FR-02: Kê đơn thuốc Điện tử (e-Prescribing)
*   **Mô tả:** Bác sĩ chọn thuốc từ danh mục chuẩn y tế, nhập liều dùng và xuất đơn thuốc điện tử có chữ ký số.
*   **Ràng buộc:** Tự động cảnh báo dị ứng thuốc dựa trên tiền sử bệnh nhân.

#### FR-03: Tra cứu & Tìm kiếm Hồ sơ
*   **Mô tả:** Cho phép Nhân viên y tế tìm kiếm hồ sơ bệnh nhân nhanh chóng theo Mã BN, Số CCCD/CMND hoặc Số điện thoại.

---

### 1.3 Yêu cầu Phi chức năng - Bảo mật & An toàn dữ liệu (NFR Security)

#### NFR-SEC-01: Mã hóa Dữ liệu (Encryption Standards)
*   Tất cả dữ liệu hồ sơ bệnh nhân ở trạng thái lưu trữ (Data-at-Rest) phải được mã hóa theo chuẩn **AES-256**.
*   Dữ liệu truyền tải trên mạng (Data-in-Transit) phải mã hóa bằng **TLS 1.3 / HTTPS**.

#### NFR-SEC-02: Kiểm soát Truy cập theo Vai trò (RBAC - Role-Based Access Control)
*   Chỉ Bác sĩ được phân công mới có quyền xem toàn bộ nội dung bệnh án.
*   Lễ tân chỉ xem được thông tin hành chính (Họ tên, SĐT, Ngày sinh), không truy cập được lịch sử khám bệnh.

#### NFR-SEC-03: Ghi vết Hệ thống (Audit Logging)
*   Hệ thống phải tự động ghi nhật ký (Audit Log) đối với mọi thao tác Xem (Read), Thêm (Create), Sửa (Update), Xóa (Delete) hồ sơ bệnh nhân. Nhật ký bao gồm: User ID, Thời gian (Timestamp), IP Address và Nội dung thao tác.

---

## PHẦN 2: TRIỂN KHAI DỮ LIỆU EXCEL CHUẨN DATABASE MINDSET

### 2.1 Giải pháp xử lý "Bẫy dữ liệu" (Xử lý hàng nghìn ô trống trong AutoFilter)

#### **Vấn đề (Bẫy dữ liệu):**
Khi bảng dữ liệu 1.000 hồ sơ có nhiều dòng hoặc ô bị trống thông tin, nếu chọn 1 ô rồi nhấn `Ctrl + Shift + L`, Excel sẽ dừng vùng lọc (AutoFilter) ngay tại dòng trống đầu tiên. Điều này khiến dữ liệu phía dưới dòng trống bị bỏ sót khỏi bộ lọc.

#### **Quy trình xử lý chuẩn Database Mindset:**
1.  **Bước 1: Chọn toàn bộ vùng dữ liệu chuẩn xác**
    *   Không chọn 1 ô duy nhất.
    *   Nhấp chọn **toàn bộ các cột** chứa dữ liệu (ví dụ: bôi đen từ Cột `A` đến Cột `H`), hoặc bấm `Ctrl + A` sau khi đã chọn vùng dữ liệu mở rộng đảm bảo bao phủ đến dòng cuối cùng.
2.  **Bước 2: Kích hoạt AutoFilter chuẩn**
    *   Nhấn tổ hợp phím **`Ctrl + Shift + L`** (hoặc truy cập `Data` $\rightarrow$ `Filter`).
    *   Lúc này thanh AutoFilter sẽ được áp dụng cho toàn bộ 1.000 dòng dữ liệu bất kể có dòng trống.
3.  **Bước 3: Chuẩn hóa / Điền dữ liệu trống (Optional - Best Practice)**
    *   Bấm `Ctrl + G` $\rightarrow$ Chọn `Special...` $\rightarrow$ Tích chọn **Blanks** $\rightarrow$ Click `OK`.
    *   Nhập `N/A` hoặc `Chưa cập nhật`, sau đó nhấn **`Ctrl + Enter`** để điền đồng loạt vào tất cả các ô trống.

---

### 2.2 Cấu trúc Bảng dữ liệu Hồ sơ Bệnh nhân (1000 records sample)

| Mã Bệnh Nhân | Họ và Tên | Năm Sinh | Giới Tính | Khoa Khám | Chẩn Đoán | Trạng Thái Điều Trị | Ghi Chú |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| BN-2026-001 | Nguyễn Văn A | 1985 | Nam | Nội tổng hợp | Viêm dạ dày | Đang điều trị | *Chưa cập nhật* |
| BN-2026-002 | Trần Thị B | 1992 | Nữ | Nhi khoa | Sốt siêu vi | Đã xuất viện |  |
| BN-2026-003 | Lê Hoàng C | 1978 | Nam | Tim mạch | Tăng huyết áp | Theo dõi định kỳ | Tiền sử dị ứng |
| BN-2026-004 | Phạm Minh D | 2001 | Nam | Tai Mũi Họng | Viêm Amydal | Đã xuất viện |  |
| BN-2026-005 | Hoàng Thu E | 1995 | Nữ | Da liễu | Viêm da cơ địa | Đang điều trị | *Chưa cập nhật* |

---

### 2.3 Visualizing Dữ liệu bằng Donut Chart Tối giản (PowerPoint / Excel)

#### **Nguyên tắc thiết kế Chart tối giản (Minimalist Donut Chart):**
*   **Mục tiêu:** Trình bày tỷ lệ phân bố các ca bệnh dịch/khoa khám một cách trực quan, sạch sẽ.
*   **Loại biểu đồ:** Doughnut Chart.
*   **Quy chuẩn thiết kế:**
    *   Đường kính lỗ giữa (Doughnut Hole Size): Đặt ở mức $65\% - 75\%$.
    *   Loại bỏ các thành phần rác: Xóa Gridlines (đường lưới), xóa Legend (chú giải bên dưới) và thêm Data Labels trực tiếp trên/cạnh các lát cắt.
    *   Ở tâm biểu đồ (Center Label): Đặt 1 Textbox thể hiện tổng số ca bệnh (VD: `1,000 Ca`).

#### **Cấu trúc Màu sắc:**
*   Lát chiếm tỷ lệ lớn nhất/Cần lưu ý: Accent Color (Đỏ nhung hoặc Cam đậm).
*   Các lát còn lại: Các tông màu xám trung tính (Light Gray, Medium Gray).

---

## PHẦN 3: THAO TÁC SỬA LỖI ĐỊNH DẠNG TRÊN MICROSOFT WORD (KIỂM THỬ)

### Lỗi giả định:
Trong quá trình gán nhãn Heading Styles để xây dựng mục lục cho tài liệu SRS, người dùng vô tình chọn nhầm một đoạn văn bản thường (Body text) thành **Heading 1 / Heading 2**, khiến đoạn văn đó hiển thị trên Navigation Pane và làm hỏng cấu trúc mục lục.

### Quy trình khắc phục bằng tổ hợp phím `Ctrl + Shift + N`:

1.  **Bước 1: Xác định vị trí lỗi**
    *   Mở thanh **Navigation Pane** (`Ctrl + F` $\rightarrow$ Chọn tab *Headings*).
    *   Nhấp chuột vào tiêu đề bị lỗi (đoạn văn bị gán nhầm thành Heading). Con trỏ chuột sẽ tự động nhảy đến đúng đoạn văn đó trong văn bản.
2.  **Bước 2: Xóa bỏ Style Heading và trả về định dạng chuẩn**
    *   Bôi đen toàn bộ đoạn văn bản bị gán nhầm (hoặc đặt con trỏ chuột vào trong đoạn văn đó).
    *   Nhấn tổ hợp phím **`Ctrl + Shift + N`**.
    *   *Tác dụng:* Phím tắt `Ctrl + Shift + N` sẽ ngay lập tức chuyển style của đoạn văn bản hiện tại về kiểu **Normal** (văn bản thường), loại bỏ cấp độ Heading.
3.  **Bước 3: Cập nhật lại Mục tự động (Update Table of Contents)**
    *   Click chuột phải vào bảng Mục lục ở đầu trang.
    *   Chọn **Update Field** $\rightarrow$ Chọn **Update entire table** $\rightarrow$ Nhấn **OK**.
    *   Kiểm tra lại Navigation Pane để đảm bảo cấu trúc văn bản đã chuẩn xác.

---

## BẢNG KIỂM TRA ĐÁP ỨNG RÀNG BUỘC (CHECKLIST)

* [x] **Phần 1**: Đã soạn thảo SRS đầy đủ các FR và NFR về bảo mật cho module Bệnh án điện tử.
* [x] **Phần 2**: Đã trình bày giải pháp xử lý bẫy ô trống bằng `Ctrl + Shift + L` trên vùng bôi đen đầy đủ theo Database Mindset.
* [x] **Phần 2**: Đã hướng dẫn thiết kế Donut Chart tối giản để trình bày số liệu dịch bệnh.
* [x] **Phần 3**: Đã hướng dẫn quy trình sửa lỗi gán nhầm Heading bằng tổ hợp phím `Ctrl + Shift + N` và `Update entire table`.