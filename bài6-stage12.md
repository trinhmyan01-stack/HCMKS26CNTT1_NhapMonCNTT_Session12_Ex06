
## PHẦN 1: PHÂN TÍCH ĐẶC ĐIỂM BẢNG EXCEL "XẤU" VS BẢNG "ĐÚNG" (DATABASE MINDSET)

Trong quản trị và phân tích dữ liệu chuyên nghiệp, việc tuân thủ các quy tắc thiết kế cơ sở dữ liệu trên bảng tính quyết định trực tiếp đến hiệu suất xử lý và độ chính xác của báo cáo:

1. **Bảng Excel "Xấu" (Vi phạm Quy tắc Vàng):**
   - **Có Merge Cell (Trộn ô) ở hàng tiêu đề hoặc dữ liệu:** Làm mất cấu trúc phẳng của bảng, khiến các hàm tìm kiếm (`VLOOKUP`, `XLOOKUP`) và bộ lọc (`AutoFilter`) không thể nhận diện đúng ranh giới cột/hàng.
   - **Tồn tại hàng trống hoặc cột trống ngắt quãng:** Khiến Excel hiểu lầm rằng bảng dữ liệu đã kết thúc tại ô trống đó, làm đứt gãy vùng quét của PivotTable và AutoFilter.
   - **Dữ liệu lẫn lộn định dạng (Ngày tháng, số, chữ viết tay không đồng nhất):** Gây lỗi tính toán, không thể sắp xếp (Sort) hay tổng hợp số liệu tự động.
   - **Thiếu tính nguyên tử của dữ liệu (Atomicity):** Gộp cả họ tên, địa chỉ, số điện thoại vào chung một ô thay vì tách thành các cột riêng biệt (Field/Record).

2. **Bảng Excel "Đúng" (Chuẩn Database Mindset):**
   - **Cấu trúc dạng Bảng phẳng (Flat Table):** Mỗi cột đại diện cho một thuộc tính duy nhất (**Field**), mỗi hàng đại diện cho một bản ghi độc lập (**Record**).
   - **Hàng tiêu đề duy nhất:** Nằm ở dòng đầu tiên, không có ô gộp (No Merge Cell), tên cột rõ ràng không chứa ký tự đặc biệt.
   - **Đồng nhất kiểu dữ liệu:** Cột số ra số, cột ngày tháng chuẩn định dạng ngày tháng, giúp hàm tính toán và công cụ phân tích hoạt động hoàn hảo tuyệt đối.

---

## PHẦN 2: TRIỂN KHAI HỌC LIỆU DỰ ÁN SMART-HOME

### 1. Tài liệu SRS Mô tả Nền tảng Đấu giá Nhà đất Trực tuyến (Cấu trúc Heading 1, 2, 3)

# 1. Tổng quan Dự án Smart-Home & Đấu giá Nhà đất Trực tuyến
## 1.1. Mục đích hệ thống
Xây dựng nền tảng thương mại điện tử kết hợp đấu giá tài sản bất động sản trực tuyến, đảm bảo tính minh bạch, bảo mật và tốc độ xử lý giao dịch thời gian thực cho người mua và nhà đầu tư.
## 1.2. Phạm vi ứng dụng
Hỗ trợ đăng tin tài sản, kiểm duyệt pháp lý tự động, tổ chức phiên đấu giá trực tuyến theo thời gian thực và quản lý hợp đồng điện tử.

# 2. Yêu cầu Chức năng (Functional Requirements)
## 2.1. Module Đăng ký & Xác thực Người dùng
- Phân quyền tài khoản: Người mua (Buyer), Người bán/Môi giới (Broker), Quản trị viên (Admin).
- Xác thực danh tính sinh trắc học hoặc CCCD gắn chip liên kết VNeID.
## 2.2. Module Đấu giá Trực tuyến (Live Bidding)
- Cho phép người dùng đặt giá thầu (Place Bid) theo thời gian thực với độ trễ dưới 1 giây qua giao thức WebSocket.
- Hệ thống tự động đếm ngược thời gian kết thúc phiên đấu giá và chốt người trả giá cao nhất hợp lệ.

# 3. Yêu cầu Phi chức năng (Non-functional Requirements)
## 3.1. Bảo mật & Hiệu năng
- Mã hóa toàn bộ dữ liệu giao dịch và thông tin định danh cá nhân bằng chuẩn AES-256.
- Khả năng chịu tải (Scalability): Hỗ trợ tối thiểu 10,000 phiên đấu giá diễn ra đồng thời mà không bị gián đoạn.

---

### 2. Triển khai Bảng Quản lý Danh sách Nhà đất & Tra cứu Hoa hồng bằng VLOOKUP
Để tra cứu chính xác tỷ lệ hoa hồng môi giới dựa trên phân khúc giá trị nhà đất mà không bị lệch dữ liệu, ta thiết lập bảng và công thức chuẩn:

| Mã BĐS | Tên Tài Sản | Giá trị định giá (VNĐ) | Phân khúc | Tỷ lệ hoa hồng | Tiền hoa hồng thực tế |
| :--- | :--- | :---: | :---: | :---: | :--- |
| **BDS01** | Căn hộ Vinhomes Grand Park | 3,500,000,000 | Phổ thông | `=VLOOKUP(D2, $H$2:$I$5, 2, FALSE)` | `=C2*E2` |
| **BDS02** | Biệt thự Thảo Điền Quận 2 | 25,000,000,000 | Cao cấp | `=VLOOKUP(D2, $H$2:$I$5, 2, FALSE)` | `=C2*E2` |

* **Bảng tham chiếu hoa hồng ($H$2:$I$5):**
  - Cột H (Phân khúc): Phổ thông, Trung cấp, Cao cấp, Siêu sang.
  - Cột I (Tỷ lệ hoa hồng tương ứng): 1.0%, 1.5%, 2.0%, 2.5%.

---

### 3. Thiết kế Slide Pitch Deck Gọi vốn (Tuân thủ Quy tắc 10/20/30 & White Space)
Bộ slide gọi vốn cho dự án **Smart-Home** được thiết kế tối giản, không vượt quá 10 slide, tuân thủ nguyên tắc **1 Slide - 1 Thông điệp** và duy trì **White Space** (khoảng trắng tối thiểu 40% diện tích slide) để làm nổi bật trọng tâm:

- **Slide 1: Trang bìa**
  - *Tiêu đề:* Smart-Home: Kỷ nguyên Đấu giá Bất động sản Minh bạch.
  - *Thông điệp:* Nền tảng PropTech toàn diện cho thế hệ nhà đầu tư số.
- **Slide 2: Vấn đề thị trường**
  - *Thông điệp:* Thị trường bất động sản truyền thống thiếu minh bạch về giá và thủ tục pháp lý phức tạp.
- **Slide 3: Giải pháp Smart-Home**
  - *Thông điệp:* Minh bạch hóa toàn bộ quy trình định giá và đấu giá trực tuyến qua công nghệ Blockchain & IoT.
- **Slide 4: Mô hình Kinh doanh (Business Model)**
  - *Thông điệp:* Thu phí dịch vụ trung gian (Commission) và gói hội viên VIP phân tích dữ liệu thị trường.
- **Slide 5: Kêu gọi Vốn (Call to Action)**
  - *Thông điệp:* Kêu gọi 500.000 USD cho giai đoạn mở rộng thị phần toàn quốc.

---

## PHẦN 3: KIỂM THỬ VÀ KHẮC PHỤC LỖI TRÔI VÙNG DỮ LIỆU VỚI DẤU `$` (DOLLAR SIGN)

* **Vấn đề giả định:** Khi viết công thức VLOOKUP tra cứu hoa hồng `=VLOOKUP(D2, H2:I5, 2, FALSE)` ở dòng 2 và kéo (fill down) xuống các dòng phía dưới, kết quả trả về lỗi `#N/A` ở các dòng tiếp theo.
* **Nguyên nhân:** Khi kéo công thức xuống, vùng tham chiếu bảng phụ `H2:I5` không được cố định, nó sẽ tự động dịch chuyển thành `H3:I6`, `H4:I7`... khiến ô cần tra cứu rơi ra ngoài vùng dữ liệu tham chiếu (hiện tượng trôi vùng dữ liệu).
* **Cách khắc phục bằng Địa chỉ tuyệt đối (Dollar Sign - `$`)**:
  - Thêm ký tự dấu đô la (`$`) vào trước tên cột và số hàng của vùng tham chiếu để biến nó thành **Địa chỉ tuyệt đối**.
  - Công thức chuẩn được sửa lại là:

    =VLOOKUP(D2, $H$2:$I$5, 2, FALSE)
  - **Tác dụng:** Dấu `$` khóa chặt tọa độ của bảng tham chiếu từ hàng 2 đến hàng 5 của cột H và I. Khi kéo công thức xuống hàng trăm dòng bên dưới, ô `D2` tự động đổi thành `D3`, `D4`... nhưng vùng bảng phụ `$H$2:$I$5` được giữ nguyên cố định tuyệt đối, giúp việc tra cứu hoa hồng diễn ra chính xác 100% mà không bao giờ bị lỗi trôi vùng.