# CNTT1-nh-p-m-n-c-ng-ngh-th-ng-tin-Session12_Ex04
# Phần 1.Phân tích : mapping 4 tham số của VLOOKUP vào mô hình "Nhân viên kho tra hồ sơ"
 ## cú pháp chuẩn của hàm VLOOKUP : VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup]
| Tham số VLOOKUP | Tên gọi kỹ thuật | Tương đương trong mô hinh "Nhân Viên Kho | Ý nghĩa thực tế FastMove |
| --- | --- | --- | --- | 
| look_value | giá trị cần tra cứu | Tờ giấy ghi mã vận đơn mà nhân viên kho đang cầm trên tay | Mã vận đơn duy nhất ( ví dụ : `FM-10023` ) tại ô A2 của sheet danh sách đơn hàng |
| table_array | bảng chứa dữ liệu tra cứu | tủ tài liệu/Bảng giá cước lưu trữ ở căn phòng khác ( sheet `BangGiaCuoc` ). | Vùng dữ liệu từ cột mã đến cột giá cước , sử dụng tham chiếu tuyệt đối  `BangGiaCuoc!$A$2:$C$100`
|col_index_num | thứ tự cột lấy dữ liệu | số thứ tự của ngăn kéo trong tủ tài liệu chứa thông giá cước ( cột thứ N ) | Chỉ số cột chứa thông tin giá cước |
|range_lookup | phương thức tìm kiếm | quy tắc đối soát: Tìm chính xác 100% | bắt buộc chọn `FALSE` để đảm bảo mã vận đơn phải trùng hớp hoàn toàn mới lấy giá cước |
# Phần 2. Triển khai : giải pháp công nghệ cho FastMove
 ## Word 
  cấu trúc dữ liệu: 
  1. Quy trình tiếp nhận và phân loại hàng hóa ( heading 1 )

      1.1 Quét mã vận đơn và đối soát cước tự động ( heading 2 )

      1.2 kiểm tra trọng lượng và kích thước quy đổi ( heading 2 )

  2. Hướng dẫn shippe xuất giao và giao hàng ( heading 1 )
 ## Excel - nối 2 sheet khác nhau
  Đảm bảo không vi phạm 5 quy tắc vàng :
   1. Không merge cell: Tuyệt đối không gộp ô ở vùng dữ liệu để công thức kéo không bị lỗi `#VALUE` hoặc chệch hàng
   2. Dòng header duy nhất : Dòng 1 chứa tiêu đề cột ( mã vận đơn , khối lượng , vùng giao , giá cước )
   3. dữ liệu liên tục : không chứa dòng trống/cột trống xen kẽ giữa các bản ghi.
   4. Đúng dữ liệu ( database)
   5. không chứa dòng tổng cộng ở giữa dữ liệu
* Công thức VLOOKUP liên kết 2 sheet cho 500 dòng

  -sheet 1 ( `danhsachdonhang` ): chứa danhh sách 500 đơn hàng . Cột A là mã vận đơn , cột D là giá cước cần điền
  -sheet 2 ( `banggiacuoc` ): chứa bảng tra cứu từ ô A2 đến C50
  Công thức tại D2: =VLOOKUP(A2,banggiacuoc!A$2:$C$50,3,FALSE)
## PowerPoimt - slide báo cáo hiệu suất kpi và hướng dẫn shipper
 * Thiết kế slide hướng dẫn shipper nhận đơn giao hàng gồm đúng 3 bước ngắn gọn :
   1. Quét mã QR trên đơn hàng: sử dụng app FastMove driver để nhận diện mã vận đơn trong 1 giây
   2. đối soát giá cước hiển thị
   3. xác nhận xuất kho
# Phần 3. Kiểm thử: - Xử lý lỗi #N/A khi mã vận đơn có ký tự lạ bằng hàm IFERROR.
1. Nguyên nhân xuất hiện lỗi #N/A
Khi nhập Mã vận đơn mới hoặc mã có ký tự lạ/khoảng trắng thừa (ví dụ: FM-99999*), hàm VLOOKUP với tham số FALSE sẽ dò tìm chính xác 100%. Khi không tìm thấy mã này trong BangGiaCuoc, Excel sẽ trả về lỗi #N/A (Not Available).

Lỗi #N/A làm cho toàn bộ các phép tính tổng tiền, tính trung bình ở cuối bảng bị lỗi dây chuyền (#N/A).

2. Xử lý triệt để bằng hàm IFERROR kết hợp TRIM
Công thức lồng IFERROR chuẩn hóa:

Excel
=IFERROR(VLOOKUP(TRIM(A2), BangGiaCuoc!$A$2:$C$50, 3, FALSE), "Chưa có biểu cước")
Quy trình vận hành 3 bước khi phát hiện lỗi:

Bước 1 (Lọc lỗi): Bật AutoFilter (Ctrl + Shift + L), lọc cột Giá cước theo giá trị "Chưa có biểu cước" hoặc chứa #N/A.

Bước 2 (Kiểm tra ký tự lạ / khoảng trắng): Hàm TRIM(A2) đã tự động xóa khoảng trắng thừa. Nếu vẫn báo lỗi, kiểm tra xem Mã vận đơn có bị gõ sai định dạng hoặc chứa ký tự đặc biệt (như dấu *, ?, khoảng trắng không ngắt NBSP) hay không.

Bước 3 (Bổ sung Bảng giá): Nếu mã vận đơn hợp lệ nhưng chưa được đăng ký trong hệ thống, bộ phận kho tiến hành cập nhật mã mới vào BangGiaCuoc. Bảng tính Excel sẽ lập tức tự động cập nhật giá cước mà không cần sửa công thức.
# Phần 1. Phân tích : yêu cầu chức năng và phi chức năng
Trong việc xây dựng hệ thống phần mềm y tế (Healthcare Software), việc tách bạch rõ ràng giữa Yêu cầu chức năng (Functional Requirements) và Yêu cầu phi chức năng (Non-functional Requirements - đặc biệt là bảo mật) là yếu tố then chốt:
| Tiêu chí phân loại | yêu cầu chức năng | yêu cầu phi chức năng |
| --- | --- | --- | 
| Định Nghĩa | Mô tả hệ thống phải làm gì để đáp ứng các nghiệp vụ khám chữa bệnh thực tế của phòng khám | Mô tả hệ thống hoạt động ra sao về mặt chất lượng, hiệu năng và đặc biệt là chuẩn an toàn thông tin y tế |
| ví dụ | - tiếp nhận thông tin bệnh nhân | - Bảo mật dữ liệu  
|       |   - Ghi nhận huyết áp , nhiệt độ , nhịp tim | - phân quyền chặt chẽ : bác sĩ chỉ xem đc bệnh án của , lễ tân chỉ xem được thông tin hành chính |
|       | - lập phiếu khám bệnh và kê đơn | -theo dõi truy cập
# PHẦN 2. TRIỂN KHAI: HỆ THỐNG CLINIC-CLOUD TÍCH HỢP
1. Word – Tài liệu Đặc tả yêu cầu (SRS) Module Bệnh án điện tử
Quản lý cấu trúc bằng Navigation Pane: Tài liệu SRS được phân cấp chuẩn từ Heading 1 đến Heading 3, giúp dễ dàng kiểm soát các điều khoản về bảo mật dữ liệu y tế và quy trình khám bệnh.

Các nội dung cốt lõi trong SRS:

Tổng quan hệ thống phần mềm quản lý phòng khám Startup Clinic.   
JPG

Đặc tả chi tiết các tác vụ nghiệp vụ khám chữa bệnh (Functional).

Khung tiêu chuẩn bảo mật dữ liệu theo quy định y tế (Non-functional) gồm mã hóa AES-256, sao lưu định kỳ trên đám mây và cơ chế phân quyền RBAC.   
JPG

2. Excel – Quản lý 1000 hồ sơ bệnh nhân (Database Mindset & Xử lý bẫy dữ liệu)
Xử lý Bẫy dữ liệu (Ô trống do bệnh nhân không điền đủ thông tin):

Nguy cơ: Nếu bảng dữ liệu chứa hàng nghìn ô trống (Null), tính năng AutoFilter hoặc các hàm thống kê (COUNTA, VLOOKUP) sẽ dễ bị ngắt quãng, hiểu nhầm là kết thúc bảng hoặc lọc thiếu dữ liệu.

Giải pháp tiêu chuẩn Database:

Tuyệt đối không để ô trống hoàn toàn ở các trường quan trọng. Đối với dữ liệu chữ còn thiếu, chuẩn hóa bằng cách điền giá trị mặc định như "Chưa cập nhật" hoặc "N/A".

Đảm bảo tuân thủ 5 Quy tắc Vàng: Không dùng tính năng Merge Cell trong vùng dữ liệu và duy trì đúng 1 hàng tiêu đề (Header Row) ở dòng đầu tiên.   

Cấu trúc bảng quản lý mẫu (Mô phỏng giao diện minh họa):

Cột A: ID (Ví dụ: P001, P002, P003...)   

Cột B: Họ tên (Ví dụ: Nguyễn Văn A, Trần Thị B...)   

Cột C: Ngày sinh

Cột D: Giới tính (Nam / Nữ)   
JPG

Cột E: SĐT

Cột F: Trạng thái (Đang điều trị, Tái khám, Hoàn tất)   

3. PowerPoint – Slide báo cáo dịch bệnh & Trực quan hóa
Tiêu đề Assertion Title (Mang tính kết luận):

Tiêu đề sai (Chung chung): "Biểu đồ bệnh lý tháng này"

Tiêu đề chuẩn (Assertion Title): "Số hóa bệnh án giúp giảm 40% thời gian chờ đợi và bảo mật tuyệt đối 100% hồ sơ bệnh nhân"

Thiết kế trực quan:

Sử dụng Donut Chart tối giản để phân chia tỷ lệ các nhóm bệnh hoặc trạng thái điều trị của 1000 bệnh nhân.

Tích hợp khu vực hiển thị Checklist Bảo mật y tế (Phân quyền, Mã hóa dữ liệu, Sao lưu định kỳ, Theo dõi truy cập) giúp slide trực quan và súc tích.   
JPG

PHẦN 3. KIỂM THỬ: KHẮC PHỤC LỖI GÁN NHẦM HEADING STYLES BẰNG CTRL + SHIFT + N
1. Nguyên nhân và Hậu quả lỗi gán nhầm Heading
Nguyên nhân: Khi soạn thảo tài liệu SRS, người dùng vô tình bôi đen một đoạn văn bản mô tả triệu chứng hoặc ghi chú dài nhưng lại bấm nhầm vào Heading 1 hoặc Heading 2 trên thanh Ribbon.

Hậu quả: Đoạn văn bản thông thường bị biến thành tiêu đề cấp lớn, làm phình to Mục lục tự động (TOC) một cách vô lý và làm hỏng cấu trúc Navigation Pane của tài liệu.

2. Quy trình xử lý nhanh bằng tổ hợp phím Ctrl + Shift + N
Khi phát hiện đoạn văn bản bị gán nhầm style tiêu đề, bạn chỉ cần thực hiện 2 bước đơn giản sau:

Bước 1 (Chọn vùng bị lỗi): Bôi đen (hoặc đặt con trỏ chuột vào) đoạn văn bản đang bị dính định dạng Heading nhầm.

Bước 2 (Khôi phục Normal Style): Nhấn tổ hợp phím Ctrl + Shift + N.

Tác dụng: Lệnh này lập tức đưa định dạng của đoạn văn bản trở về Normal Style (văn bản thường mặc định của Word), loại bỏ hoàn toàn thuộc tính tiêu đề, giúp loại bỏ đoạn văn đó ra khỏi bảng Mục lục tự động và trả lại sự gọn gàng cho Navigation Pane.
