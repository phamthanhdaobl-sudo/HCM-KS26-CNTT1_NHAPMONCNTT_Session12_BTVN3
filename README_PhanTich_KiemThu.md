FASTMOVE — BÀI 3: CHUẨN HÓA QUY TRÌNH GIAO VẬN VÀ TRA CỨU GIÁ CƯỚC

Bộ bài nộp gồm:
1. FastMove_QuyTrinhVanHanhKho.docx — Quy trình vận hành kho có Styles, màu thương hiệu xanh navy/xanh teal.
2. FastMove_TraCuuGiaCuoc_KPI.xlsx — Hai sheet DonHang và GiaCuoc; công thức VLOOKUP có vùng tuyệt đối; sheet KPI có Bar Chart.
3. FastMove_HuongDanShipper_KPI.pptx — Slide hướng dẫn shipper, tối đa 4 bullet/slide, tiêu đề dạng kết luận.

PHẦN 1 — PHÂN TÍCH: 4 THAM SỐ VLOOKUP
- lookup_value: mã dịch vụ ở đơn cần tra — giống mã hồ sơ nhân viên đưa ra.
- table_array: GiaCuoc!$A$2:$B$6 — giống tủ/ngăn hồ sơ cố định. Dấu $ giữ nguyên vùng khi kéo công thức.
- col_index_num: 2 — lấy cột thứ hai trong vùng, tức Giá cước.
- range_lookup: FALSE — chỉ chấp nhận khớp chính xác, tránh lấy nhầm hồ sơ gần giống.

PHẦN 2 — TRIỂN KHAI
- Chuẩn hóa mã dịch vụ và dữ liệu đơn hàng.
- Dùng công thức tại cột Giá cước: =IFERROR(VLOOKUP(B2,GiaCuoc!$A$2:$B$6,2,FALSE),"Không tìm thấy mã")
- Kéo công thức xuống 500 dòng. Vùng bảng giá được khóa bằng $.
- Dùng Bar Chart để so sánh hiệu suất hiện tại với mục tiêu và làm nổi bật điểm nghẽn 16 giờ/ngày.
- Trong PowerPoint, tiêu đề slide là kết luận; mỗi slide giữ tối đa 3–4 ý theo Miller's Law.

PHẦN 3 — KIỂM THỬ LỖI #N/A
- Sheet DonHang có dòng kiểm thử FM-TEST-NA với mã dịch vụ FM-NOI# (ký tự lạ).
- Kết quả kỳ vọng ở cột Giá cước là “Không tìm thấy mã”, không hiển thị #N/A.
- Nếu vẫn không khớp: kiểm tra khoảng trắng đầu/cuối, ký tự ẩn, định dạng Text/Number và danh mục mã dịch vụ.
- Không dùng IFERROR để che lỗi nghiệp vụ; cần theo dõi các dòng “Không tìm thấy mã” và sửa dữ liệu nguồn.

5 QUY TẮC VÀNG DỮ LIỆU ĐƯỢC ÁP DỤNG
1. Một dòng cho một bản ghi.
2. Một cột cho một trường dữ liệu.
3. Không gộp ô trong vùng dữ liệu.
4. Không để dòng/cột trống xen giữa bảng dữ liệu.
5. Dùng tiêu đề cột rõ ràng, định dạng nhất quán và không dùng màu làm thông tin duy nhất.

Lưu ý: dữ liệu trong workbook là dữ liệu minh họa phục vụ bài tập, không phải dữ liệu vận hành thực tế.
