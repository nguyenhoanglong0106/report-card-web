# Web tạo phiếu báo điểm

## Chức năng
- Upload file Excel kết quả học tập, đọc dữ liệu từ `Sheet1`.
- Tự tạo lại `Sheet2` theo thứ tự cột của file mẫu.
- Tự tính tổng chuyên cần, trung bình năm, xếp hạng, xếp loại và kết quả lên lớp.
- Mẫu `templates/report-card-template.docx` được gắn cố định trong project.
- Tải Excel đã có Sheet2.
- Tải phiếu Word từng học sinh hoặc toàn bộ lớp dưới dạng ZIP.

## Chạy
1. Cài Node.js 20 LTS trở lên.
2. Mở Terminal tại thư mục project.
3. Chạy `npm install`.
4. Chạy `npm start`.
5. Mở `http://localhost:3000`.

## Quy ước file đầu vào
Dữ liệu học sinh bắt đầu từ dòng 10 giống file mẫu. Các cột Sheet1 dùng như file bạn cung cấp: K:N điểm HKI, O TB học lực HKI, P:Q chuyên cần; V:Y điểm HKII, Z TB học lực HKII, AA:AB chuyên cần, U/AF là trung bình học kỳ.

## Lưu ý
ExcelJS không tự tính công thức như Excel. Project ghi cả `formula` và `result` vào Sheet2 để khi mở file vẫn thấy kết quả ngay.
