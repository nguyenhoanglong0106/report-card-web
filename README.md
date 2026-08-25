# Web tạo phiếu báo điểm (Cloudflare Workers)

Ứng dụng serverless chạy hoàn toàn trên Cloudflare Workers — không cần VPS,
không cần Node server chạy 24/7. Sau khi deploy, máy tính cá nhân có thể tắt,
web vẫn hoạt động qua URL Cloudflare cấp.

## Chức năng
- Upload file Excel kết quả học tập, đọc dữ liệu từ `Sheet1`.
- Tính xếp loại, xếp hạng (tương đương `RANK.EQ`), tổng chuyên cần, TB cả năm.
- Xem danh sách học sinh ngay trên trình duyệt, sau đó chủ động chọn xuất file.
- Tải Excel đã có `Sheet2` theo đúng định dạng mẫu (giữ nguyên `Sheet1`).
- Tải phiếu Word cho từng học sinh, hoặc phiếu cả lớp (1 file `.docx`, mỗi học
  sinh 1 trang, ngắt trang tự động) — mẫu Word nằm cố định trong project.

## Kiến trúc

```
report-card-web/
├── src/
│   ├── worker.js              Hono app — toàn bộ route /api/*
│   ├── excel/
│   │   ├── readSheet1.js      Đọc & parse Sheet1 (ExcelJS)
│   │   ├── calculateGrades.js Xếp loại + xếp hạng (RANK.EQ)
│   │   └── exportSheet2.js    Dựng lại Sheet2 (style, formula, merge)
│   ├── reports/
│   │   └── generateReport.js  Render Word bằng thay thế {TAG} trực tiếp trên XML
│   └── utils/
│       └── format.js          Hàm dùng chung (round2, fmt, safeName...)
├── public/                    Frontend, deploy chung qua Cloudflare Assets
│   ├── index.html
│   ├── app.js                 Toàn bộ logic UI + gọi API
│   └── zip-lite.js            ZIP reader/writer tối giản, dùng để ghép phiếu cả lớp
│   └── style.css
├── templates/
│   └── report-card-template.docx   Mẫu Word, bundle thẳng vào Worker
├── wrangler.jsonc
└── package.json
```

**Stateless hoàn toàn**: Worker không lưu session, không lưu file, không dùng
database. Sau khi `/api/analyze` trả JSON, trình duyệt giữ toàn bộ dữ liệu
học sinh trong bộ nhớ JS. Khi tải phiếu Word, trình duyệt gửi thẳng dữ liệu
JSON đã tính về `/api/report` — không cần gửi lại file Excel. Khi tải Excel
Sheet2, trình duyệt gửi lại file Excel gốc (bắt buộc, vì cần giữ nguyên
`Sheet1`) và Worker tính lại (kết quả xác định, không đổi giữa các lần).

**Mẫu Word** được import trực tiếp dưới dạng binary (`import ... from
"../../templates/report-card-template.docx"`, cấu hình `rules` kiểu `Data`
trong `wrangler.jsonc`) — bundle thẳng vào Worker script, không cần
`fs.readFileSync`, không public URL riêng cho file mẫu.

**Phiếu Word**: không dùng docxtemplater — mẫu chỉ có placeholder dạng
`{TAG}` phẳng (không loop/điều kiện), nên Worker thay thế trực tiếp trên
chuỗi XML của `word/document.xml` (đã parse sẵn một lần lúc khởi động), rồi
đóng gói lại bằng PizZip với nén `STORE`. Cách này rẻ hơn nhiều lần so với
việc dựng DOM đầy đủ của docxtemplater — quan trọng vì Cloudflare Workers Free
plan chỉ cho **10ms CPU/request** (xem phần "Giới hạn CPU" bên dưới).

**Phiếu cả lớp**: thay vì gộp nhiều học sinh trong một request phía Worker
(sẽ vượt xa 10ms), trình duyệt tự gọi `/api/report` cho từng học sinh (mỗi
request rẻ, ~5-10ms) rồi **ghép các file `.docx` lại ngay trên trình
duyệt** bằng `public/zip-lite.js` — một ZIP reader/writer tối giản (không cần
thư viện ngoài), ghép `word/document.xml` của từng học sinh với ngắt trang
(`<w:br w:type="page"/>`) giữa các trang, dùng chung phần media/style từ học
sinh đầu tiên (giống hệt nhau vì cùng một mẫu). Nhờ vậy Worker không bao giờ
phải xử lý cả lớp trong một request.

## API

| Route | Method | Vào | Ra |
|---|---|---|---|
| `/api/health` | GET | — | trạng thái Worker |
| `/api/analyze` | POST | `multipart/form-data` (`file`) | JSON: `schoolYear`, `className`, `glv`, `students[]` |
| `/api/export-excel` | POST | `multipart/form-data` (`file` gốc) | file `.xlsx` (Sheet1 + Sheet2) |
| `/api/report` | POST | JSON `{ student, schoolYear, className, teacher, phone, contactDate }` | file `.docx` 1 học sinh |

Không có route "phiếu cả lớp" phía server — trình duyệt tự lặp gọi
`/api/report` rồi ghép (xem trên).

## Giới hạn CPU trên Cloudflare Free plan (quan trọng)

Cloudflare Workers **Free plan giới hạn 10ms CPU time cho mỗi request** (khác
với giới hạn RAM/bộ nhớ). Đây là giới hạn cứng của nền tảng, không phải bug
của project. Đã đo thực tế qua `wrangler dev` (runtime Workers thật):

| Thao tác | CPU time đo được (warm) |
|---|---|
| `/api/report` (1 phiếu Word) | ~5–11ms — **vừa trong giới hạn Free** |
| `/api/analyze` | ~8–13ms — sát ngưỡng |
| `/api/export-excel` | ~19–29ms — **vượt giới hạn Free** |

Phần Word đã được tối ưu tối đa (bỏ docxtemplater, dùng thay thế chuỗi trực
tiếp, nén STORE) nên nằm gọn trong 10ms. Phần Excel dùng ExcelJS để giữ đúng
style/công thức/merge của Sheet2 — bản thân engine đọc/ghi `.xlsx` của
ExcelJS đã tốn ~15–25ms bất kể tối ưu thêm, vì đây là chi phí dựng
`styles.xml`/`theme.xml`/shared-strings vốn có của định dạng OOXML. Viết lại
toàn bộ engine đọc/ghi Excel bằng tay để né nốt phần này là một việc rất lớn
và rủi ro (dễ tạo file `.xlsx` hỏng/sai style), nên project **giữ nguyên
ExcelJS** và chấp nhận rủi ro: trên Free plan, `/api/analyze` và
`/api/export-excel` **có thể thỉnh thoảng lỗi** (Cloudflare trả lỗi 1102 —
"exceeded resource limits") với lớp học lớn. Khi gặp lỗi này:

- Người dùng thấy thông báo lỗi rõ ràng, có thể **thử lại** — không mất dữ
  liệu, không crash.
- Nếu lỗi này xảy ra thường xuyên và gây khó chịu, cách xử lý dứt điểm là
  nâng cấp **Cloudflare Workers Paid ($5/tháng)** — giới hạn CPU tăng lên 30s
  (gấp ~1000 lần nhu cầu thực tế), không cần sửa code gì thêm.

## Chạy local

```powershell
npm install
npm run dev
```

Wrangler sẽ khởi động dev server (mặc định `http://localhost:8787`) chạy đúng
runtime Cloudflare Workers (không phải Node thường) — dùng để test trước khi
deploy. Lưu ý: `wrangler dev` **không** enforce giới hạn 10ms CPU của Free
plan, nên mọi thứ sẽ luôn chạy được ở local dù có vượt ngưỡng khi lên
production Free plan thật.

## Deploy lên Cloudflare

```powershell
npm install
npx wrangler login
npm run deploy
```

Sau khi deploy xong, Cloudflare trả về URL dạng:

```
https://phieu-bao-diem.<account>.workers.dev
```

Mở URL đó là dùng được ngay — không cần server nào khác, có thể tắt máy tính
cá nhân.

## Quy ước file Excel đầu vào

Dữ liệu học sinh bắt đầu từ dòng 10 của `Sheet1`, đúng theo file mẫu: cột
K:N điểm HKI, O TB học lực HKI, P:Q chuyên cần; V:Y điểm HKII, Z TB học lực
HKII, AA:AB chuyên cần, U/AF là trung bình học kỳ. Nếu không tìm thấy sheet
tên đúng là `Sheet1`, hệ thống báo lỗi thay vì đoán sheet khác.

## Ghi chú kỹ thuật

- ExcelJS không tự tính công thức như Excel — Sheet2 ghi cả `formula` lẫn
  `result` để mở file là thấy kết quả ngay, không cần Excel tính lại.
- Không dùng `nodejs_compat`: ExcelJS và PizZip chạy tốt trên Workers runtime
  thuần (đã kiểm tra qua `wrangler dev`), nên không bật flag không cần thiết.
- File Word/Excel xuất ra dùng nén `STORE` (không nén) thay vì `DEFLATE` để
  giảm CPU time — file lớn hơn một chút nhưng vẫn nhỏ, không đáng kể so với
  lợi ích về tốc độ.
- Giới hạn file upload: 20MB.
