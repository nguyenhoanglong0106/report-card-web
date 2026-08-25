import { Hono } from 'hono';
import { processExcel, AnalyzeError } from './excel/readSheet1.js';
import { createSheet2 } from './excel/exportSheet2.js';
import { reportData, renderReportDocx } from './reports/generateReport.js';
import { safeName } from './utils/format.js';

const MAX_FILE_SIZE = 20 * 1024 * 1024;

const app = new Hono();

function fail(c, message, status = 400) {
  return c.json({ success: false, message }, status);
}

async function readUploadedFile(c) {
  let form;
  try {
    form = await c.req.formData();
  } catch {
    throw new AnalyzeError('Vui lòng chọn file Excel.');
  }
  const file = form.get('file');
  if (!file || typeof file.arrayBuffer !== 'function') {
    throw new AnalyzeError('Vui lòng chọn file Excel.');
  }
  if (file.size > MAX_FILE_SIZE) {
    throw new AnalyzeError('File vượt quá giới hạn cho phép (20MB).');
  }
  return await file.arrayBuffer();
}

app.get('/api/health', (c) => c.json({ success: true, status: 'ok', runtime: 'cloudflare-worker' }));

app.post('/api/analyze', async (c) => {
  try {
    const arrayBuffer = await readUploadedFile(c);
    const { students, meta } = await processExcel(arrayBuffer);
    return c.json({
      success: true,
      schoolYear: meta.schoolYear,
      className: meta.className,
      glv: meta.glv,
      parish: meta.parish,
      students,
    });
  } catch (e) {
    if (e instanceof AnalyzeError) return fail(c, e.message);
    return fail(c, 'Không đọc được dữ liệu từ file Excel.', 500);
  }
});

app.post('/api/export-excel', async (c) => {
  try {
    const arrayBuffer = await readUploadedFile(c);
    const { wb, students, meta } = await processExcel(arrayBuffer);
    createSheet2(wb, students, meta);
    const buffer = await wb.xlsx.writeBuffer({ zip: { compression: 'STORE' } });
    c.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    c.header('Content-Disposition', 'attachment; filename="KetQuaHocTap_DaXuLy.xlsx"');
    return c.body(buffer);
  } catch (e) {
    if (e instanceof AnalyzeError) return fail(c, e.message);
    return fail(c, 'Không thể tạo file Excel.', 500);
  }
});

app.post('/api/report', async (c) => {
  try {
    const body = await c.req.json();
    const { student, schoolYear, className, teacher, phone, contactDate } = body || {};
    if (!student) return fail(c, 'Thiếu dữ liệu học sinh.');
    const data = reportData(student, { schoolYear, className, glv: teacher }, { schoolYear, className, glv: teacher, phone, contactDate });
    const out = renderReportDocx(data);
    c.header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    c.header('Content-Disposition', `attachment; filename="PhieuBaoDiem_${safeName(student.fullName)}.docx"`);
    return c.body(out);
  } catch (e) {
    return fail(c, e.message || 'Không thể tạo phiếu Word.', 500);
  }
});

app.notFound((c) => c.json({ success: false, message: 'Không tìm thấy.' }, 404));
app.onError((err, c) => {
  console.error(err);
  return c.json({ success: false, message: 'Có lỗi xảy ra trên máy chủ.' }, 500);
});

export default app;
