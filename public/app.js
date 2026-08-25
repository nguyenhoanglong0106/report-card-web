import { mergeDocx } from './zip-lite.js';

const $ = (s) => document.querySelector(s);
let currentFile = null;
let students = [];

const file = $('#file'), msg = $('#msg'), tbody = $('#tbody');
const summaryTotal = $('#summaryTotal'), summaryRows = $('#summaryRows');
const classifyOrder = ['Xuất Sắc', 'Giỏi', 'Khá', 'Trung Bình', 'Yếu'];

function percent(count, total) {
  if (!total) return '0%';
  const value = (count / total) * 100;
  return `${Number.isInteger(value) ? value : value.toFixed(1)}%`;
}

function renderSummary() {
  const total = students.length;
  summaryTotal.textContent = `${total} học sinh`;
  const counts = Object.fromEntries(classifyOrder.map((name) => [name, 0]));
  students.forEach((student) => {
    const key = classifyOrder.find((name) => name.toLowerCase() === String(student.yearClassify || '').toLowerCase());
    if (key) counts[key]++;
  });
  summaryRows.innerHTML = classifyOrder
    .map((name) => {
      const count = counts[name];
      const pct = percent(count, total);
      return `<div class="summary-row">
        <span>${name}</span>
        <div class="summary-bar"><i style="width:${pct}"></i></div>
        <b>${count}</b>
        <em>${pct}</em>
      </div>`;
    })
    .join('');
}

function setMsg(t, type = '') {
  msg.textContent = t;
  msg.className = 'msg ' + type;
}

function formValues() {
  return {
    schoolYear: $('#schoolYear').value,
    className: $('#className').value,
    teacher: $('#glv').value,
    phone: $('#phone').value,
    contactDate: $('#contactDate').value,
  };
}

async function extractErrorMessage(res) {
  try {
    const data = await res.json();
    return data.message || data.error || `Có lỗi xảy ra (${res.status}).`;
  } catch {
    return `Có lỗi xảy ra (${res.status}).`;
  }
}

function triggerDownload(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

async function fetchReportDocx(student) {
  const r = await fetch('/api/report', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ student, ...formValues() }),
  });
  if (!r.ok) throw new Error(await extractErrorMessage(r));
  return r.arrayBuffer();
}

async function downloadOne(student) {
  setMsg('Đang tạo phiếu...');
  try {
    const buf = await fetchReportDocx(student);
    triggerDownload(new Blob([buf]), `PhieuBaoDiem_${student.fullName}.docx`);
    setMsg('Đã tạo file thành công.', 'ok');
  } catch (e) {
    setMsg(e.message || 'Không thể tạo phiếu Word.', 'error');
  }
}

// Fetches every student's report individually (each request is small and
// fast) and merges them into one .docx in the browser — keeps the Worker from
// ever having to render a whole class inside a single request.
async function downloadAll() {
  if (!students.length) return;
  setMsg(`Đang tạo phiếu cả lớp (0/${students.length})...`);
  try {
    const CONCURRENCY = 4;
    const buffers = new Array(students.length);
    let next = 0;
    let done = 0;
    async function worker() {
      while (next < students.length) {
        const i = next++;
        buffers[i] = await fetchReportDocx(students[i]);
        done++;
        setMsg(`Đang tạo phiếu cả lớp (${done}/${students.length})...`);
      }
    }
    await Promise.all(Array.from({ length: Math.min(CONCURRENCY, students.length) }, worker));
    setMsg('Đang ghép phiếu cả lớp...');
    const merged = await mergeDocx(buffers);
    triggerDownload(merged, 'PhieuBaoDiem_CaLop.docx');
    setMsg('Đã tạo file thành công.', 'ok');
  } catch (e) {
    setMsg(e.message || 'Không thể tạo phiếu cả lớp.', 'error');
  }
}

async function downloadExcel() {
  if (!currentFile) return;
  setMsg('Đang tạo file...');
  try {
    const fd = new FormData();
    fd.append('file', currentFile);
    const r = await fetch('/api/export-excel', { method: 'POST', body: fd });
    if (!r.ok) throw new Error(await extractErrorMessage(r));
    const b = await r.blob();
    triggerDownload(b, 'KetQuaHocTap_DaXuLy.xlsx');
    setMsg('Đã tạo file thành công.', 'ok');
  } catch (e) {
    setMsg(e.message || 'Không thể kết nối tới máy chủ.', 'error');
  }
}

function renderTable() {
  renderSummary();
  if (!students.length) {
    tbody.innerHTML = '<tr><td colspan="15" class="empty">Chưa có dữ liệu</td></tr>';
    return;
  }
  tbody.innerHTML = students
    .map(
      (s, i) => `<tr>
        <td>${s.stt ?? ''}</td><td>${s.code ?? ''}</td><td>${s.saintName ?? ''}</td><td>${s.lastName ?? ''}</td><td>${s.firstName ?? ''}</td>
        <td>${s.dob ?? ''}</td><td>${s.hk1Avg ?? ''}</td><td>${s.hk1Classify ?? ''}</td><td>${s.hk2Avg ?? ''}</td><td>${s.hk2Classify ?? ''}</td>
        <td>${s.yearAvg ?? ''}</td><td>${s.yearAvgRank ?? ''}</td><td>${s.yearClassify ?? ''}</td><td>${s.result ?? ''}</td>
        <td><button class="mini" data-i="${i}">Tải phiếu</button></td>
      </tr>`
    )
    .join('');
}

file.addEventListener('change', async () => {
  currentFile = file.files[0];
  $('#fileName').textContent = currentFile?.name || 'Chưa chọn file';
  if (!currentFile) return;
  setMsg('Đang đọc dữ liệu...');
  students = [];
  renderTable();
  $('#excelBtn').disabled = true;
  $('#allBtn').disabled = true;
  try {
    const fd = new FormData();
    fd.append('file', currentFile);
    const r = await fetch('/api/analyze', { method: 'POST', body: fd });
    const data = await r.json();
    if (!r.ok || !data.success) throw new Error(data.message || 'Không đọc được dữ liệu.');
    students = data.students;
    $('#schoolYear').value = data.schoolYear || '';
    $('#className').value = data.className || '';
    $('#glv').value = data.glv || '';
    $('#count').textContent = `${students.length} học sinh`;
    renderTable();
    $('#excelBtn').disabled = false;
    $('#allBtn').disabled = false;
    setMsg(`Đọc thành công ${students.length} học sinh.`, 'ok');
  } catch (e) {
    students = [];
    renderTable();
    $('#count').textContent = '';
    setMsg(e.message || 'Không đọc được dữ liệu.', 'error');
  }
});

tbody.addEventListener('click', (e) => {
  const b = e.target.closest('[data-i]');
  if (!b) return;
  downloadOne(students[b.dataset.i]);
});

$('#excelBtn').addEventListener('click', downloadExcel);
$('#allBtn').addEventListener('click', downloadAll);
renderSummary();
