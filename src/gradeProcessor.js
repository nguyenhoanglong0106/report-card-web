const ExcelJS = require('exceljs');

const round2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
const num = (v) => {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(',', '.').trim());
  return Number.isFinite(n) ? n : null;
};
const text = (v) => (v === null || v === undefined ? '' : String(v).trim());
const attendanceValue = (v) => {
  if (typeof v === 'number') return v;
  const m = text(v).match(/-?\d+(?:[.,]\d+)?/);
  return m ? Number(m[0].replace(',', '.')) : 0;
};
const classify = (score) => {
  const s = num(score) ?? 0;
  if (s >= 9) return 'Xuất Sắc';
  if (s >= 8) return 'Giỏi';
  if (s >= 6.5) return 'Khá';
  if (s >= 5) return 'Trung Bình';
  return 'Yếu';
};
const rankScores = (students, key) => {
  const vals = students.map(s => num(s[key]) ?? -Infinity);
  students.forEach((s, i) => {
    const v = vals[i];
    s[key + 'Rank'] = 1 + vals.filter(x => x > v).length;
  });
};
const extractMeta = (sheet) => {
  const title = text(sheet.getCell('A4').value);
  const years = title.match(/\b20\d{2}\s*[-–]\s*20\d{2}\b/);
  let className = '';
  const lines = title.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const detail = lines.find(s => /20\d{2}\s*[-–]\s*20\d{2}/.test(s)) || lines[lines.length - 1] || '';
  if (detail) {
    className = detail.replace(/KẾT QUẢ HỌC TẬP CHI TIẾT/i, '').replace(/\s*[-–]\s*[^-–]*\s*[-–]\s*20\d{2}\s*[-–]\s*20\d{2}.*/i, '').trim();
    if (!className) className = detail.split(/\s+-\s+/)[0].trim();
  }
  const glvs = [text(sheet.getCell('A2').value), text(sheet.getCell('A3').value)].filter(Boolean);
  return {
    parish: text(sheet.getCell('A1').value),
    schoolYear: years ? years[0].replace(/\s/g, '') : '',
    className,
    glv: glvs.join(' - ')
  };
};

async function readWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet1 = wb.getWorksheet('Sheet1') || wb.worksheets[0];
  if (!sheet1) throw new Error('Không tìm thấy Sheet1 trong file Excel.');
  return { wb, sheet1 };
}

function buildStudents(sheet1) {
  const students = [];
  for (let r = 10; r <= sheet1.rowCount + 5; r++) {
    const stt = sheet1.getCell(r, 1).value;
    const code = sheet1.getCell(r, 2).value;
    if ((stt === null || stt === '') && (code === null || code === '')) continue;
    if (!code && !sheet1.getCell(r, 5).value && !sheet1.getCell(r, 6).value) continue;

    const s = {
      sourceRow: r,
      stt: text(stt), code: text(code), saintName: text(sheet1.getCell(r, 4).value),
      lastName: text(sheet1.getCell(r, 5).value), firstName: text(sheet1.getCell(r, 6).value),
      dob: text(sheet1.getCell(r, 7).text || sheet1.getCell(r, 7).value), status: text(sheet1.getCell(r, 8).value),
      hk1Quiz: num(sheet1.getCell(r, 11).value), hk1p15: num(sheet1.getCell(r, 12).value),
      hk1Mid: num(sheet1.getCell(r, 13).value), hk1Exam: num(sheet1.getCell(r, 14).value),
      hk1Academic: num(sheet1.getCell(r, 15).value), hk1Mass: attendanceValue(sheet1.getCell(r, 16).value),
      hk1Class: attendanceValue(sheet1.getCell(r, 17).value), hk1Avg: num(sheet1.getCell(r, 21).value),
      hk2Quiz: num(sheet1.getCell(r, 22).value), hk2p15: num(sheet1.getCell(r, 23).value),
      hk2Mid: num(sheet1.getCell(r, 24).value), hk2Exam: num(sheet1.getCell(r, 25).value),
      hk2Academic: num(sheet1.getCell(r, 26).value), hk2Mass: attendanceValue(sheet1.getCell(r, 27).value),
      hk2Class: attendanceValue(sheet1.getCell(r, 28).value), hk2Avg: num(sheet1.getCell(r, 32).value)
    };
    s.hk1Attendance = s.hk1Mass + s.hk1Class;
    s.hk2Attendance = s.hk2Mass + s.hk2Class;
    s.totalAttendance = s.hk1Attendance + s.hk2Attendance;
    s.yearAcademic = round2(((s.hk1Avg ?? 0) + (s.hk2Avg ?? 0)) / 2);
    s.yearAvg = s.yearAcademic;
    s.hk1Classify = classify(s.hk1Avg);
    s.hk2Classify = classify(s.hk2Avg);
    s.yearClassify = classify(s.yearAvg);
    s.result = s.yearAcademic < 3.5 ? 'Ở lại' : 'Lên Lớp';
    s.fullName = [s.saintName, s.lastName, s.firstName].filter(Boolean).join(' ');
    students.push(s);
  }
  rankScores(students, 'hk1Avg');
  rankScores(students, 'hk2Avg');
  rankScores(students, 'yearAvg');
  return students;
}

function styleSheet2(ws, count) {
  ws.views = [{ state: 'frozen', ySplit: 9 }];
  const widths = [6,12,14,22,12,13,14,9,9,9,9,10,11,13,11,12,9,12,9,9,9,9,10,11,13,11,12,9,12,11,10,10,10,11,11];
  widths.forEach((w, i) => ws.getColumn(i + 1).width = w);
  for (let r = 6; r <= 9 + count; r++) {
    ws.getRow(r).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  ['A6:AI9'].forEach(range => {
    ws.getCell('A6');
    const rows = ws.getRows(6, 4);
    rows.forEach(row => row.eachCell({includeEmpty:true}, c => {
      c.font = { name: 'Calibri', size: 10, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      c.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    }));
  });
  for (let r = 10; r < 10 + count; r++) {
    ws.getRow(r).eachCell({includeEmpty:true}, c => {
      c.font = { name: 'Calibri', size: 10 };
      c.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws.getCell(r,4).alignment = { horizontal:'left', vertical:'middle' };
  }
}

function createSheet2(wb, students, meta) {
  const old = wb.getWorksheet('Sheet2');
  if (old) wb.removeWorksheet(old.id);
  const ws = wb.addWorksheet('Sheet2');
  ws.getCell('A1').value = meta.parish || 'GIÁO XỨ BẮC THẦN';
  ws.getCell('A2').value = meta.glv.split(' - ')[0] || '';
  ws.getCell('A3').value = meta.glv.split(' - ')[1] || '';
  ws.getCell('A4').value = `KẾT QUẢ HỌC TẬP CHI TIẾT\n ${meta.className || ''} - ${meta.schoolYear || ''}`;

  const headers = {
    A6:'Stt',B6:'Mã số',C6:'Tên Thánh',D6:'Họ',E6:'Tên',F6:'Ngày sinh',G6:'Tình trạng',
    H6:'Học kỳ I (Quy định 100% học lực và 0% C. Cần)',Q6:'Xếp hạng',R6:'Xếp loại',
    S6:'Học kỳ II (Quy định 100% học lực và 0% C. Cần)',AB6:'Xếp hạng',AC6:'Xếp loại',
    AD6:'TB C.Cần',AE6:'TB H.Lực',AF6:'Tổng TB',AG6:'Xếp loại',AH6:'Xếp hạng',AI6:'Kết quả',
    G7:'Tình trạng',H7:'Học lực',M7:'Chuyên cần',S7:'Học lực',X7:'Chuyên cần',
    H8:'Đầu giờ',I8:'15 phút',J8:'45 phút',K8:'Thi',L8:'Tb. Học lực',M8:'Số lần dự T.Lễ',N8:'Số lần đi học G.Lý',O8:'Tb. Chuyên cần',P8:'Trung bình học kỳ',
    S8:'Đầu giờ',T8:'15 phút',U8:'45 phút',V8:'Thi',W8:'Tb. Học lực',X8:'Số lần dự T.Lễ',Y8:'Số lần đi học G.Lý',Z8:'Tb. Chuyên cần',AA8:'Trung bình học kỳ',
    D9:'Hệ số',H9:1,I9:1,J9:2,K9:3,S9:1,T9:1,U9:2,V9:3
  };
  Object.entries(headers).forEach(([a,v]) => ws.getCell(a).value = v);
  ws.mergeCells('H6:P6'); ws.mergeCells('S6:AA6');
  ws.mergeCells('H7:L7'); ws.mergeCells('M7:P7'); ws.mergeCells('S7:W7'); ws.mergeCells('X7:AA7');

  students.forEach((s, idx) => {
    const r = 10 + idx;
    const values = [s.stt,s.code,s.saintName,s.lastName,s.firstName,s.dob,s.status,s.hk1Quiz,s.hk1p15,s.hk1Mid,s.hk1Exam,s.hk1Academic,s.hk1Mass,s.hk1Class,null,s.hk1Avg,null,null,s.hk2Quiz,s.hk2p15,s.hk2Mid,s.hk2Exam,s.hk2Academic,s.hk2Mass,s.hk2Class,null,s.hk2Avg,null,null,null,null,null,null,null,null];
    values.forEach((v,i)=> ws.getCell(r,i+1).value=v);
    ws.getCell(r,15).value={formula:`M${r}+N${r}`,result:s.hk1Attendance};
    ws.getCell(r,17).value={formula:`RANK(P${r},$P$10:$P$${9+students.length},0)`,result:s.hk1AvgRank};
    ws.getCell(r,18).value={formula:`IF(P${r}>=9,"Xuất Sắc",IF(P${r}>=8,"Giỏi",IF(P${r}>=6.5,"Khá",IF(P${r}>=5,"Trung Bình","Yếu"))))`,result:s.hk1Classify};
    ws.getCell(r,26).value={formula:`X${r}+Y${r}`,result:s.hk2Attendance};
    ws.getCell(r,28).value={formula:`RANK(AA${r},$AA$10:$AA$${9+students.length},0)`,result:s.hk2AvgRank};
    ws.getCell(r,29).value={formula:`IF(AA${r}>=9,"Xuất Sắc",IF(AA${r}>=8,"Giỏi",IF(AA${r}>=6.5,"Khá",IF(AA${r}>=5,"Trung Bình","Yếu"))))`,result:s.hk2Classify};
    ws.getCell(r,30).value={formula:`O${r}+Z${r}`,result:s.totalAttendance};
    ws.getCell(r,31).value={formula:`ROUND((P${r}+AA${r})/2,2)`,result:s.yearAcademic};
    ws.getCell(r,32).value={formula:`AE${r}`,result:s.yearAvg};
    ws.getCell(r,33).value={formula:`RANK(AF${r},$AF$10:$AF$${9+students.length},0)`,result:s.yearAvgRank};
    ws.getCell(r,34).value={formula:`IF(AF${r}>=9,"Xuất Sắc",IF(AF${r}>=8,"Giỏi",IF(AF${r}>=6.5,"Khá",IF(AF${r}>=5,"Trung Bình","Yếu"))))`,result:s.yearClassify};
    ws.getCell(r,35).value={formula:`IF(AE${r}<3.5,"Ở lại","Lên Lớp")`,result:s.result};
  });
  styleSheet2(ws, students.length);
  return ws;
}

async function processExcel(buffer) {
  const { wb, sheet1 } = await readWorkbook(buffer);
  const meta = extractMeta(sheet1);
  const students = buildStudents(sheet1);
  if (!students.length) throw new Error('Không tìm thấy dữ liệu học sinh từ dòng 10 của Sheet1.');
  createSheet2(wb, students, meta);
  return { wb, students, meta };
}

function fmt(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(round2(v));
  return String(v);
}

function reportData(s, meta, form={}) {
  return {
    NIEN_KHOA: form.schoolYear || meta.schoolYear || '', HOC_SINH: s.fullName,
    NGAY_SINH: s.dob, LOP: form.className || meta.className || '', GLV: form.glv || meta.glv || '',
    SDT: form.phone || '', NGAY_LIEN_HE: form.contactDate || '', KET_QUA: s.result,
    HK1_KHAO_BAI:fmt(s.hk1Quiz),HK1_15P:fmt(s.hk1p15),HK1_GIUA_KY:fmt(s.hk1Mid),HK1_HOC_KY:fmt(s.hk1Exam),HK1_TB:fmt(s.hk1Avg),HK1_XEP_LOAI:s.hk1Classify,
    HK1_DI_LE:fmt(s.hk1Mass),HK1_DI_HOC_GL:fmt(s.hk1Class),HK1_TB_CC:fmt(s.hk1Attendance),HK1_XEP_HANG:fmt(s.hk1AvgRank),
    HK2_KHAO_BAI:fmt(s.hk2Quiz),HK2_15P:fmt(s.hk2p15),HK2_GIUA_KY:fmt(s.hk2Mid),HK2_HOC_KY:fmt(s.hk2Exam),HK2_TB:fmt(s.hk2Avg),HK2_XEP_LOAI:s.hk2Classify,
    HK2_DI_LE:fmt(s.hk2Mass),HK2_DI_HOC_GL:fmt(s.hk2Class),HK2_TB_CC:fmt(s.hk2Attendance),HK2_XEP_HANG:fmt(s.hk2AvgRank),
    CN_KHAO_BAI:fmt(round2(((s.hk1Quiz||0)+(s.hk2Quiz||0))/2)),CN_15P:fmt(round2(((s.hk1p15||0)+(s.hk2p15||0))/2)),
    CN_GIUA_KY:fmt(round2(((s.hk1Mid||0)+(s.hk2Mid||0))/2)),CN_HOC_KY:fmt(round2(((s.hk1Exam||0)+(s.hk2Exam||0))/2)),
    CN_TB:fmt(s.yearAvg),CN_XEP_LOAI:s.yearClassify,CN_DI_LE:fmt(s.hk1Mass+s.hk2Mass),CN_DI_HOC_GL:fmt(s.hk1Class+s.hk2Class),
    CN_TB_CC:fmt(s.totalAttendance),CN_XEP_HANG:fmt(s.yearAvgRank)
  };
}

module.exports={processExcel, reportData};
