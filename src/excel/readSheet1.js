import ExcelJS from 'exceljs';
import { round2, num, text, attendanceValue } from '../utils/format.js';
import { classify, rankScores } from './calculateGrades.js';

export class AnalyzeError extends Error {}

function extractMeta(sheet) {
  const title = text(sheet.getCell('A4').value);
  const years = title.match(/\b20\d{2}\s*[-–]\s*20\d{2}\b/);
  let className = '';
  const lines = title.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  const detail = lines.find((s) => /20\d{2}\s*[-–]\s*20\d{2}/.test(s)) || lines[lines.length - 1] || '';
  if (detail) {
    className = detail
      .replace(/KẾT QUẢ HỌC TẬP CHI TIẾT/i, '')
      .replace(/\s*[-–]\s*[^-–]*\s*[-–]\s*20\d{2}\s*[-–]\s*20\d{2}.*/i, '')
      .trim();
    if (!className) className = detail.split(/\s+-\s+/)[0].trim();
  }
  const glvs = [text(sheet.getCell('A2').value), text(sheet.getCell('A3').value)].filter(Boolean);
  return {
    parish: text(sheet.getCell('A1').value),
    schoolYear: years ? years[0].replace(/\s/g, '') : '',
    className,
    glv: glvs.join(' - '),
  };
}

function buildStudents(sheet1) {
  const students = [];
  // Snapshot rowCount once: ExcelJS's rowCount is a live view of the underlying
  // row array, and getCell()/getRow() extend that array on any access past the
  // current bound. Re-reading sheet1.rowCount in the loop condition would make
  // touching one new empty row push the bound forward by one every iteration,
  // an unbounded loop that eventually exhausts memory.
  const lastRow = sheet1.rowCount;
  for (let r = 10; r <= lastRow + 5; r++) {
    const stt = sheet1.getCell(r, 1).value;
    const code = sheet1.getCell(r, 2).value;
    if ((stt === null || stt === '') && (code === null || code === '')) continue;
    if (!code && !sheet1.getCell(r, 5).value && !sheet1.getCell(r, 6).value) continue;

    const s = {
      sourceRow: r,
      stt: text(stt),
      code: text(code),
      saintName: text(sheet1.getCell(r, 4).value),
      lastName: text(sheet1.getCell(r, 5).value),
      firstName: text(sheet1.getCell(r, 6).value),
      dob: text(sheet1.getCell(r, 7).text || sheet1.getCell(r, 7).value),
      status: text(sheet1.getCell(r, 8).value),
      hk1Quiz: num(sheet1.getCell(r, 11).value),
      hk1p15: num(sheet1.getCell(r, 12).value),
      hk1Mid: num(sheet1.getCell(r, 13).value),
      hk1Exam: num(sheet1.getCell(r, 14).value),
      hk1Academic: num(sheet1.getCell(r, 15).value),
      hk1Mass: attendanceValue(sheet1.getCell(r, 16).value),
      hk1Class: attendanceValue(sheet1.getCell(r, 17).value),
      hk1Avg: num(sheet1.getCell(r, 21).value),
      hk2Quiz: num(sheet1.getCell(r, 22).value),
      hk2p15: num(sheet1.getCell(r, 23).value),
      hk2Mid: num(sheet1.getCell(r, 24).value),
      hk2Exam: num(sheet1.getCell(r, 25).value),
      hk2Academic: num(sheet1.getCell(r, 26).value),
      hk2Mass: attendanceValue(sheet1.getCell(r, 27).value),
      hk2Class: attendanceValue(sheet1.getCell(r, 28).value),
      hk2Avg: num(sheet1.getCell(r, 32).value),
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

export async function readWorkbook(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(arrayBuffer);
  } catch (e) {
    throw new AnalyzeError('File Excel không đúng định dạng.');
  }
  const sheet1 = wb.getWorksheet('Sheet1');
  if (!sheet1) throw new AnalyzeError('Không tìm thấy Sheet1 trong file Excel.');
  return { wb, sheet1 };
}

export async function processExcel(arrayBuffer) {
  const { wb, sheet1 } = await readWorkbook(arrayBuffer);
  const meta = extractMeta(sheet1);
  const students = buildStudents(sheet1);
  if (!students.length) throw new AnalyzeError('Không tìm thấy dữ liệu học sinh từ dòng 10 của Sheet1.');
  return { wb, students, meta };
}
