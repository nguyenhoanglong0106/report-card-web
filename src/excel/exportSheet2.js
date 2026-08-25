function styleSheet2(ws, count) {
  ws.views = [{ state: 'frozen', ySplit: 9 }];
  const widths = [6, 12, 14, 22, 12, 13, 14, 9, 9, 9, 9, 10, 11, 13, 11, 12, 9, 12, 9, 9, 9, 9, 10, 11, 13, 11, 12, 9, 12, 11, 10, 10, 10, 11, 11];
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  for (let r = 6; r <= 9 + count; r++) {
    ws.getRow(r).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  const rows = ws.getRows(6, 4);
  rows.forEach((row) =>
    row.eachCell({ includeEmpty: true }, (c) => {
      c.font = { name: 'Calibri', size: 10, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    })
  );
  for (let r = 10; r < 10 + count; r++) {
    ws.getRow(r).eachCell({ includeEmpty: true }, (c) => {
      c.font = { name: 'Calibri', size: 10 };
      c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws.getCell(r, 4).alignment = { horizontal: 'left', vertical: 'middle' };
  }
}

export function createSheet2(wb, students, meta) {
  const old = wb.getWorksheet('Sheet2');
  if (old) wb.removeWorksheet(old.id);
  const ws = wb.addWorksheet('Sheet2');
  ws.getCell('A1').value = meta.parish || 'GIÁO XỨ BẮC THẦN';
  ws.getCell('A2').value = meta.glv.split(' - ')[0] || '';
  ws.getCell('A3').value = meta.glv.split(' - ')[1] || '';
  ws.getCell('A4').value = `KẾT QUẢ HỌC TẬP CHI TIẾT\n ${meta.className || ''} - ${meta.schoolYear || ''}`;

  const headers = {
    A6: 'Stt', B6: 'Mã số', C6: 'Tên Thánh', D6: 'Họ', E6: 'Tên', F6: 'Ngày sinh', G6: 'Tình trạng',
    H6: 'Học kỳ I (Quy định 100% học lực và 0% C. Cần)', Q6: 'Xếp hạng', R6: 'Xếp loại',
    S6: 'Học kỳ II (Quy định 100% học lực và 0% C. Cần)', AB6: 'Xếp hạng', AC6: 'Xếp loại',
    AD6: 'TB C.Cần', AE6: 'TB H.Lực', AF6: 'Tổng TB', AG6: 'Xếp loại', AH6: 'Xếp hạng', AI6: 'Kết quả',
    G7: 'Tình trạng', H7: 'Học lực', M7: 'Chuyên cần', S7: 'Học lực', X7: 'Chuyên cần',
    H8: 'Đầu giờ', I8: '15 phút', J8: '45 phút', K8: 'Thi', L8: 'Tb. Học lực', M8: 'Số lần dự T.Lễ', N8: 'Số lần đi học G.Lý', O8: 'Tb. Chuyên cần', P8: 'Trung bình học kỳ',
    S8: 'Đầu giờ', T8: '15 phút', U8: '45 phút', V8: 'Thi', W8: 'Tb. Học lực', X8: 'Số lần dự T.Lễ', Y8: 'Số lần đi học G.Lý', Z8: 'Tb. Chuyên cần', AA8: 'Trung bình học kỳ',
    D9: 'Hệ số', H9: 1, I9: 1, J9: 2, K9: 3, S9: 1, T9: 1, U9: 2, V9: 3,
  };
  Object.entries(headers).forEach(([a, v]) => (ws.getCell(a).value = v));
  ws.mergeCells('H6:P6');
  ws.mergeCells('S6:AA6');
  ws.mergeCells('H7:L7');
  ws.mergeCells('M7:P7');
  ws.mergeCells('S7:W7');
  ws.mergeCells('X7:AA7');

  students.forEach((s, idx) => {
    const r = 10 + idx;
    const values = [s.stt, s.code, s.saintName, s.lastName, s.firstName, s.dob, s.status, s.hk1Quiz, s.hk1p15, s.hk1Mid, s.hk1Exam, s.hk1Academic, s.hk1Mass, s.hk1Class, null, s.hk1Avg, null, null, s.hk2Quiz, s.hk2p15, s.hk2Mid, s.hk2Exam, s.hk2Academic, s.hk2Mass, s.hk2Class, null, s.hk2Avg, null, null, null, null, null, null, null, null];
    values.forEach((v, i) => (ws.getCell(r, i + 1).value = v));
    ws.getCell(r, 15).value = { formula: `M${r}+N${r}`, result: s.hk1Attendance };
    ws.getCell(r, 17).value = { formula: `RANK(P${r},$P$10:$P$${9 + students.length},0)`, result: s.hk1AvgRank };
    ws.getCell(r, 18).value = { formula: `IF(P${r}>=9,"Xuất Sắc",IF(P${r}>=8,"Giỏi",IF(P${r}>=6.5,"Khá",IF(P${r}>=5,"Trung Bình","Yếu"))))`, result: s.hk1Classify };
    ws.getCell(r, 26).value = { formula: `X${r}+Y${r}`, result: s.hk2Attendance };
    ws.getCell(r, 28).value = { formula: `RANK(AA${r},$AA$10:$AA$${9 + students.length},0)`, result: s.hk2AvgRank };
    ws.getCell(r, 29).value = { formula: `IF(AA${r}>=9,"Xuất Sắc",IF(AA${r}>=8,"Giỏi",IF(AA${r}>=6.5,"Khá",IF(AA${r}>=5,"Trung Bình","Yếu"))))`, result: s.hk2Classify };
    ws.getCell(r, 30).value = { formula: `O${r}+Z${r}`, result: s.totalAttendance };
    ws.getCell(r, 31).value = { formula: `ROUND((P${r}+AA${r})/2,2)`, result: s.yearAcademic };
    ws.getCell(r, 32).value = { formula: `AE${r}`, result: s.yearAvg };
    ws.getCell(r, 33).value = { formula: `RANK(AF${r},$AF$10:$AF$${9 + students.length},0)`, result: s.yearAvgRank };
    ws.getCell(r, 34).value = { formula: `IF(AF${r}>=9,"Xuất Sắc",IF(AF${r}>=8,"Giỏi",IF(AF${r}>=6.5,"Khá",IF(AF${r}>=5,"Trung Bình","Yếu"))))`, result: s.yearClassify };
    ws.getCell(r, 35).value = { formula: `IF(AE${r}<3.5,"Ở lại","Lên Lớp")`, result: s.result };
  });
  styleSheet2(ws, students.length);
  return ws;
}
