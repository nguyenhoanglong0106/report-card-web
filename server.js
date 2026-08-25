const path = require('path');
const fs = require('fs');
const express = require('express');
const multer = require('multer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const JSZip = require('jszip');
const { processExcel, reportData } = require('./src/gradeProcessor');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const templatePath = path.join(__dirname, 'templates', 'report-card-template.docx');
app.use(express.static(path.join(__dirname, 'public')));

function formData(req) {
  return { glv:req.body.glv||'', phone:req.body.phone||'', className:req.body.className||'', schoolYear:req.body.schoolYear||'', contactDate:req.body.contactDate||'' };
}
function safeName(s) { return String(s||'hoc-sinh').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').replace(/[^a-zA-Z0-9-_ ]/g,'').trim().replace(/\s+/g,'_'); }
function renderDoc(data) {
  const content=fs.readFileSync(templatePath,'binary');
  const zip=new PizZip(content);
  const doc=new Docxtemplater(zip,{paragraphLoop:true,linebreaks:true,nullGetter:()=>''});
  doc.render(data);
  return doc.getZip().generate({type:'nodebuffer',compression:'DEFLATE'});
}

app.post('/api/analyze', upload.single('file'), async (req,res)=>{
  try {
    if(!req.file) return res.status(400).json({error:'Vui lòng chọn file Excel.'});
    const {students,meta}=await processExcel(req.file.buffer);
    res.json({meta,count:students.length,students:students.map((s,i)=>({index:i,stt:s.stt,code:s.code,name:s.fullName,dob:s.dob,hk1:s.hk1Avg,hk2:s.hk2Avg,year:s.yearAvg,rank:s.yearAvgRank,classify:s.yearClassify,result:s.result}))});
  } catch(e){ res.status(400).json({error:e.message}); }
});

app.post('/api/export-excel', upload.single('file'), async (req,res)=>{
  try {
    const {wb}=await processExcel(req.file.buffer);
    const buffer=await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="KetQuaHocTap_DaTaoSheet2.xlsx"');
    res.send(Buffer.from(buffer));
  } catch(e){ res.status(400).json({error:e.message}); }
});

app.post('/api/report/:index', upload.single('file'), async (req,res)=>{
  try {
    const {students,meta}=await processExcel(req.file.buffer);
    const idx=Number(req.params.index); const s=students[idx];
    if(!s) return res.status(404).json({error:'Không tìm thấy học sinh.'});
    const out=renderDoc(reportData(s,meta,formData(req)));
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="PhieuBaoDiem_${safeName(s.fullName)}.docx"`);
    res.send(out);
  } catch(e){ res.status(400).json({error:e.message}); }
});

app.post('/api/reports-all', upload.single('file'), async (req,res)=>{
  try {
    const {students,meta}=await processExcel(req.file.buffer); const zip=new JSZip();
    const form=formData(req);
    students.forEach((s,i)=>zip.file(`${String(i+1).padStart(2,'0')}_PhieuBaoDiem_${safeName(s.fullName)}.docx`,renderDoc(reportData(s,meta,form))));
    const out=await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
    res.setHeader('Content-Type','application/zip'); res.setHeader('Content-Disposition','attachment; filename="PhieuBaoDiem_CaLop.zip"'); res.send(out);
  } catch(e){ res.status(400).json({error:e.message}); }
});

const port=process.env.PORT||3000;
app.listen(port,()=>console.log(`Web đang chạy: http://localhost:${port}`));
