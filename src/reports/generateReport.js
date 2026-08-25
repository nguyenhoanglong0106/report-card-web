import PizZip from 'pizzip';
import templateBuffer from '../../templates/report-card-template.docx';
import { round2, fmt } from '../utils/format.js';

// Parsed once when the module loads (counts against Workers' one-time startup
// budget, not the per-request CPU budget). The template only uses flat {TAG}
// placeholders — no {#loop} / {/loop} sections — so a direct string substitution
// on this cached XML replaces docxtemplater's full XML-DOM build, which alone
// cost ~35ms per document and blew the 10ms Workers CPU-time budget on its own.
const templateXml = new PizZip(templateBuffer).files['word/document.xml'].asText();

const XML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&apos;', '"': '&quot;' };
const escapeXml = (value) => String(value ?? '').replace(/[&<>'"]/g, (ch) => XML_ESCAPES[ch]);

function renderXml(data) {
  return templateXml.replace(/\{([A-Z0-9_]+)\}/g, (_match, tag) => escapeXml(data[tag]));
}

export function reportData(s, meta, form = {}) {
  return {
    NIEN_KHOA: form.schoolYear || meta.schoolYear || '',
    HOC_SINH: s.fullName,
    NGAY_SINH: s.dob,
    LOP: form.className || meta.className || '',
    GLV: form.glv || meta.glv || '',
    SDT: form.phone || '',
    NGAY_LIEN_HE: form.contactDate || '',
    KET_QUA: s.result,
    HK1_KHAO_BAI: fmt(s.hk1Quiz), HK1_15P: fmt(s.hk1p15), HK1_GIUA_KY: fmt(s.hk1Mid), HK1_HOC_KY: fmt(s.hk1Exam), HK1_TB: fmt(s.hk1Avg), HK1_XEP_LOAI: s.hk1Classify,
    HK1_DI_LE: fmt(s.hk1Mass), HK1_DI_HOC_GL: fmt(s.hk1Class), HK1_TB_CC: fmt(s.hk1Attendance),
    HK1_XEP_HANG: fmt(s.hk1AvgRank), HK1_XEP_HANG_GL: fmt(s.hk1AvgRank), HK1_XEP_HANG_CC: fmt(s.hk1AttendanceRank),
    HK2_KHAO_BAI: fmt(s.hk2Quiz), HK2_15P: fmt(s.hk2p15), HK2_GIUA_KY: fmt(s.hk2Mid), HK2_HOC_KY: fmt(s.hk2Exam), HK2_TB: fmt(s.hk2Avg), HK2_XEP_LOAI: s.hk2Classify,
    HK2_DI_LE: fmt(s.hk2Mass), HK2_DI_HOC_GL: fmt(s.hk2Class), HK2_TB_CC: fmt(s.hk2Attendance),
    HK2_XEP_HANG: fmt(s.hk2AvgRank), HK2_XEP_HANG_GL: fmt(s.hk2AvgRank), HK2_XEP_HANG_CC: fmt(s.hk2AttendanceRank),
    CN_KHAO_BAI: fmt(round2(((s.hk1Quiz || 0) + (s.hk2Quiz || 0)) / 2)), CN_15P: fmt(round2(((s.hk1p15 || 0) + (s.hk2p15 || 0)) / 2)),
    CN_GIUA_KY: fmt(round2(((s.hk1Mid || 0) + (s.hk2Mid || 0)) / 2)), CN_HOC_KY: fmt(round2(((s.hk1Exam || 0) + (s.hk2Exam || 0)) / 2)),
    CN_TB: fmt(s.yearAvg), CN_XEP_LOAI: s.yearClassify, CN_DI_LE: fmt((s.hk1Mass || 0) + (s.hk2Mass || 0)), CN_DI_HOC_GL: fmt((s.hk1Class || 0) + (s.hk2Class || 0)),
    CN_TB_CC: fmt(s.totalAttendance),
    CN_XEP_HANG: fmt(s.yearAvgRank), CN_XEP_HANG_GL: fmt(s.yearAvgRank), CN_XEP_HANG_CC: fmt(s.totalAttendanceRank),
  };
}

export function renderReportDocx(data) {
  const zip = new PizZip(templateBuffer);
  zip.file('word/document.xml', renderXml(data));
  // STORE (no compression) trades a larger file for much lower CPU cost than
  // DEFLATE — deflate alone cost ~20ms per document, the single biggest lever
  // after removing docxtemplater's DOM build.
  return zip.generate({ type: 'uint8array', compression: 'STORE' });
}
