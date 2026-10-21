const SHEET_ID = '10VsSDUgrGtgh5vAs_cWanBNmtKIdlbf1aH7_ZLixsdk';

// ===== คอลัมน์ข้อมูลทั่วไป (เหมือนกันทุกโครงการ) =====
const LEADING_HEADERS = [
  "Timestamp", "ชื่อ-สกุล (ใบประกาศ)", "ครั้งที่เข้าอบรม", "เพศ", "อายุ (ปี)", "ตำแหน่ง", "หน่วยงาน"
];
const TRAILING_HEADERS = ["ข้อเสนอแนะเพิ่มเติม"];

/**
 * รับข้อมูลแบบประเมินจากหน้าเว็บ แล้วบันทึกลง "แท็บคำตอบของโครงการนั้น"
 * payload ต้องมี: projectId, sheetTab (ชื่อแท็บคำตอบ เช่น ai2568_ตอบ) และข้อมูลคำตอบ
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const payload = JSON.parse(e.postData.contents);

    const projectId = (payload.projectId || '').toString().trim();
    // ชื่อแท็บคำตอบ: ใช้จาก payload ก่อน, ถ้าไม่มีให้ประกอบจาก projectId
    const tabName = (payload.sheetTab || (projectId ? projectId + '_ตอบ' : 'ประเมินโครงการ2')).toString().trim();

    // สร้างแท็บคำตอบถ้ายังไม่มี
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);

    // ถ้าแท็บว่าง -> สร้างหัวคอลัมน์จาก "แท็บคำถาม" ของโครงการ
    if (sheet.getLastRow() === 0) {
      const questions = getProjectQuestions(ss, projectId);
      const headers = LEADING_HEADERS.concat(questions).concat(TRAILING_HEADERS);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
    }

    // จับคู่ข้อมูลตามหัวคอลัมน์ที่มีอยู่จริง
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = headers.map(header => {
      const key = String(header).trim();
      if (key === 'Timestamp' || key === 'ประทับเวลา') return new Date();
      if (payload[key] !== undefined) return payload[key];
      const foundKey = Object.keys(payload).find(k => k.trim() === key);
      return foundKey !== undefined ? payload[foundKey] : '';
    });
    sheet.appendRow(rowData);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'บันทึกข้อมูลเรียบร้อย', tab: tabName }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// อ่านรายการคำถามจากแท็บ "<projectId>_ถาม" (คอลัมน์ชื่อ "คำถาม")
function getProjectQuestions(ss, projectId) {
  const questions = [];
  if (!projectId) return questions;
  const qSheet = ss.getSheetByName(projectId + '_ถาม');
  if (!qSheet || qSheet.getLastRow() < 2) return questions;

  const values = qSheet.getDataRange().getValues();
  const headerRow = values[0].map(h => String(h).replace(/\s+/g, '').toLowerCase());
  let qIdx = headerRow.findIndex(h => h.indexOf('คำถาม') !== -1 || h.indexOf('ข้อ') !== -1);
  if (qIdx === -1) qIdx = values[0].length - 1; // สำรอง: ใช้คอลัมน์สุดท้าย

  for (let i = 1; i < values.length; i++) {
    const q = String(values[i][qIdx] || '').trim();
    if (q) questions.push(q);
  }
  return questions;
}

// ดึงข้อมูลดิบออกจากแท็บ (GET) — ?sheet=<ชื่อแท็บ>
function doGet(e) {
  try {
    const sheetName = (e.parameter.sheet || '').trim();
    if (!sheetName) {
      return ContentService.createTextOutput('Survey multi-project service is running.');
    }
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, index) => { obj[header] = row[index]; });
      return obj;
    });
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
