/**
 * Google Apps Script — ส่งเกียรติบัตรเข้าอีเมล
 * สำหรับเว็บ SSWH ศูนย์รวมการประเมินและใบประกาศ
 *
 * วิธีใช้:
 * 1) เปิด https://script.google.com  →  New project (แยกจากสคริปต์บันทึกแบบประเมิน)
 * 2) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้ลงไป
 * 3) กด Deploy → New deployment → เลือก type = "Web app"
 *      - Execute as: Me (บัญชีของคุณ)
 *      - Who has access: Anyone
 * 4) กด Deploy แล้วอนุญาตสิทธิ์ (Authorize) การส่งอีเมล
 * 5) คัดลอก "Web app URL" ที่ได้ (ลงท้ายด้วย /exec)
 * 6) นำไปวางแทนค่า CERT_MAIL_URL ในไฟล์ index.html
 *
 * หมายเหตุโควตาการส่งอีเมล/วัน:
 *   - บัญชี Gmail ทั่วไป: 100 ฉบับ/วัน
 *   - บัญชี Google Workspace: 1,500 ฉบับ/วัน
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'sendCertificate') {
      // แปลง base64 กลับเป็นไฟล์รูปภาพ JPG
      var bytes = Utilities.base64Decode(data.image);
      var fileName = (data.filename || 'certificate') + '.jpg';
      var blob = Utilities.newBlob(bytes, 'image/jpeg', fileName);

      MailApp.sendEmail({
        to: data.email,
        name: 'SSWH ศูนย์รวมการประเมิน',
        subject: 'เกียรติบัตร - ' + (data.projectName || 'การประเมิน'),
        htmlBody:
            'เรียน คุณ' + (data.name || '') + '<br><br>' +
            'ขอส่งเกียรติบัตรการเข้าร่วม<br>' +
            (data.projectName ? '<b>' + data.projectName + '</b><br>' : '') +
            (data.projectDate ? data.projectDate + '<br>' : '') +
            '<br>กรุณาดูเกียรติบัตรจากไฟล์แนบ<br><br>' +
            'ขอแสดงความนับถือ',
        attachments: [blob]
      });

      return ContentService
          .createTextOutput(JSON.stringify({ status: 'ok' }))
          .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
        .createTextOutput(JSON.stringify({ status: 'ignored' }))
        .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

// (ทางเลือก) ใช้ทดสอบว่า Web App ทำงานไหม เปิด URL /exec บนเบราว์เซอร์แล้วต้องเห็นข้อความนี้
function doGet() {
  return ContentService.createTextOutput('Certificate mail service is running.');
}
