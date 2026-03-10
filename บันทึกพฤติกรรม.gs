// ==========================================
// ไฟล์: บันทึกพฤติกรรม.gs (ระบบบันทึกคะแนน)
// ==========================================

const LOG_SHEET_NAME = 'บันทึกพฤติกรรม';

// ฟังก์ชันเปิดหน้าต่างบันทึกพฤติกรรม (เรียกจากเมนูหน้าบ้าน)
function openBehaviorLogModal() {
  const html = HtmlService.createTemplateFromFile('BehaviorLogUI')
    .evaluate()
    .setWidth(600) // ขนาดกว้างพอดีๆ
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'บันทึกพฤติกรรม / ตัดคะแนน');
}

// 1. ดึงข้อมูลนักเรียนทั้งหมด (สำหรับทำ Search Box)
function getAllStudentsForSearch() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('รายชื่อนักเรียน');
    const data = sheet.getDataRange().getDisplayValues();
    const students = [];
    
    // ข้าม Header (แถว 1)
    for (let i = 1; i < data.length; i++) {
      if(data[i][0]) { // ถ้ารหัสไม่ว่าง
        students.push({
          id: data[i][0],      // รหัส
          name: data[i][1],    // ชื่อ
          grade: data[i][3],   // ชั้น
          room: data[i][4]     // ห้อง
        });
      }
    }
    return { status: 'success', data: students };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// 2. ดึงเกณฑ์คะแนนทั้งหมด (สำหรับทำ Dropdown 2 ชั้น)
function getBehaviorOptions() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('เกณฑ์คะแนน');
    const data = sheet.getDataRange().getDisplayValues();
    const options = [];

    for (let i = 1; i < data.length; i++) {
      if(data[i][0] !== "") {
        options.push({
          level: data[i][0],    // ระดับ (0, U, 1, 2...)
          code: data[i][1],     // รหัส (0.1, U.1, 1.1...)
          detail: data[i][2],   // รายละเอียด
          points: data[i][3]    // คะแนน
        });
      }
    }
    return { status: 'success', data: options };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// 3. บันทึกข้อมูลลงชีต
function saveBehaviorLog(data) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET_NAME);
    if (!sheet) return { status: 'error', message: 'ไม่พบชีต "บันทึกพฤติกรรม"' };

    // เตรียมข้อมูลลงแถว
    const timestamp = new Date();
    const newRow = [
      timestamp,          // A: เวลา
      data.studentId,     // B: รหัส
      data.studentName,   // C: ชื่อ
      data.studentClass,  // D: ชั้น/ห้อง
      data.items,         // E: รายการ (String ยาวๆ)
      data.totalPoints,   // F: คะแนนรวม
      data.recorder || '' // G: ผู้บันทึก
    ];

    sheet.appendRow(newRow);
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}
