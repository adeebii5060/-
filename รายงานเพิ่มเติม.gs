// ==========================================
// ไฟล์: รายงานเพิ่มเติม.gs
// ==========================================

const REPORT_SHEET_NAME = 'รายงานเพิ่มเติม';

// 1. บันทึกรายงานลงชีต
function saveAdditionalReport(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(REPORT_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(REPORT_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'หัวข้อ', 'รายละเอียด', 'หมายเหตุ']);
    }

    sheet.appendRow([new Date(), data.title, data.detail, data.note]);
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// 2. ดึงข้อมูลรายงานล่าสุด 10 รายการมาโชว์ในหน้าเว็บ
function getRecentReports() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REPORT_SHEET_NAME);
    if (!sheet) return { status: 'success', data: [] };
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { status: 'success', data: [] };

    const results = [];
    // วนลูปย้อนหลังจากล่างขึ้นบน เอาแค่ 10 อันล่าสุด
    for (let i = data.length - 1; i >= Math.max(1, data.length - 10); i--) {
      results.push({
        date: data[i][0],
        title: data[i][1],
        detail: data[i][2],
        note: data[i][3]
      });
    }
    return { status: 'success', data: results };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}
