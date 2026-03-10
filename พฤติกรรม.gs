// ==========================================
// ไฟล์: พฤติกรรม.gs (ระบบจัดการเกณฑ์คะแนน)
// ==========================================

const CRITERIA_SHEET_NAME = 'เกณฑ์คะแนน';

function getCriteriaSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(CRITERIA_SHEET_NAME);
}

// เช็ครหัสข้อซ้ำ
function checkDuplicateBehaviorCode(code, excludeRowIdx = -1) {
  const data = getCriteriaSheet().getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(code).trim() && (i + 1) !== excludeRowIdx) {
      return true; 
    }
  }
  return false;
}

// โหลดข้อมูลเกณฑ์คะแนน
function getBehaviorCriteria() {
  try {
    const sheet = getCriteriaSheet();
    if (!sheet) return { status: 'error', message: 'ไม่พบชีตชื่อ เกณฑ์คะแนน' };

    const data = sheet.getDataRange().getDisplayValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] !== "") { // เช็คว่ามีข้อมูลระดับไหม
        result.push({
          rowIdx: i + 1,
          level: data[i][0],    // คอลัมน์ A (ระดับ)
          code: data[i][1],     // คอลัมน์ B (รหัสข้อ)
          detail: data[i][2],   // คอลัมน์ C (รายละเอียด)
          points: data[i][3]    // คอลัมน์ D (คะแนนที่หัก)
        });
      }
    }
    return { status: 'success', data: result };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// เพิ่มเกณฑ์ใหม่
function addBehavior(data) {
  try {
    if (checkDuplicateBehaviorCode(data.code)) return { status: 'error', message: 'รหัสข้อนี้มีในระบบแล้ว!' };
    getCriteriaSheet().appendRow([data.level, data.code, data.detail, data.points]);
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// แก้ไขเกณฑ์
function updateBehavior(data) {
  try {
    if (checkDuplicateBehaviorCode(data.code, data.rowIdx)) return { status: 'error', message: 'รหัสข้อนี้ไปซ้ำกับข้ออื่น!' };
    const sheet = getCriteriaSheet();
    sheet.getRange(data.rowIdx, 1, 1, 4).setValues([[data.level, data.code, data.detail, data.points]]);
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// ลบเกณฑ์
function deleteBehavior(rowIdx) {
  try {
    getCriteriaSheet().deleteRow(rowIdx);
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}
