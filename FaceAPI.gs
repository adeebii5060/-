// =========================================================
// ไฟล์สำหรับจัดการ API สแกนใบหน้าโดยเฉพาะ (ทำหน้าที่เป็น Backend)
// =========================================================

// ตัวแปรตั้งค่าเวลา (แก้ไขได้ตามต้องการ)
const CUTOFF_LATE_TIME = "08:00:00"; // หลังเวลานี้ถือว่าสายตอนเช้า (ได้ 0 คะแนน)
const CUTOFF_NOON_TIME = "12:00:00"; // หลังเวลานี้ถือว่าเป็นการสแกน "ออก"

function processFacePostRequest(data) {
  const action = data.action;
  
  if (action === 'registerUser') {
    return registerFaceData(data.name, data.faceDescriptor); // data.name ในที่นี้เราจะใช้ รหัสนักเรียน
  } else if (action === 'logAttendance') {
    return logFaceAttendance(data.name); // data.name คือ รหัสนักเรียน ที่แสกนเจอ
  }
  
  return { error: 'Unknown action' };
}

// 1. ดึงข้อมูลใบหน้าไปให้หน้าเว็บ
function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('รายชื่อนักเรียน');
  if (!sheet) return { error: 'ไม่พบชีต รายชื่อนักเรียน' };

  const data = sheet.getDataRange().getValues();
  const faces = [];

  // เริ่ม loop จากแถว 2 (ข้าม Header)
  for (let i = 1; i < data.length; i++) {
    const studentId = data[i][0]; // คอลัมน์ A: รหัสนักเรียน
    const faceDataStr = data[i][6]; // คอลัมน์ G: ข้อมูลใบหน้า

    if (studentId && faceDataStr) {
      try {
        const descriptor = JSON.parse(faceDataStr);
        faces.push({
          name: studentId.toString(), // ส่งรหัสนักเรียนไปเป็นชื่อ label ใบหน้า
          descriptors: [descriptor]
        });
      } catch (e) {
        // ข้ามหากข้อมูลใบหน้าไม่ใช่ JSON ที่ถูกต้อง
      }
    }
  }

  return faces;
}

// 2. บันทึกใบหน้าใหม่ลงชีต "รายชื่อนักเรียน"
function registerFaceData(studentId, faceDescriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('รายชื่อนักเรียน');
  if (!sheet) return { success: false, message: 'ไม่พบชีต รายชื่อนักเรียน' };

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;

  // ค้นหานักเรียนจาก รหัสนักเรียน (คอลัมน์ A)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === studentId.toString()) {
      foundRow = i + 1; // +1 เพราะ index เริ่มที่ 0 แต่แถวชีตเริ่มที่ 1
      break;
    }
  }

  if (foundRow > 0) {
    // บันทึก JSON ลงคอลัมน์ G (7)
    sheet.getRange(foundRow, 7).setValue(JSON.stringify(faceDescriptor));
    return { success: true, message: 'บันทึกใบหน้า รหัส ' + studentId + ' สำเร็จ' };
  } else {
    return { success: false, message: 'ไม่พบรหัสนักเรียน ' + studentId + ' ในระบบ กรุณาตรวจสอบ' };
  }
}

// 3. บันทึกการมาเรียนและคิดคะแนน
function logFaceAttendance(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 3.1 ดึงข้อมูลนักเรียนก่อน
  const studentSheet = ss.getSheetByName('รายชื่อนักเรียน');
  const studentData = studentSheet.getDataRange().getValues();
  let studentName = "ไม่ทราบชื่อ";
  let studentRoom = "-";
  
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][0].toString() === studentId.toString()) {
      studentName = studentData[i][1]; // ชื่อ-สกุล
      studentRoom = studentData[i][4]; // ห้อง (อ้างอิงคอลัมน์ E)
      break;
    }
  }

  // 3.2 จัดการชีตเวลาเรียน
  const logSheet = ss.getSheetByName('ฐานข้อมูลการมาเรียน');
  if (!logSheet) return { success: false, message: 'ไม่พบชีต ฐานข้อมูลการมาเรียน' };

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  const isMorning = timeStr < CUTOFF_NOON_TIME;
  const isLate = timeStr > CUTOFF_LATE_TIME;

  const logs = logSheet.getDataRange().getValues();
  let foundRow = -1;
  let currentInScore = 0;

  // 3.3 ค้นหาว่าวันนี้ รหัสนี้ เคยสแกนหรือยัง? (ค้นหาจากล่างขึ้นบน)
  for (let i = logs.length - 1; i >= 1; i--) {
    if (logs[i][0] === dateStr && logs[i][1].toString() === studentId.toString()) {
      foundRow = i + 1;
      currentInScore = Number(logs[i][5]) || 0; // คอลัมน์ F (คะแนนเข้า)
      break;
    }
  }

  if (isMorning) {
    // === สแกนเข้า (เช้า) ===
    let inScore = isLate ? 0 : 3;
    let status = isLate ? "สาย" : "มาเรียน";

    if (foundRow > 0) {
      return { success: true, message: `คุณ ${studentName} ได้สแกนเข้าเช้าไปแล้ว` };
    } else {
      // แถวใหม่: วันที่(A) / รหัส(B) / ชื่อ(C) / ห้อง(D) / เวลาเข้า(E) / คะแนนเข้า(F) / เวลาออก(G) / คะแนนออก(H) / รวม(I) / สถานะ(J)
      logSheet.appendRow([
        dateStr,          
        studentId,        
        studentName,      
        studentRoom,      
        timeStr,          
        inScore,          
        "",               
        "",               
        inScore,          
        status            
      ]);
      return { success: true, message: `สแกนเข้าสำเร็จ [${status}] +${inScore} คะแนน` };
    }

  } else {
    // === สแกนออก (บ่าย/เย็น) ===
    let outScore = 3; 

    if (foundRow > 0) {
      // เคยสแกนเข้าแล้ว
      let totalScore = currentInScore + outScore;
      
      logSheet.getRange(foundRow, 7).setValue(timeStr);   // G: เวลาออก
      logSheet.getRange(foundRow, 8).setValue(outScore);  // H: คะแนนออก
      logSheet.getRange(foundRow, 9).setValue(totalScore); // I: คะแนนรวม
      
      return { success: true, message: `สแกนออกสำเร็จ รวมได้ ${totalScore} คะแนน` };
    } else {
      // ขาดตอนเช้า มาสแกนออกอย่างเดียว
      let inScore = 0;
      let totalScore = inScore + outScore;
      
      logSheet.appendRow([
        dateStr,
        studentId,
        studentName,
        studentRoom,
        "-",              
        inScore,          
        timeStr,          
        outScore,         
        totalScore,       
        "ขาดตอนเช้า"
      ]);
      return { success: true, message: `สแกนออกสำเร็จ (ไม่พบข้อมูลเข้าเช้า) รวม ${totalScore} คะแนน` };
    }
  }
}
