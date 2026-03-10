/**
 * ============================================================
 * ระบบจัดการฝ่ายปกครองอัจฉริยะ (Smart Discipline System)
 * โรงเรียนศาสนูปถัมภ์ (Sasanupatham School)
 * ฉบับสมบูรณ์: รวมทุกฟีเจอร์ (Student, Behavior, OCR, Confiscated, Tarbiya)
 * ============================================================
 */

const SHEET_ID = '1riA4s_wXhxn88xjdOdsNZ_taCz_CNOXcFJcukq6hQFM'; // <--- แก้ไข ID ของคุณตรงนี้

/**
 * ส่วนที่ 1: ระบบหลักและการแสดงผล (Core System)
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ฝ่ายปกครอง')
    .addItem('เปิดระบบจัดการ', 'openDashboard')
    .addToUi();
}

function openDashboard() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate()
    .setWidth(1200)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'ระบบฝ่ายปกครอง Sasanupatham');
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบฝ่ายปกครอง | Sasanupatham')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ส่วนที่ 2: ระบบตรวจสอบประวัติรายบุคคล (Student 360 Profile)
 */
function getStudentFullProfile(studentId) {
  if (!studentId) return { status: 'error', message: 'กรุณาระบุรหัสนักเรียน' };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    studentId = studentId.toString().trim();

    // 1. ดึงข้อมูลพื้นฐาน (Profile)
    const studentSheet = ss.getSheetByName('รายชื่อนักเรียน');
    const studentData = studentSheet.getDataRange().getDisplayValues();
    const pData = studentData.find(r => r[0].toString() === studentId);
    
    if (!pData) return { status: 'error', message: 'ไม่พบนักเรียนรหัส: ' + studentId };
    
    const profile = {
      id: pData[0],
      name: pData[1],
      gender: pData[2],
      grade: pData[3],
      room: pData[4]
    };

    // 2. ดึงข้อมูลพฤติกรรม (Behavior Logs)
    const bSheet = ss.getSheetByName('บันทึกพฤติกรรม');
    const bData = bSheet.getDataRange().getDisplayValues();
    let behaviorLogs = [];
    let behaviorScoreSum = 0;

    bData.slice(1).forEach(row => {
      if (row[1].toString() === studentId) {
        const pts = parseInt(row[5]) || 0;
        behaviorScoreSum += pts;
        behaviorLogs.push({
          date: row[0],
          items: row[4],
          points: pts,
          recorder: row[6]
        });
      }
    });

    // 3. ดึงสถิติการมาเรียน (Attendance Stats)
    const aSheet = ss.getSheetByName('ฐานข้อมูลการมาเรียน');
    let stats = { normal: 0, late: 0, absent: 0, leave: 0, missingScan: 0 };
    let attendLogs = [];
    let attendScoreSum = 0;

    if (aSheet) {
      const aData = aSheet.getDataRange().getDisplayValues();
      aData.slice(1).forEach(row => {
        if (row[1].toString() === studentId) {
          const status = row[9];
          const scoreIn = parseInt(row[5]) || 0;
          const scoreOut = parseInt(row[7]) || 0;
          attendScoreSum += (scoreIn + scoreOut);

          if (status === 'ปกติ') stats.normal++;
          else if (status === 'สาย') stats.late++;
          else if (status === 'ขาด') stats.absent++;
          else if (status === 'ลา') stats.leave++;
          
          attendLogs.push({
            date: row[0],
            inTime: row[4],
            outTime: row[6],
            status: status,
            score: (scoreIn + scoreOut)
          });
        }
      });
    }

    return {
      status: 'success',
      profile: profile,
      behavior: behaviorLogs.reverse(),
      attendance: attendLogs.reverse(),
      stats: stats,
      totalPoints: attendScoreSum + behaviorScoreSum
    };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * ส่วนที่ 3: ระบบส่งต่อฝ่ายตัรบียะห์ (Tarbiya Referral System)
 */
function submitToTarbiya(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('การส่งต่อฝ่ายตัรบียะห์');
    if (!sheet) {
      sheet = ss.insertSheet('การส่งต่อฝ่ายตัรบียะห์');
      sheet.appendRow(['วันที่ส่งต่อ','รหัสนักเรียน','ชื่อ-นามสกุล','สรุปพฤติกรรม','บันทึกจากครูปกครอง','สถานะ','หมายเหตุ']);
    }
    
    sheet.appendRow([
      new Date(), 
      d.studentId, 
      d.studentName, 
      d.summary, 
      d.note, 
      "รอการขัดเกลา", 
      ""
    ]);
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * ส่วนที่ 4: ระบบจัดการรายชื่อนักเรียน (Student Management)
 */
function getStudentsByClass(grade, room) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายชื่อนักเรียน');
    const data = sheet.getDataRange().getDisplayValues();
    const filtered = data.slice(1)
      .filter(r => r[3] === grade && r[4] === room.toString())
      .map(r => ({
        rowIdx: data.indexOf(r) + 1,
        id: r[0],
        name: r[1],
        gender: r[2],
        grade: r[3],
        room: r[4]
      }));
    return { status: 'success', data: filtered };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function addStudent(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายชื่อนักเรียน');
    sheet.appendRow([d.id, d.name, d.gender, d.grade, d.room, '-']);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function updateStudent(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายชื่อนักเรียน');
    sheet.getRange(d.rowIdx, 1, 1, 5).setValues([[d.id, d.name, d.gender, d.grade, d.room]]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function deleteStudent(idx) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายชื่อนักเรียน');
    sheet.deleteRow(idx);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function getAllStudentsForSearch() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const data = ss.getSheetByName('รายชื่อนักเรียน').getDataRange().getDisplayValues();
    const list = data.slice(1).map(r => ({ id: r[0], name: r[1], grade: r[3], room: r[4] }));
    return { status: 'success', data: list };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

/**
 * ส่วนที่ 5: ระบบจัดการพฤติกรรม (Behavior Management)
 */
function getBehaviorOptions() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const data = ss.getSheetByName('เกณฑ์คะแนน').getDataRange().getDisplayValues();
    const options = data.slice(1).map(r => ({ level: r[0], code: r[1], detail: r[2], points: r[3] }));
    return { status: 'success', data: options };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function saveBehaviorLog(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('บันทึกพฤติกรรม');
    sheet.appendRow([new Date(), d.studentId, d.studentName, d.studentClass, d.items, d.totalPoints, "ครูฝ่ายปกครอง"]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function getBehaviorCriteria() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const data = ss.getSheetByName('เกณฑ์คะแนน').getDataRange().getDisplayValues();
    const criteria = data.slice(1).map((r, i) => ({ rowIdx: i + 2, level: r[0], code: r[1], detail: r[2], points: r[3] }));
    return { status: 'success', data: criteria };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function addBehavior(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ss.getSheetByName('เกณฑ์คะแนน').appendRow([d.level, d.code, d.detail, d.points]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function updateBehavior(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ss.getSheetByName('เกณฑ์คะแนน').getRange(d.rowIdx, 1, 1, 4).setValues([[d.level, d.code, d.detail, d.points]]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function deleteBehavior(idx) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ss.getSheetByName('เกณฑ์คะแนน').deleteRow(idx);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

/**
 * ส่วนที่ 6: ระบบจัดการของกลาง (Confiscated Management)
 */
function getConfiscatedFolder() {
  const folderName = "Confiscated_Items_Photos";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function saveConfiscatedItem(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('รายการของกลาง');
    if (!sheet) {
      sheet = ss.insertSheet('รายการของกลาง');
      sheet.appendRow(['วันที่ยึด','รหัสนักเรียน','ชื่อนักเรียน','รายการของกลาง','รูปภาพ','กำหนดคืน','วันที่คืนจริง','สถานะ','ครูผู้ยึด','หมายเหตุ']);
    }

    // หาชื่อนักเรียนจากรหัสอัตโนมัติ
    const studentSheet = ss.getSheetByName('รายชื่อนักเรียน');
    const sData = studentSheet.getDataRange().getDisplayValues();
    const student = sData.find(r => r[0].toString() === data.studentId.toString());
    const studentName = student ? student[1] : "ไม่ทราบชื่อ";

    let imageUrl = "";
    if (data.imageFile && data.imageFile.includes("base64")) {
      const folder = getConfiscatedFolder();
      const bytes = Utilities.base64Decode(data.imageFile.split(',')[1]);
      const blob = Utilities.newBlob(bytes, "image/jpeg", "item_" + data.studentId + "_" + new Date().getTime() + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(), 
      data.studentId, 
      studentName, 
      data.itemName, 
      imageUrl, 
      data.returnDueDate, 
      "-", 
      "รอคืน", 
      "ครูฝ่ายปกครอง", 
      ""
    ]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function getConfiscatedItems() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายการของกลาง');
    if (!sheet) return { status: 'success', data: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const result = data.slice(1).map((r, i) => ({
      rowIdx: i + 2,
      dateSeized: r[0],
      studentId: r[1],
      studentName: r[2],
      itemName: r[3],
      imageUrl: r[4],
      returnDueDate: r[5],
      actualReturnDate: r[6],
      status: r[7]
    })).reverse();
    return { status: 'success', data: result };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function markAsReturned(rowIdx) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายการของกลาง');
    sheet.getRange(rowIdx, 7, 1, 2).setValues([[new Date(), "คืนแล้ว"]]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

/**
 * ส่วนที่ 7: ระบบรายงานและสถิติการมาเรียน (OCR & Report)
 */
function saveAdditionalReport(d) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายงานเพิ่มเติม');
    sheet.appendRow([new Date(), d.title, d.detail, d.note]);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function getRecentReports() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายงานเพิ่มเติม');
    if (!sheet) return { status: 'success', data: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const reports = data.slice(-10).reverse().map(r => ({ date: r[0], title: r[1], detail: r[2], note: r[3] }));
    return { status: 'success', data: reports };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

// ฟังก์ชัน OCR PDF
function uploadAttendancePDF(base64Data, fileName) {
  try {
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, 'application/pdf', fileName);
    const resource = { title: fileName, mimeType: 'application/pdf' };
    const file = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'th' });
    const fullText = DocumentApp.openById(file.id).getBody().getText();
    Drive.Files.remove(file.id);
    return robustAttendanceParser(fullText);
  } catch (e) { return { status: 'error', message: 'OCR Error: ' + e.toString() }; }
}

function robustAttendanceParser(text) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('ฐานข้อมูลการมาเรียน') || ss.insertSheet('ฐานข้อมูลการมาเรียน');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['วันที่','รหัส','ชื่อ','ห้อง','เข้า','คะแนนเข้า','ออก','คะแนนออก','รวม','สถานะ']);
    }

    const dateRegex = /\d{1,2}\s+(ธ\.ค\.|ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.)\s+\d{2}/g;
    let foundDates = text.match(dateRegex) || ["N/A"];
    foundDates = [...new Set(foundDates)].slice(0, 5);

    const studentPattern = /([S\d]{4,10})\s+([ก-๙\s]+?)\s+((?:มัธยม|ซานาวีย์)\s+\d\/\d)/g;
    let matches = [];
    let match;
    while ((match = studentPattern.exec(text)) !== null) {
      matches.push({ id: match[1], name: match[2].trim(), room: match[3].trim(), pos: match.index, raw: match[0] });
    }

    let uploadRows = [];
    matches.forEach((curr, i) => {
      const nextPos = matches[i+1] ? matches[i+1].pos : text.length;
      const segment = text.substring(curr.pos + curr.raw.length, nextPos);
      const times = segment.match(/(\d{1,2}:\d{2}|ข\.|ล\.|-)/g) || [];
      
      for (let d = 0; d < 5; d++) {
        const inT = times[d*2] || "-";
        const outT = times[d*2+1] || "-";
        let status = "ปกติ", inS = 0, outS = 0;

        if (inT === "ข.") status = "ขาด";
        else if (inT === "ล.") status = "ลา";
        else if (inT.includes(":")) {
          const [h, m] = inT.split(':').map(Number);
          if (h > 8 || (h === 8 && m > 15)) status = "สาย"; 
          else inS = 3;
        }
        
        if (outT.includes(":") && outT !== "-" && status !== "ขาด" && status !== "ลา") {
          outS = 3;
        }
        
        uploadRows.push([
          foundDates[d] || "N/A", 
          curr.id, curr.name, curr.room, 
          inT, inS, outT, outS, (inS+outS), status
        ]);
      }
    });

    if (uploadRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, uploadRows.length, 10).setValues(uploadRows);
    }
    return { status: 'success', message: 'ประมวลผลสำเร็จ ' + matches.length + ' รายการ' };
  } catch (e) {
    return { status: 'error', message: 'Parser Error: ' + e.toString() };
  }
}
