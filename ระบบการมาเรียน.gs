/**
 * ==========================================
 * ไฟล์: ระบบการมาเรียน.gs (ฉบับสมบูรณ์ 100%)
 * นโยบายคะแนน: 
 * - มาสาย (>08:15): เข้า 0 / ออก 3 (ถ้ามีเวลาออก)
 * - มาปกติ (<=08:15): เข้า 3 / ออก 3
 * - ขาด/ลา: เข้า 0 / ออก 0
 * ==========================================
 */

function uploadAttendancePDF(base64Data, fileName) {
  try {
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, 'application/pdf', fileName);

    // 1. ทำ OCR ผ่าน Google Drive (ต้องเปิดบริการ Drive API v2 ในเมนู Services แถบซ้ายมือ)
    const resource = { title: fileName, mimeType: 'application/pdf' };
    const file = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'th' });
    
    // 2. ดึงข้อความดิบ
    const doc = DocumentApp.openById(file.id);
    const fullText = doc.getBody().getText();
    
    // 3. ลบไฟล์ชั่วคราวทิ้งทันที
    Drive.Files.remove(file.id);

    // 4. ส่งไปประมวลผล
    return robustAttendanceParser(fullText);
    
  } catch (e) {
    return { status: 'error', message: 'ข้อผิดพลาด OCR: ' + e.toString() };
  }
}

function robustAttendanceParser(text) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('ฐานข้อมูลการมาเรียน');
  
  if (!sheet) {
    sheet = ss.insertSheet('ฐานข้อมูลการมาเรียน');
    sheet.appendRow(['วันที่', 'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ห้อง', 'เวลาเข้า', 'คะแนนเข้า', 'เวลาออก', 'คะแนนออก', 'คะแนนรวมวันนั้น', 'สถานะ']);
  }

  // --- ระบบเช็คข้อมูลซ้ำเพื่อป้องกันการบันทึกซ้ำ (วันที่_รหัส) ---
  const existingValues = sheet.getDataRange().getValues();
  const duplicateSet = new Set();
  for (let i = 1; i < existingValues.length; i++) {
    duplicateSet.add(existingValues[i][0] + "_" + existingValues[i][1]); 
  }

  // --- ค้นหาวันที่จาก PDF (เช่น 22 ธ.ค. 68) ---
  const dateRegex = /\d{1,2}\s+(ธ\.ค\.|ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.)\s+\d{2}/g;
  let foundDates = text.match(dateRegex) || ["ไม่ทราบวันที่"];
  foundDates = [...new Set(foundDates)].slice(0, 5); // จำกัด 5 วัน

  // --- ค้นหานักเรียน (รหัส | ชื่อ | ชั้นเรียน) ---
  const studentPattern = /([S\d]{4,10})\s+([ก-๙\s]+?)\s+((?:มัธยมศึกษาปีที่|ซานาวีย์)\s+\d\/\d)/g;
  let matches = [];
  let match;
  while ((match = studentPattern.exec(text)) !== null) {
    matches.push({ id: match[1], name: match[2].trim(), room: match[3].trim(), pos: match.index, raw: match[0] });
  }

  if (matches.length === 0) return { status: 'error', message: 'ไม่พบรายชื่อนักเรียนในไฟล์' };

  let uploadRows = [];
  let duplicateCount = 0;

  // --- ประมวลผลข้อมูลรายคน ---
  for (let i = 0; i < matches.length; i++) {
    const current = matches[i];
    const nextPos = (i + 1 < matches.length) ? matches[i + 1].pos : text.length;
    const segment = text.substring(current.pos + current.raw.length, nextPos);
    
    // ดึงเวลาหรือสถานะ (07:32, ข., ล., -)
    const times = segment.match(/(\d{1,2}:\d{2}|ข\.|ล\.|-)/g) || [];

    for (let d = 0; d < 5; d++) {
      const dateKey = foundDates[d] + "_" + current.id;
      
      if (duplicateSet.has(dateKey)) {
        duplicateCount++;
        continue;
      }

      const inTime = times[d * 2] || "-";
      const outTime = times[d * 2 + 1] || "-";

      let status = "ปกติ";
      let inScore = 0;
      let outScore = 0;

      // 1. ตัดสิน "ขาเข้า" และ "สถานะ"
      if (inTime === "ข.") {
        status = "ขาด";
        inScore = 0;
      } else if (inTime === "ล.") {
        status = "ลา";
        inScore = 0;
      } else if (inTime.includes(":")) {
        // แยกชั่วโมงและนาทีเพื่อความแม่นยำ
        const timeParts = inTime.split(':');
        const hour = parseInt(timeParts[0], 10);
        const minute = parseInt(timeParts[1], 10);

        // เช็คเกณฑ์สาย (8:15)
        if (hour > 8 || (hour === 8 && minute > 15)) {
          status = "สาย";
          inScore = 0; // สายขาเข้า = 0
        } else {
          inScore = 3; // ปกติขาเข้า = 3
        }
      }

      // 2. ตัดสิน "ขาออก" (ให้ 3 คะแนนเสมอถ้ามีเวลา และไม่ใช่ ขาด/ลา)
      if (outTime.includes(":") && outTime !== "-" && status !== "ขาด" && status !== "ลา") {
        outScore = 3;
      } else {
        outScore = 0;
      }

      uploadRows.push([
        foundDates[d],        // คอลัมน์ A
        current.id,           // คอลัมน์ B
        current.name,         // คอลัมน์ C
        current.room,         // คอลัมน์ D
        inTime,               // คอลัมน์ E
        inScore,              // คอลัมน์ F
        outTime,              // คอลัมน์ G
        outScore,             // คอลัมน์ H
        (inScore + outScore), // คอลัมน์ I
        status                // คอลัมน์ J
      ]);
    }
  }

  // --- บันทึกข้อมูลลงชีต ---
  if (uploadRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, uploadRows.length, 10).setValues(uploadRows);
    return { 
      status: 'success', 
      message: `บันทึกข้อมูลใหม่ ${uploadRows.length} รายการ (พบข้อมูลซ้ำข้ามไป ${duplicateCount} รายการ)` 
    };
  } else {
    return { status: 'success', message: `ไม่พบข้อมูลใหม่ (พบข้อมูลซ้ำทั้งหมด ${duplicateCount} รายการ)` };
  }
}
