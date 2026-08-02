// ==========================================
// CONFIGURATION (กำหนดค่าเริ่มต้นตรงนี้)
// ==========================================
// Folder ID จาก Google Drive ของคุณ (ใส่ให้เรียบร้อยแล้ว)
const GOOGLE_DRIVE_FOLDER_ID = "1s9w3texzqb8Ons1cJKySdvT1qOl1-0Wg"; 

function doPost(e) {
  try {
    // รับข้อมูล JSON ที่ส่งมาจากหน้าเว็บ
    const data = JSON.parse(e.postData.contents);
    
    // ตรวจสอบว่าเป็นข้อมูลบันทึกคะแนนสอบ
    if (data.type === "quizScore" || data.action === "submitQuiz" || data.action === "submitMidterm") {
      const targetSheetName = (data.unitId === "Midterm" || data.sheetName === "Mid_Scores" || data.unit === "Midterm") ? "Mid_Scores" : "Quiz_Scores";
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) {
        sheet = ss.insertSheet(targetSheetName);
        sheet.appendRow([
          "Timestamp", 
          "รหัสนักเรียน (Student ID)", 
          "ชื่อ-นามสกุล (Student Name)", 
          "ห้องเรียน (Classroom)", 
          "เลขหน่วยการเรียนรู้ (Unit ID)", 
          "ชื่อบทเรียน (Unit Title)", 
          "คะแนนสอบที่ได้ (Score)", 
          "จำนวนข้อทั้งหมด (Total Questions)", 
          "คะแนนคิดเป็นร้อยละ (Percentage)"
        ]);
        // ตกแต่งหัวตารางของชีตคะแนนสอบให้สวยงาม
        sheet.getRange(1, 1, 1, 9)
             .setFontWeight("bold")
             .setBackground("#2563eb")
             .setFontColor("#ffffff")
             .setHorizontalAlignment("center");
      }
      
      sheet.appendRow([
        data.timestamp || new Date().toLocaleString('th-TH'),
        data.studentId,
        data.studentName,
        data.classRoom || data.classroom,
        data.unitId || data.unit || "Midterm",
        data.unitTitle || data.quizTitle || "ข้อสอบกลางภาค (หน่วยที่ 1-2)",
        data.score,
        data.totalQuestions || data.total || 40,
        (data.percentage !== undefined ? data.percentage : Math.round((data.score / (data.totalQuestions || 40)) * 100)) + "%"
      ]);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "บันทึกคะแนนสอบลง Google Sheets เรียบร้อยแล้ว"
      }))
      .setMimeType(ContentService.MimeType.JSON);
    }
    
    // ตรวจสอบว่าเป็นข้อมูลบันทึกคะแนนและคำแนะนำจากครู
    if (data.type === "updateGrade") {
      const passcode = data.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "unauthorized",
          message: "รหัสผ่านผู้สอนไม่ถูกต้อง"
        }))
        .setMimeType(ContentService.MimeType.JSON);
      }
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const sheetData = sheet.getDataRange().getValues();
      let targetRow = -1;
      
      const inputStudentId = normalizeStr(data.studentId);
      const inputAssignment = normalizeStr(data.assignment);
      
      for (let i = 1; i < sheetData.length; i++) {
        const sheetReceiptId = sheetData[i][1];
        const sheetStudentId = normalizeStr(sheetData[i][2]);
        const sheetAssignment = normalizeStr(sheetData[i][5]);
        
        if ((data.receiptId && sheetReceiptId === data.receiptId) || 
            (sheetStudentId === inputStudentId && sheetAssignment === inputAssignment)) {
          targetRow = i + 1; // ลำดับแถวจริงใน Sheet (1-based)
          break;
        }
      }
      
      if (targetRow !== -1) {
        sheet.getRange(targetRow, 9).setValue("checked");            // Status (Col I)
        sheet.getRange(targetRow, 10).setValue(data.grade);          // Grade (Col J)
        sheet.getRange(targetRow, 11).setValue(data.feedback);       // Teacher Feedback (Col K)
        
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          message: "บันทึกคะแนนและข้อเสนอแนะเรียบร้อยแล้ว"
        }))
        .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          status: "notFound",
          message: "ไม่พบข้อมูลการส่งงานของนักเรียนในระบบ"
        }))
        .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // 1. จัดการอัปโหลดไฟล์ไปยัง Google Drive
    const folder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
    
    // แปลงไฟล์จาก Base64 กลับมาเป็นไฟล์ปกติ
    const contentType = data.mimeType || "application/octet-stream";
    const decodedFile = Utilities.base64Decode(data.fileData);
    const blob = Utilities.newBlob(decodedFile, contentType, data.fileName);
    
    // บันทึกไฟล์ลงใน Drive
    const file = folder.createFile(blob);
    // กำหนดสิทธิ์ให้ทุกคนที่มีลิงก์สามารถเปิดดูหรือดาวน์โหลดได้ (สำหรับคุณครูตรวจงาน)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    // 2. บันทึกข้อมูลลงใน Google Sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetData = sheet.getDataRange().getValues();
    
    let receiptId = data.id; // ใช้ ID ที่ส่งมาจากฝั่งไคลเอนต์เป็นค่าเริ่มต้น
    let isUpdate = false;
    let targetRow = -1;
    
    // ตรวจสอบว่านักเรียนคนนี้เคยส่งใบงานนี้ไปหรือยัง (อ้างอิงจาก รหัสนักเรียน คอลัมน์ C และ ใบงาน คอลัมน์ F)
    // คอลัมน์ C คือ ดัชนีที่ 2, คอลัมน์ F คือ ดัชนีที่ 5
    for (let i = 1; i < sheetData.length; i++) {
      const sheetStudentId = normalizeStr(sheetData[i][2]);
      const sheetAssignment = normalizeStr(sheetData[i][5]);
      const inputStudentId = normalizeStr(data.studentId);
      const inputAssignment = normalizeStr(data.assignment);
      
      if (sheetStudentId === inputStudentId && sheetAssignment === inputAssignment) {
        targetRow = i + 1; // ลำดับแถวจริงใน Sheet (1-based)
        receiptId = sheetData[i][1]; // ดึง Receipt ID เดิมมาใช้ซ้ำ เพื่อให้ใบเสร็จนักเรียนมีรหัสเดิม
        isUpdate = true;
        break;
      }
    }
    
    if (isUpdate && targetRow !== -1) {
      // เขียนทับข้อมูลในแถวเดิม เพื่อไม่ให้แถวใน Google Sheet ซ้ำซ้อนกัน
      sheet.getRange(targetRow, 1).setValue(data.timestamp);       // Timestamp (Col A)
      sheet.getRange(targetRow, 2).setValue(receiptId);            // Receipt ID (Col B)
      sheet.getRange(targetRow, 7).setValue(fileUrl);              // ลิงก์ไฟล์งานตัวใหม่ (Col G)
      sheet.getRange(targetRow, 8).setValue(data.comments);         // หมายเหตุใหม่ (Col H)
      sheet.getRange(targetRow, 9).setValue("pending");            // รีเซ็ตสถานะเป็น รอตรวจ (Col I)
      sheet.getRange(targetRow, 10).setValue("-");                 // รีเซ็ตคะแนนใหม่ (Col J)
      sheet.getRange(targetRow, 11).setValue("-");                 // รีเซ็ตความเห็นครู (Col K)
    } else {
      // เพิ่มแถวใหม่ (กรณีส่งใบงานนี้เป็นครั้งแรก)
      sheet.appendRow([
        data.timestamp,       // Timestamp (Col A)
        receiptId,            // Receipt ID (Col B)
        data.studentId,       // Student ID (Col C)
        data.studentName,     // Student Name (Col D)
        data.classRoom,       // Classroom (Col E)
        data.assignment,      // Assignment (Col F)
        fileUrl,              // File Link (Col G)
        data.comments,        // Comments (Col H)
        "pending",            // Status (Col I)
        "-",                  // Grade (Col J)
        "-"                   // Teacher Feedback (Col K)
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      fileUrl: fileUrl,
      receiptId: receiptId,
      isUpdate: isUpdate,
      message: isUpdate ? "อัปเดตไฟล์งานเรียบร้อยแล้ว" : "ข้อมูลถูกบันทึกเรียบร้อยแล้ว"
    }))
    .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

// คีย์รหัสผ่านคุณครูสำหรับเข้าถึงแดชบอร์ด
const TEACHER_PASSCODE = "teacher123";

// ค้นหาข้อมูลสถานะการตรวจและคะแนนจาก Google Sheets (เพื่อแสดงผลบนหน้าเว็บนักเรียนตามจริง)
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === "checkStatus") {
      const receiptId = e.parameter.id;
      if (!receiptId) {
        return createJsonResponse({ status: "error", message: "ไม่พบรหัสใบเสร็จในพารามิเตอร์" });
      }
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const data = sheet.getDataRange().getValues();
      
      // วนลูปหาแถวที่ตรงกับ Receipt ID (เริ่มแถวที่ 2 เนื่องจากแถวแรกเป็น Header)
      for (let i = 1; i < data.length; i++) {
        // Col B คือคอลัมน์ที่ 2 (ดัชนีที่ 1) เก็บ Receipt ID
        if (data[i][1] === receiptId) {
          return createJsonResponse({
            status: "success",
            submissionStatus: data[i][8],      // คอลัมน์ I (ดัชนีที่ 8) คือ Status
            grade: data[i][9],                 // คอลัมน์ J (ดัชนีที่ 9) คือ Grade
            teacherFeedback: data[i][10]       // คอลัมน์ K (ดัชนีที่ 10) คือ Teacher Feedback
          });
        }
      }
      
      return createJsonResponse({ status: "notFound", message: "ไม่พบรหัสใบเสร็จนี้ในระบบฐานข้อมูล" });
    }
    
    // ตรวจสอบประวัติการส่งงานของนักเรียนจากเลขประจำตัว 5 หลัก
    if (action === "getStudentSubmissions") {
      const studentId = e.parameter.studentId;
      if (!studentId) {
        return createJsonResponse({ status: "error", message: "ไม่พบรหัสประจำตัวนักเรียน" });
      }
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const submissions = [];
      const normalizedQuery = normalizeStr(studentId);
      
      for (let i = 1; i < data.length; i++) {
        if (normalizeStr(data[i][2]) === normalizedQuery) { // Col C (Index 2) คือ Student ID
          submissions.push({
            timestamp: data[i][0],      // Col A
            receiptId: data[i][1],      // Col B
            studentId: data[i][2],      // Col C
            studentName: data[i][3],    // Col D
            classRoom: data[i][4],      // Col E
            assignment: data[i][5],     // Col F
            fileLink: data[i][6],       // Col G
            comments: data[i][7],       // Col H
            status: data[i][8],         // Col I
            grade: data[i][9],          // Col J
            teacherFeedback: data[i][10] // Col K
          });
        }
      }
      
      return createJsonResponse({
        status: "success",
        submissions: submissions
      });
    }
    
    // ดึงข้อมูลการส่งงานทั้งหมดสำหรับแดชบอร์ดคุณครู
    if (action === "getAllSubmissions") {
      const passcode = e.parameter.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const submissions = [];
      
      for (let i = 1; i < data.length; i++) {
        submissions.push({
          timestamp: data[i][0],      // Col A
          receiptId: data[i][1],      // Col B
          studentId: data[i][2],      // Col C
          studentName: data[i][3],    // Col D
          classRoom: data[i][4],      // Col E
          assignment: data[i][5],     // Col F
          fileLink: data[i][6],       // Col G
          comments: data[i][7],       // Col H
          status: data[i][8],         // Col I
          grade: data[i][9],          // Col J
          teacherFeedback: data[i][10] // Col K
        });
      }
      
      return createJsonResponse({
        status: "success",
        submissions: submissions
      });
    }
    
    return createJsonResponse({ status: "error", message: "Action ไม่ถูกต้อง" });
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันช่วยทำความสะอาดข้อความเพื่อเปรียบเทียบ (ลบเว้นวรรค เครื่องหมายจุด ทวิภาค ฯลฯ)
function normalizeStr(str) {
  if (str === null || str === undefined) return "";
  return str.toString()
    .toLowerCase()
    .replace(/[\s\:\.\-\_\/\\]/g, "")
    .trim();
}
