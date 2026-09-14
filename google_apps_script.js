// ==========================================
// CONFIGURATION (กำหนดค่าเริ่มต้น)
// ==========================================
const SPREADSHEET_ID = "1tICg3H5kfGAp_uoQrUI1WWAqdfGSGqurgqOZQvuOd0A";
const GOOGLE_DRIVE_FOLDER_ID = "1s9w3texzqb8Ons1cJKySdvT1qOl1-0Wg"; 
const TEACHER_PASSCODE = "teacher123";

// ฟังก์ชันดึง Spreadsheet ตาม ID หรือ Active Spreadsheet
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ฟังก์ชันระบุชีตส่งใบงานหลัก
function getSubmissionSheet(ss) {
  let sheet = ss.getSheetByName("Submissions");
  if (!sheet) sheet = ss.getSheetByName("Form Responses 1");
  if (!sheet) sheet = ss.getSheetByName("Sheet1");
  if (!sheet) sheet = ss.getSheets()[0];
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();
    
    // 1. บันทึกคะแนนสอบ (Midterm / Final / Quiz)
    if (data.type === "quizScore" || data.action === "submitQuiz" || data.action === "submitMidterm" || data.action === "submitFinal") {
      let targetSheetName = "Quiz_Scores";
      if (data.sheetName) {
        targetSheetName = data.sheetName;
      } else if (data.unitId === "Midterm" || data.unit === "Midterm") {
        targetSheetName = "Mid_Scores";
      } else if (data.unitId === "Final" || data.unit === "Final") {
        targetSheetName = "Final_Scores";
      }

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
        data.unitId || data.unit || "Exam",
        data.unitTitle || data.quizTitle || "ข้อสอบวัดผลสัมฤทธิ์",
        data.score,
        data.totalQuestions || data.total || 40,
        (data.percentage !== undefined ? data.percentage : Math.round((data.score / (data.totalQuestions || 40)) * 100)) + "%"
      ]);
      
      return createJsonResponse({
        status: "success",
        message: "บันทึกคะแนนสอบลง Google Sheets เรียบร้อยแล้ว"
      });
    }

    // 2. บันทึก/ซิงก์คลังข้อสอบแยกชีต (Mid_Questions, Final_Questions, Quiz_Questions)
    if (data.action === "saveQuestionBank") {
      const passcode = data.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }

      const examType = data.examType || "midterm";
      let targetSheetName = "Mid_Questions";
      if (examType === "final") targetSheetName = "Final_Questions";
      if (examType === "quiz") targetSheetName = "Quiz_Questions";

      let sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) {
        sheet = ss.insertSheet(targetSheetName);
      } else {
        sheet.clear();
      }

      sheet.appendRow([
        "ID", "Unit", "Unit_Title", "Question", 
        "Choice_A", "Choice_B", "Choice_C", "Choice_D", 
        "Correct_Answer", "Explanation"
      ]);
      sheet.getRange(1, 1, 1, 10)
           .setFontWeight("bold")
           .setBackground("#1e293b")
           .setFontColor("#ffffff")
           .setHorizontalAlignment("center");

      const questions = data.questions || [];
      questions.forEach((q, idx) => {
        const choices = q.choices || {};
        sheet.appendRow([
          idx + 1,
          q.unit || 1,
          q.unitTitle || "",
          q.question || "",
          choices.A || choices["A"] || "",
          choices.B || choices["B"] || "",
          choices.C || choices["C"] || "",
          choices.D || choices["D"] || "",
          q.correctAnswer || "A",
          q.explanation || ""
        ]);
      });

      return createJsonResponse({
        status: "success",
        message: `บันทึกคลังข้อสอบลงแท็บ ${targetSheetName} จำนวน ${questions.length} ข้อเรียบร้อยแล้ว`
      });
    }

    // 3. บันทึกรายการไฟล์ใบงาน (Worksheets_Bank)
    if (data.action === "saveWorksheets") {
      const passcode = data.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }

      let sheet = ss.getSheetByName("Worksheets_Bank");
      if (!sheet) {
        sheet = ss.insertSheet("Worksheets_Bank");
      } else {
        sheet.clear();
      }

      sheet.appendRow(["Unit", "Title", "URL"]);
      sheet.getRange(1, 1, 1, 3)
           .setFontWeight("bold")
           .setBackground("#059669")
           .setFontColor("#ffffff")
           .setHorizontalAlignment("center");

      const worksheets = data.worksheets || [];
      worksheets.forEach(w => {
        sheet.appendRow([w.unit || "", w.title || "", w.url || ""]);
      });

      return createJsonResponse({
        status: "success",
        message: `บันทึกรายการใบงานลงแท็บ Worksheets_Bank จำนวน ${worksheets.length} รายการเรียบร้อยแล้ว`
      });
    }

    // บันทึกคลังสื่อการเรียนรู้ (Media_Bank)
    if (data.action === "saveMediaBank") {
      const passcode = data.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }

      let sheet = ss.getSheetByName("Media_Bank");
      if (!sheet) {
        sheet = ss.insertSheet("Media_Bank");
      } else {
        sheet.clear();
      }

      sheet.appendRow(["Category", "Title", "Description", "Type", "Link_URL", "Thumbnail_URL"]);
      sheet.getRange(1, 1, 1, 6)
           .setFontWeight("bold")
           .setBackground("#7c3aed")
           .setFontColor("#ffffff")
           .setHorizontalAlignment("center");

      const mediaItems = data.mediaItems || [];
      mediaItems.forEach(m => {
        sheet.appendRow([
          m.category || "Video",
          m.title || "",
          m.description || "",
          m.type || "",
          m.url || "",
          m.thumbnail || ""
        ]);
      });

      return createJsonResponse({
        status: "success",
        message: `บันทึกสื่อการเรียนรู้ลงแท็บ Media_Bank จำนวน ${mediaItems.length} รายการเรียบร้อยแล้ว`
      });
    }

    // 4. บันทึกสิทธิ์อนุญาตสอบซ้ำ (Exam_Locks)
    if (data.action === "unlockStudentExam") {
      const passcode = data.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }

      let sheet = ss.getSheetByName("Exam_Locks");
      if (!sheet) {
        sheet = ss.insertSheet("Exam_Locks");
        sheet.appendRow(["Timestamp", "Student_ID", "Exam_Type", "Status"]);
        sheet.getRange(1, 1, 1, 4)
             .setFontWeight("bold")
             .setBackground("#dc2626")
             .setFontColor("#ffffff")
             .setHorizontalAlignment("center");
      }

      sheet.appendRow([
        new Date().toLocaleString('th-TH'),
        data.studentId,
        data.examType || "midterm",
        "unlocked"
      ]);

      return createJsonResponse({
        status: "success",
        message: `ปลดล็อกการสอบสำหรับนักเรียนรหัส ${data.studentId} เรียบร้อยแล้ว`
      });
    }

    // 5. ตรวจและบันทึกคะแนนใบงานโดยครูผู้สอน
    if (data.action === "gradeSubmission" || data.action === "updateGrade" || data.type === "updateGrade" || data.type === "gradeSubmission") {
      const passcode = data.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านครูผู้สอนไม่ถูกต้อง" });
      }

      const receiptId = data.receiptId;
      const studentIdStr = String(data.studentId || "").trim();
      const assignmentStr = String(data.assignment || "").trim();
      const grade = data.grade;
      const teacherFeedback = data.teacherFeedback || data.feedback || "-";
      
      const sheet = getSubmissionSheet(ss);
      const sheetData = sheet.getDataRange().getValues();
      let updated = false;

      for (let i = 1; i < sheetData.length; i++) {
        const rowReceipt = String(sheetData[i][1]).trim();
        const rowStudentId = String(sheetData[i][2]).trim();
        const rowAssignment = String(sheetData[i][5]).trim();

        if ((receiptId && rowReceipt === String(receiptId).trim()) || (studentIdStr && assignmentStr && rowStudentId === studentIdStr && rowAssignment === assignmentStr)) {
          sheet.getRange(i + 1, 9).setValue("checked");         
          sheet.getRange(i + 1, 10).setValue(grade);            
          sheet.getRange(i + 1, 11).setValue(teacherFeedback);   
          updated = true;
          break;
        }
      }

      if (updated) {
        return createJsonResponse({ status: "success", message: "บันทึกผลการตรวจเรียบร้อยแล้ว" });
      } else {
        return createJsonResponse({ status: "notFound", message: "ไม่พบรายการส่งงานรหัสนี้" });
      }
    }
    
    // 6. การอัปโหลดไฟล์ใบงานของนักเรียนลง Google Drive
    if (data.fileData || data.action === "uploadWorksheet" || data.action === "submitWorksheet") {
      const folder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
      const contentType = data.mimeType || "application/octet-stream";
      const bytes = Utilities.base64Decode(data.fileData);
      const blob = Utilities.newBlob(bytes, contentType, data.fileName || "file_submission");
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      const fileUrl = file.getUrl();
      const sheet = getSubmissionSheet(ss);
      
      const sheetData = sheet.getDataRange().getValues();
      let targetRow = -1;
      let receiptId = "";
      let isUpdate = false;
      
      const studentIdStr = String(data.studentId).trim();
      const assignmentStr = String(data.assignment).trim();
      
      for (let i = 1; i < sheetData.length; i++) {
        const existingStudentId = String(sheetData[i][2]).trim();
        const existingAssignment = String(sheetData[i][5]).trim();
        
        if (existingStudentId === studentIdStr && existingAssignment === assignmentStr) {
          targetRow = i + 1;
          receiptId = sheetData[i][1];
          isUpdate = true;
          break;
        }
      }
      
      if (!isUpdate) {
        const datePart = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd");
        const randomPart = Math.floor(1000 + Math.random() * 9000);
        receiptId = "DT-" + datePart + "-" + randomPart;
      }
      
      if (isUpdate && targetRow !== -1) {
        sheet.getRange(targetRow, 1).setValue(data.timestamp);       
        sheet.getRange(targetRow, 2).setValue(receiptId);            
        sheet.getRange(targetRow, 7).setValue(fileUrl);              
        sheet.getRange(targetRow, 8).setValue(data.comments);         
        sheet.getRange(targetRow, 9).setValue("pending");            
        sheet.getRange(targetRow, 10).setValue("-");                 
        sheet.getRange(targetRow, 11).setValue("-");                 
      } else {
        sheet.appendRow([
          data.timestamp,
          receiptId,
          data.studentId,
          data.studentName,
          data.classRoom,
          data.assignment,
          fileUrl,
          data.comments,
          "pending",
          "-",
          "-"
        ]);
      }

      return createJsonResponse({
        status: "success",
        fileUrl: fileUrl,
        receiptId: receiptId,
        isUpdate: isUpdate,
        message: isUpdate ? "อัปเดตไฟล์งานเรียบร้อยแล้ว" : "ข้อมูลถูกบันทึกเรียบร้อยแล้ว"
      });
    }

    return createJsonResponse({
      status: "error",
      message: "ไม่พบคำสั่ง (Action) ที่ระบุ หรือไฟล์ข้อมูลไม่ถูกต้อง"
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

// ค้นหาข้อมูลสถานะการตรวจ คะแนน คลังข้อสอบ ใบงาน และสิทธิ์สอบซ้ำจาก Google Sheets
function doGet(e) {
  try {
    const action = e.parameter.action;
    const ss = getSpreadsheet();
    
    // 1. ดึงคลังข้อสอบแยกชีต (Mid_Questions, Final_Questions, Quiz_Questions)
    if (action === "getQuestionBank") {
      const getQuestionsFromSheet = (sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return null;
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return [];
        
        const questions = [];
        for (let i = 1; i < data.length; i++) {
          if (!data[i][3]) continue;
          questions.push({
            id: data[i][0] || i,
            unit: data[i][1] || 1,
            unitTitle: data[i][2] || "",
            question: data[i][3] || "",
            choices: {
              "A": data[i][4] || "",
              "B": data[i][5] || "",
              "C": data[i][6] || "",
              "D": data[i][7] || ""
            },
            correctAnswer: String(data[i][8] || "A").trim().toUpperCase(),
            explanation: data[i][9] || ""
          });
        }
        return questions;
      };

      return createJsonResponse({
        status: "success",
        midtermQuestions: getQuestionsFromSheet("Mid_Questions"),
        finalQuestions: getQuestionsFromSheet("Final_Questions"),
        quizQuestions: getQuestionsFromSheet("Quiz_Questions")
      });
    }

    // 2. ดึงรายการใบงาน (Worksheets_Bank)
    if (action === "getWorksheets") {
      const sheet = ss.getSheetByName("Worksheets_Bank");
      if (!sheet) return createJsonResponse({ status: "success", worksheets: [] });
      const data = sheet.getDataRange().getValues();
      const worksheets = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i][1]) continue;
        worksheets.push({
          unit: data[i][0] || "",
          title: data[i][1] || "",
          url: data[i][2] || ""
        });
      }
      return createJsonResponse({ status: "success", worksheets: worksheets });
    }

    // ดึงคลังสื่อการเรียนรู้ (Media_Bank)
    if (action === "getMediaBank") {
      const sheet = ss.getSheetByName("Media_Bank");
      if (!sheet) return createJsonResponse({ status: "success", mediaItems: [] });
      const data = sheet.getDataRange().getValues();
      const mediaItems = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i][1]) continue;
        mediaItems.push({
          category: data[i][0] || "Video",
          title: data[i][1] || "",
          description: data[i][2] || "",
          type: data[i][3] || "",
          url: data[i][4] || "",
          thumbnail: data[i][5] || ""
        });
      }
      return createJsonResponse({ status: "success", mediaItems: mediaItems });
    }

    // 3. ดึงสิทธิ์อนุญาตสอบซ้ำ (Exam_Locks)
    if (action === "getExamLocks") {
      const sheet = ss.getSheetByName("Exam_Locks");
      if (!sheet) return createJsonResponse({ status: "success", locks: [] });
      const data = sheet.getDataRange().getValues();
      const locks = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i][1]) continue;
        locks.push({
          timestamp: data[i][0],
          studentId: String(data[i][1]).trim(),
          examType: data[i][2] || "midterm",
          status: data[i][3] || "unlocked"
        });
      }
      return createJsonResponse({ status: "success", locks: locks });
    }

    // 4. ค้นหาตาม Receipt ID
    if (action === "checkStatus") {
      const receiptId = e.parameter.id;
      if (!receiptId) {
        return createJsonResponse({ status: "error", message: "ไม่พบรหัสใบเสร็จในพารามิเตอร์" });
      }
      
      const sheet = getSubmissionSheet(ss);
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === receiptId) {
          return createJsonResponse({
            status: "success",
            submissionStatus: data[i][8],
            grade: data[i][9],
            teacherFeedback: data[i][10]
          });
        }
      }
      
      return createJsonResponse({ status: "notFound", message: "ไม่พบรหัสใบเสร็จนี้ในระบบฐานข้อมูล" });
    }
    
    // 5. ตรวจสอบประวัติการส่งงานของนักเรียนจากรหัสประจำตัว 5 หลัก
    if (action === "getStudentSubmissions") {
      const studentId = e.parameter.studentId;
      if (!studentId) {
        return createJsonResponse({ status: "error", message: "ไม่พบรหัสประจำตัวนักเรียน" });
      }
      
      const sheet = getSubmissionSheet(ss);
      const data = sheet.getDataRange().getValues();
      const submissions = [];
      const normalizedQuery = normalizeStr(studentId);
      
      for (let i = 1; i < data.length; i++) {
        if (normalizeStr(data[i][2]) === normalizedQuery) {
          submissions.push({
            timestamp: data[i][0],
            receiptId: data[i][1],
            studentId: data[i][2],
            studentName: data[i][3],
            classRoom: data[i][4],
            assignment: data[i][5],
            fileLink: data[i][6],
            comments: data[i][7],
            status: data[i][8],
            grade: data[i][9],
            teacherFeedback: data[i][10]
          });
        }
      }
      
      return createJsonResponse({
        status: "success",
        submissions: submissions
      });
    }
    
    // 6. ดึงข้อมูลการส่งงานทั้งหมดสำหรับแดชบอร์ดคุณครู
    if (action === "getAllSubmissions") {
      const passcode = e.parameter.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }
      
      const sheet = getSubmissionSheet(ss);
      const data = sheet.getDataRange().getValues();
      const submissions = [];
      
      for (let i = 1; i < data.length; i++) {
        submissions.push({
          timestamp: data[i][0],
          receiptId: data[i][1],
          studentId: data[i][2],
          studentName: data[i][3],
          classRoom: data[i][4],
          assignment: data[i][5],
          fileLink: data[i][6],
          comments: data[i][7],
          status: data[i][8],
          grade: data[i][9],
          teacherFeedback: data[i][10]
        });
      }
      
      return createJsonResponse({
        status: "success",
        submissions: submissions
      });
    }

    // 7. ดึงข้อมูลคะแนนสอบกลางภาค ปลายภาค และแบบทดสอบย่อยสำหรับแอดมิน
    if (action === "getExamScores") {
      const passcode = e.parameter.passcode;
      if (passcode !== TEACHER_PASSCODE) {
        return createJsonResponse({ status: "unauthorized", message: "รหัสผ่านผู้สอนไม่ถูกต้อง" });
      }
      
      const midtermSheet = ss.getSheetByName("Mid_Scores");
      const finalSheet = ss.getSheetByName("Final_Scores");
      const quizSheet = ss.getSheetByName("Quiz_Scores");
      
      const midtermScores = [];
      if (midtermSheet) {
        const mData = midtermSheet.getDataRange().getValues();
        for (let i = 1; i < mData.length; i++) {
          midtermScores.push({
            timestamp: mData[i][0],
            studentId: mData[i][1],
            studentName: mData[i][2],
            classRoom: mData[i][3],
            unitId: mData[i][4],
            unitTitle: mData[i][5],
            score: mData[i][6],
            totalQuestions: mData[i][7],
            percentage: mData[i][8]
          });
        }
      }

      const finalScores = [];
      if (finalSheet) {
        const fData = finalSheet.getDataRange().getValues();
        for (let i = 1; i < fData.length; i++) {
          finalScores.push({
            timestamp: fData[i][0],
            studentId: fData[i][1],
            studentName: fData[i][2],
            classRoom: fData[i][3],
            unitId: fData[i][4],
            unitTitle: fData[i][5],
            score: fData[i][6],
            totalQuestions: fData[i][7],
            percentage: fData[i][8]
          });
        }
      }
      
      const quizScores = [];
      if (quizSheet) {
        const qData = quizSheet.getDataRange().getValues();
        for (let i = 1; i < qData.length; i++) {
          quizScores.push({
            timestamp: qData[i][0],
            studentId: qData[i][1],
            studentName: qData[i][2],
            classRoom: qData[i][3],
            unitId: qData[i][4],
            unitTitle: qData[i][5],
            score: qData[i][6],
            totalQuestions: qData[i][7],
            percentage: qData[i][8]
          });
        }
      }
      
      return createJsonResponse({
        status: "success",
        midtermScores: midtermScores,
        finalScores: finalScores,
        quizScores: quizScores
      });
    }
    
    return createJsonResponse({ status: "error", message: "Action ไม่ถูกต้อง" });
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

function normalizeStr(str) {
  if (!str) return "";
  return String(str).trim();
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
