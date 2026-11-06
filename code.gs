
const SPREADSHEET_ID = "ใส่_ID_GOOGLE_SHEETS_ตรงนี้";

const STUDENT_SHEET = "รายชื่อนักเรียน";
const LOG_SHEET = "บันทึกการใช้งาน";

// =====================================
// ตั้งค่าฐานข้อมูลครั้งแรก
// =====================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.setSpreadsheetTimeZone("Asia/Bangkok");

  let studentSheet = ss.getSheetByName(STUDENT_SHEET);
  if (!studentSheet) {
    studentSheet = ss.insertSheet(STUDENT_SHEET);
  }

  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
  }

  if (studentSheet.getLastRow() === 0) {
    studentSheet.appendRow([
      "ชั้นเรียน",
      "รหัสนักเรียน",
      "ชื่อ-นามสกุล"
    ]);
  }

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "วันที่",
      "เวลา",
      "รหัสนักเรียน",
      "ชื่อ-นามสกุล",
      "ชั้นเรียน",
      "วัตถุประสงค์"
    ]);
  }

  studentSheet.setFrozenRows(1);
  logSheet.setFrozenRows(1);

  return "สร้างฐานข้อมูลเรียบร้อย";
}

// =====================================
// อ่านข้อมูลจาก Google Sheets
// =====================================
function getStudents() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENT_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 3
  ).getDisplayValues();

  return values
    .filter(row => row[0] && row[1] && row[2])
    .map(row => ({
      className: row[0],
      studentId: row[1],
      name: row[2]
    }));
}

// =====================================
// ส่งรายชื่อนักเรียนให้ HTML
// =====================================
function doGet(e) {
  const action = e.parameter.action || "";
  const callback = e.parameter.callback || "";

  let result;

  try {
    if (action === "students") {
      result = {
        success: true,
        students: getStudents()
      };
    } else {
      result = {
        success: true,
        message: "ระบบฐานข้อมูลนักเรียนพร้อมใช้งาน"
      };
    }
  } catch (error) {
    result = {
      success: false,
      message: error.message
    };
  }

  if (callback && /^[a-zA-Z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(
        callback + "(" + JSON.stringify(result) + ");"
      )
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================
// บันทึกข้อมูลนักเรียน
// =====================================
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const data = JSON.parse(e.postData.contents);

    if (!data.studentId || !data.name ||
        !data.className || !data.purpose) {
      throw new Error("ข้อมูลไม่ครบถ้วน");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LOG_SHEET);

    if (!sheet) {
      throw new Error("ไม่พบชีตบันทึกการใช้งาน");
    }

    const now = new Date();

    const date = Utilities.formatDate(
      now,
      "Asia/Bangkok",
      "dd/MM/yyyy"
    );

    const time = Utilities.formatDate(
      now,
      "Asia/Bangkok",
      "HH:mm:ss"
    );

    sheet.appendRow([
      date,
      time,
      data.studentId,
      data.name,
      data.className,
      data.purpose
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: "บันทึกข้อมูลสำเร็จ"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
