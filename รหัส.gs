const SHEET_ID = '1gUfYTrg7pT7sIyCREhRPeQstCtWLtad2T5l2pHNqKTE';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบลงชื่อเข้าใช้งานห้องปฏิบัติการ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันรวมสำหรับเข้าถึง Sheet
function getSheet(sheetName) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
}

// -----------------------------------------------------
// ฟังก์ชันสำหรับหน้า "ลงเวลา"
// -----------------------------------------------------

function getInitialData() {
  const stdSheet = getSheet('std');
  const labSheet = getSheet('labs');
  
  let students = [];
  if (stdSheet.getLastRow() > 1) {
    const stdData = stdSheet.getRange(2, 1, stdSheet.getLastRow() - 1, 3).getValues();
    students = stdData.map(r => ({ id: r[0], name: r[1], class: r[2] }));
  }
  
  let labs = [];
  if (labSheet.getLastRow() > 1) {
    const labData = labSheet.getRange(2, 1, labSheet.getLastRow() - 1, 1).getValues();
    labs = labData.map(r => r[0]).sort(); // เรียง A-Z
  }
  
  return { students: students, labs: labs };
}

function saveCheckIn(data) {
  try {
    const sheet = getSheet('checkin');
    const now = new Date();
    // เรียงคอลัมน์: เลขประจำตัว, ชื่อ-สกุล, ชั้น, ห้องปฏิบัติการ, วันที่, เวลา
    sheet.appendRow([
      data.id,
      data.name,
      data.class,
      data.lab,
      Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd"),
      Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss")
    ]);
    return { success: true, message: 'บันทึกการลงเวลาเรียบร้อย' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// -----------------------------------------------------
// ฟังก์ชันสำหรับสรุปข้อมูล / รายงาน
// -----------------------------------------------------

function getDashboardData() {
  const sheet = getSheet('checkin');
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  // Map ข้อมูลและส่งกลับไปเรียง z-a ที่ Frontend
  return data.map(r => ({
    id: r[0], name: r[1], class: r[2], lab: r[3], date: r[4], time: r[5]
  }));
}

// -----------------------------------------------------
// ฟังก์ชันสำหรับหน้า "Admin"
// -----------------------------------------------------

function checkAdminAuth(user, pass) {
  if (user === 'admin' && pass === '@min6811') {
    return { success: true };
  }
  return { success: false };
}

function uploadCSVData(sheetName, csvData) {
  try {
    const sheet = getSheet(sheetName);
    const existingData = sheet.getDataRange().getValues();
    const isStd = sheetName === 'std'; // ตรวจสอบว่าเป็น sheet นักเรียนหรือไม่
    
    // แปลง CSV data เป็น Array
    let parsedData = Utilities.parseCsv(csvData);
    if(parsedData.length > 0 && (parsedData[0][0] == 'เลขประจำตัว' || parsedData[0][0] == 'ห้องปฏิบัติการ')) {
      parsedData.shift(); // เอา Header ออก
    }
    
    let newCount = 0;
    let updateCount = 0;

    parsedData.forEach(row => {
      let foundIndex = -1;
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] == row[0]) { // เช็ค Primary key (ID หรือ ชื่อห้อง)
          foundIndex = i + 1;
          break;
        }
      }
      
      if (foundIndex !== -1) {
        // เขียนทับข้อมูลเดิม
        sheet.getRange(foundIndex, 1, 1, row.length).setValues([row]);
        updateCount++;
      } else {
        // เพิ่มข้อมูลใหม่
        sheet.appendRow(row);
        newCount++;
      }
    });
    
    return { success: true, message: `อัปเดต: ${updateCount} รายการ, เพิ่มใหม่: ${newCount} รายการ` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function clearAllData(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  return { success: true };
}
