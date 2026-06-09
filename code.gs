const SHEET_NAME = "Data_Siswa";

function doGet(e) {
  getDatabaseSheet(); 
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PENGUMUMAN KELULUSAN SMPS IT IBNU HALIM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDatabaseSheet() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty("DB_ID");
  let ss;
  
  if (dbId) {
    try { ss = SpreadsheetApp.openById(dbId); } 
    catch (e) { props.deleteProperty("DB_ID"); }
  }
  
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } 
    catch (e) { ss = null; }
    
    if (!ss) {
      ss = SpreadsheetApp.create("Database_SMPIT_IH_Medan");
      props.setProperty("DB_ID", ss.getId());
    }
  }
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [["NISN", "Nama Lengkap", "Kelas", "Tempat, Tgl Lahir", "Status Kelulusan"]];
    sheet.getRange("A1:E1").setValues(headers);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3");
    
    // Format kolom NISN menjadi Plain Text
    sheet.getRange("A:A").setNumberFormat("@");
    
    // Data Dummy untuk pengujian
    sheet.appendRow(["0116607750", "Dapa Sabara", "9", "Kerinci, 10 Mei 2011", "LULUS"]);
    sheet.appendRow(["0111266734", "Amira", "9", "Kerinci, 28 November 2011", "LULUS"]);
    sheet.appendRow(["1243", "Andi Irawan", "9", "Padang, 01 Januari 2011", "TIDAK LULUS"]);
    
    const sheet1 = ss.getSheetByName("Sheet1");
    if (sheet1 && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet1);
    }
  }
  return sheet;
}

function searchStudent(nisnDariKlien) {
  if (!nisnDariKlien) return { success: false, message: "NISN tidak boleh kosong." };

  try {
    const sheet = getDatabaseSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return { success: false, message: "Database masih kosong." };
    
    data.shift(); // Buang baris header
    
    const student = data.find(row => String(row[0]).trim() === String(nisnDariKlien).trim());

    if (student) {
      let tanggalLahir = student[3];
      if (tanggalLahir instanceof Date) {
         tanggalLahir = Utilities.formatDate(tanggalLahir, Session.getScriptTimeZone(), "dd MMMM yyyy");
      } else {
         tanggalLahir = String(tanggalLahir);
      }

      return {
        success: true,
        data: {
          nisn: String(student[0]),
          nama: String(student[1]),
          kelas: String(student[2]),
          lahir: tanggalLahir,
          status: String(student[4]).toUpperCase()
        }
      };
    }
    return { success: false, message: "Siswa dengan NISN tersebut tidak ditemukan." };
  } catch (e) {
    return { success: false, message: "Terjadi kesalahan server: " + e.message };
  }
}