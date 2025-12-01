// ID Spreadsheet Sumber Data
const SOURCE_SHEET_ID = "1n6juQuCCmtZoZ_bxCGmwOHKjayiQK9lLMrar4Promss"; 
const SOURCE_SHEET_NAME = "Sheet1";

// ID Spreadsheet Database Nilai
const DB_SHEET_ID = "17SUH6YUHFidAhbE74jCzPG2MW--owC-FNZqYh1SyHyI"; 
const DB_SHEET_NAME = "Sheet1";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Aplikasi PABP - SMK N 1 Maluku Tengah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * FUNGSI: Mengambil Data Siswa & Nilai
 * * UPDATE STRUKTUR SHEET SUMBER:
 * Col B [0] = Nama
 * Col C [1] = NIS
 * Col D [2] = NISN
 * Col E [3] = Jenis Kelamin (JK)  <-- PERBAIKAN
 * Col F [4] = Tempat Tanggal Lahir (TTL) <-- PERBAIKAN
 */
function getStudentData() {
  const ssSource = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const sheetSource = ssSource.getSheetByName(SOURCE_SHEET_NAME);
  
  const lastRowSource = sheetSource.getLastRow();
  let sourceData = [];
  if (lastRowSource > 1) {
    // Ambil 5 Kolom (B, C, D, E, F)
    sourceData = sheetSource.getRange(2, 2, lastRowSource - 1, 5).getValues(); 
  }

  const ssDB = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheetDB = ssDB.getSheetByName(DB_SHEET_NAME);
  
  if (sheetDB.getLastRow() === 0) {
    sheetDB.appendRow(["Nama Siswa", "TP1", "TP2", "TP3", "TP4", "TP5", "LM1", "LM2", "LM3", "SAS"]);
  }

  const lastRowDB = sheetDB.getLastRow();
  let dbData = [];
  if (lastRowDB > 1) {
    dbData = sheetDB.getRange(2, 1, lastRowDB - 1, 10).getValues();
  }

  let dbMap = {};
  dbData.forEach(row => { dbMap[row[0]] = row; }); 

  let finalOutput = [];
  
  sourceData.forEach((rowSource) => {
    // Ambil data dari array source (0-4)
    let nama = rowSource[0]; // Col B
    let nis  = rowSource[1]; // Col C
    let nisn = rowSource[2]; // Col D
    let jk   = rowSource[3]; // Col E (JK)
    let ttl  = rowSource[4]; // Col F (TTL)
    
    // Kita susun ke Array Frontend. 
    // Urutan Index Internal Frontend (Bebas, asal konsisten dengan JS):
    // [0:Nama, 1:NIS, 2:NISN, 3:TTL, 4:JK, 5..13:Nilai]
    
    let gradeData = ["", "", "", "", "", "", "", "", ""]; 
    
    if (dbMap[nama]) {
      gradeData = dbMap[nama].slice(1);
    } 
    
    // Masukkan ke array output
    finalOutput.push([nama, nis, nisn, ttl, jk, ...gradeData]);
  });

  return finalOutput;
}

function saveData(frontendData) {
  try {
    const ssDB = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheetDB = ssDB.getSheetByName(DB_SHEET_NAME);
    
    // Simpan Nama (idx 0) dan Nilai (idx 5-13)
    let dbPayload = frontendData.map(row => {
      return [row[0], ...row.slice(5)];
    });

    let maxRows = sheetDB.getMaxRows();
    if (maxRows > 1) {
      sheetDB.getRange(2, 1, maxRows - 1, 10).clearContent();
    }
    if (dbPayload.length > 0) {
      sheetDB.getRange(2, 1, dbPayload.length, 10).setValues(dbPayload);
    }
    return " ✅  Data Berhasil Disimpan!";
  } catch (e) {
    return " ❌  Gagal: " + e.toString();
  }
}

function resetAllScores() {
  try {
    const ssDB = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheetDB = ssDB.getSheetByName(DB_SHEET_NAME);
    const lastRow = sheetDB.getLastRow();
    
    if (lastRow > 1) {
      sheetDB.getRange(2, 2, lastRow - 1, 9).clearContent();
    }
    return "success";
  } catch (e) {
    return "error: " + e.toString();
  }
}
