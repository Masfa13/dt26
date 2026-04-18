// --- KONFIGURASI NAMA SHEET ---
const SHEETS = {
  TATABOGA: "TATABOGA",
  TKR: "TKR",
  DASHBOARD: "DASHBOARD"
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Rekap Presensi Double Track 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi untuk membuat menu di Spreadsheet agar mudah setting awal
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 DT System')
    .addItem('Setup Spreadsheet Baru', 'initialSetup')
    .addToUi();
}

// Fungsi Setup awal untuk membuat kolom yang diperlukan jika belum ada
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  Object.values(SHEETS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name !== 'DASHBOARD') {
        sheet.appendRow(["NO", "NIS", "NAMA SISWA"]); // Header dasar
        ui.alert('Sheet ' + name + ' berhasil dibuat. Silakan isi data siswa di bawah kolom NAMA SISWA.');
      }
    }
  });
  ui.alert('Proses Setup Selesai. Pastikan nama kolom sesuai.');
}

function getDataPeserta(skill) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(skill);
  if (!sheet) return { peserta: [], dates: [], stats: {} };

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const dates = header.slice(3); // Tanggal mulai dari kolom ke-4 (indeks 3)
  
  const peserta = data.slice(1).map((row, index) => {
    const kehadiran = row.slice(3);
    const rekap = {
      h: kehadiran.filter(s => s === 'H').length,
      i: kehadiran.filter(s => s === 'I').length,
      a: kehadiran.filter(s => s === 'A').length
    };
    
    return {
      rowIndex: index + 2,
      nama: row[2], // Kolom C
      kehadiran: kehadiran,
      rekapIndividu: rekap
    };
  });

  return {
    peserta: peserta,
    dates: dates.map(d => {
      if (d instanceof Date) {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM");
      }
      return d;
    })
  };
}

function getFullDashboardData() {
  const results = {};
  ['TATABOGA', 'TKR'].forEach(skill => {
    const data = getDataPeserta(skill);
    const totalPeserta = data.peserta.length;
    const totalPertemuan = data.dates.length;
    
    let totalHadir = 0;
    data.peserta.forEach(p => totalHadir += p.rekapIndividu.h);
    
    const rataRata = (totalPeserta * totalPertemuan) > 0 
      ? (totalHadir / (totalPeserta * totalPertemuan)) 
      : 0;

    results[skill] = {
      totalPertemuan: totalPertemuan,
      rataRataHadir: rataRata
    };
  });
  return results;
}

function simpanAbsensi(skill, tanggal, dataAbsen) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(skill);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const formattedDate = Utilities.formatDate(new Date(tanggal), Session.getScriptTimeZone(), "dd/MM");
  let colIndex = header.indexOf(formattedDate);
  
  if (colIndex === -1) {
    colIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, colIndex).setValue(formattedDate);
    // Atur format tanggal agar rapi
    sheet.getRange(1, colIndex).setNumberFormat('@');
  } else {
    colIndex += 1; // Konversi ke indeks 1-based
  }
  
  dataAbsen.forEach(item => {
    sheet.getRange(item.rowIndex, colIndex).setValue(item.status);
  });
  
  return true;
}
