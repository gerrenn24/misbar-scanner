/**
 * MISBAR 2027 - Gate & Konsumsi Check-In Backend (Google Apps Script)
 * Sinkronisasi 2 arah antara:
 *   1. Form Responden Utama (Sheet Responden) -> Kolom "Check in", "Check in Konsum", & "Status"
 *   2. 6 Sheet Paroki: Gabriel, Martinus, Paulus, Bubat, Dayeuhkolot, Mikael
 *      - Kolom A (1): ID Tiket
 *      - Kolom B (2): Nama
 *      - Kolom C (3): Check in (Masuk Gerbang)
 *      - Kolom D (4): Check in Konsum (Pengambilan Konsumsi)
 */

const REQUIRE_PIN = false; // false = Fitur PIN Panitia DIMATIKAN | true = PIN Panitia AKTIF
const GATE_PIN = '2027';   // Hanya dipakai jika REQUIRE_PIN = true
const TIMEZONE = 'Asia/Jakarta';
const PAROKI_SHEETS = ['Gabriel', 'Martinus', 'Paulus', 'Bubat', 'Dayeuhkolot', 'Mikael'];
const PAROKI_HEADERS = ['ID Tiket', 'Nama', 'Check in', 'Check in Konsum'];

// Teks & Warna Status Check-In di Form Responden Utama
const STATUS_CHECKED_IN = 'Sudah Check In';
const STATUS_NOT_CHECKED_IN = 'Belum Check In';
const COLOR_GREEN_BG = '#bbf7d0';   // Hijau lembut
const COLOR_GREEN_TEXT = '#14532d'; // Hijau tua
const COLOR_RED_BG = '#fecaca';     // Merah lembut
const COLOR_RED_TEXT = '#7f1d1d';   // Merah tua

/**
 * Menu khusus di Google Sheets untuk sinkronisasi manual 1 klik
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MISBAR 2027')
    .addItem('1. Bangun & Sinkronisasi 6 Sheet Paroki + Responden Utama', 'buildAndSyncParokiSheets')
    .toUi();
}

/**
 * Jalankan fungsi ini dari editor Apps Script (atau dari menu MISBAR 2027)
 * untuk membangun/menyamakan data peserta, waktu Check in, serta kolom Status (Hijau/Merah)
 * di Form Responden Utama dan di ke-6 Sheet Paroki.
 */
function buildAndSyncParokiSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureParokiSheetsExist_(ss);
  const syncedCount = syncAllSheetsTwoWay_(ss);

  try {
    SpreadsheetApp.getUi().alert(
      'Sinkronisasi Berhasil!\n\n' +
      '6 Sheet Paroki (' + PAROKI_SHEETS.join(', ') + ') dan Form Responden Utama sudah tersinkronisasi.\n' +
      'Kolom Status (Hijau = Sudah Check In, Merah = Belum Check In) juga telah diperbarui.\n' +
      'Total data peserta diproses: ' + syncedCount + ' baris.'
    );
  } catch (e) {
    Logger.log('Selesai sinkronisasi 2 arah. Total diproses: ' + syncedCount);
  }
}

/**
 * Trigger otomatis saat ada pendaftar baru masuk dari Google Form
 * (Pasang trigger: From spreadsheet -> On form submit -> onFormSubmitSync)
 */
function onFormSubmitSync(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureParokiSheetsExist_(ss);
  syncAllSheetsTwoWay_(ss);
}

/**
 * Memastikan ke-6 sheet paroki ada beserta header Kolom A-D
 */
function ensureParokiSheetsExist_(ss) {
  PAROKI_SHEETS.forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const firstRow = sheet.getRange(1, 1, 1, PAROKI_HEADERS.length).getValues()[0];
    const needsHeader = PAROKI_HEADERS.some(function (h, idx) {
      return String(firstRow[idx] || '').trim() !== h;
    });
    if (needsHeader) {
      sheet.getRange(1, 1, 1, PAROKI_HEADERS.length).setValues([PAROKI_HEADERS]);
      sheet.getRange(1, 1, 1, PAROKI_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });
}

/**
 * Mencari sheet Form Responden Utama dan memastikan kolom "Check in", "Check in Konsum",
 * serta kolom "Status" tersedia.
 */
function getMasterSheetMeta_(ss, createMissingCols) {
  const allSheets = ss.getSheets();

  for (let s = 0; s < allSheets.length; s++) {
    const sh = allSheets[s];
    if (PAROKI_SHEETS.indexOf(sh.getName()) !== -1) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 2) continue;

    let headersRaw = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    let headers = headersRaw.map(function (h) { return String(h || '').toLowerCase().trim(); });

    const colTicket = headers.findIndex(function (h) {
      return h.indexOf('ticket') !== -1 || h.indexOf('tiket') !== -1 || h === 'id';
    });
    if (colTicket === -1) continue;

    const colNama = headers.findIndex(function (h) {
      return h.indexOf('nama') !== -1;
    });
    const colParoki = headers.findIndex(function (h) {
      return h.indexOf('paroki') !== -1 || h.indexOf('asal') !== -1;
    });

    // Deteksi kolom Status Check-In (misal: "Status", "Status Check in", "Status Cek In")
    let colStatus = headers.findIndex(function (h) {
      return h.indexOf('status') !== -1 && h.indexOf('konsum') === -1;
    });

    // Deteksi kolom Waktu Check-In Masuk (mengecualikan kolom "status" dan "konsum")
    let colMasuk = headers.findIndex(function (h) {
      const isCheckIn = h.indexOf('check in') !== -1 || h.indexOf('checkin') !== -1 ||
                        h.indexOf('cek in') !== -1 || h.indexOf('cekin') !== -1;
      return isCheckIn && h.indexOf('konsum') === -1 && h.indexOf('status') === -1;
    });

    // Deteksi kolom Waktu Check-In Konsumsi
    let colKonsum = headers.findIndex(function (h) {
      return h.indexOf('konsum') !== -1 && h.indexOf('status') === -1;
    });

    if (createMissingCols) {
      let nextCol = sh.getLastColumn() + 1;
      if (colStatus === -1) {
        sh.getRange(1, nextCol).setValue('Status Check in').setFontWeight('bold');
        colStatus = nextCol - 1;
        nextCol++;
      }
      if (colMasuk === -1) {
        sh.getRange(1, nextCol).setValue('Check in').setFontWeight('bold');
        colMasuk = nextCol - 1;
        nextCol++;
      }
      if (colKonsum === -1) {
        sh.getRange(1, nextCol).setValue('Check in Konsum').setFontWeight('bold');
        colKonsum = nextCol - 1;
        nextCol++;
      }
    }

    return {
      sheet: sh,
      colTicket: colTicket,
      colNama: colNama,
      colParoki: colParoki,
      colStatus: colStatus,
      colMasuk: colMasuk,
      colKonsum: colKonsum
    };
  }

  return null;
}

/**
 * Mengatur teks dan warna cell Status Check-In (Hijau = Sudah Check In, Merah = Belum Check In)
 */
function applyMasterStatusCell_(cell, isCheckedIn) {
  if (!cell) return;
  if (isCheckedIn) {
    cell.setValue(STATUS_CHECKED_IN)
      .setBackground(COLOR_GREEN_BG)
      .setFontColor(COLOR_GREEN_TEXT)
      .setFontWeight('bold');
  } else {
    cell.setValue(STATUS_NOT_CHECKED_IN)
      .setBackground(COLOR_RED_BG)
      .setFontColor(COLOR_RED_TEXT)
      .setFontWeight('bold');
  }
}

/**
 * Memetakan teks nama paroki dari Google Form ke salah satu dari 6 nama tab sheet
 */
function resolveParokiSheetName_(rawParoki) {
  const s = String(rawParoki || '').toLowerCase().trim();
  if (!s) return null;
  if (s.indexOf('gabriel') !== -1) return 'Gabriel';
  if (s.indexOf('martinus') !== -1) return 'Martinus';
  if (s.indexOf('paulus') !== -1) return 'Paulus';
  if (s.indexOf('bubat') !== -1 || s.indexOf('buah batu') !== -1 || s.indexOf('buahbatu') !== -1) return 'Bubat';
  if (s.indexOf('dayeuh') !== -1) return 'Dayeuhkolot';
  if (s.indexOf('mikael') !== -1) return 'Mikael';
  return null;
}

/**
 * Sinkronisasi 2 arah antara Form Responden Utama dan 6 Sheet Paroki:
 * - Peserta baru di Form Responden Utama disalin ke Sheet Paroki yang sesuai.
 * - Status Check in (Masuk) & Check in Konsum saling disamakan.
 * - Kolom Status di Form Responden Utama otomatis diwarnai Hijau ("Sudah Check In")
 *   jika ada waktu check-in, atau Merah ("Belum Check In") jika belum check-in.
 */
function syncAllSheetsTwoWay_(ss) {
  const parokiMap = {};

  PAROKI_SHEETS.forEach(function (sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    data.forEach(function (r, idx) {
      const tid = String(r[0] || '').trim().toUpperCase();
      if (tid) {
        parokiMap[tid] = {
          sheetName: sheetName,
          rowNum: idx + 2,
          masuk: formatCellTime_(r[2]),
          konsum: formatCellTime_(r[3])
        };
      }
    });
  });

  const masterMeta = getMasterSheetMeta_(ss, true);
  if (!masterMeta) return 0;

  const masterSheet = masterMeta.sheet;
  const lastRow = masterSheet.getLastRow();
  const lastCol = masterSheet.getLastColumn();
  if (lastRow < 2) return 0;

  const numRows = lastRow - 1;
  const values = masterSheet.getRange(2, 1, numRows, lastCol).getValues();

  // Siapkan array batch untuk update kolom Status secara cepat
  const statusRange = masterSheet.getRange(2, masterMeta.colStatus + 1, numRows, 1);
  const statusValues = statusRange.getValues();
  const statusBgs = statusRange.getBackgrounds();
  const statusFontColors = statusRange.getFontColors();
  const statusFontWeights = statusRange.getFontWeights();

  let processed = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const masterRowNum = i + 2;
    const tid = String(row[masterMeta.colTicket] || '').trim().toUpperCase();
    if (!tid) continue;

    const nama = masterMeta.colNama !== -1 ? String(row[masterMeta.colNama] || '').trim() : '-';
    const masterMasuk = masterMeta.colMasuk !== -1 ? formatCellTime_(row[masterMeta.colMasuk]) : '';
    const masterKonsum = masterMeta.colKonsum !== -1 ? formatCellTime_(row[masterMeta.colKonsum]) : '';

    let finalMasuk = masterMasuk;
    let finalKonsum = masterKonsum;

    const targetSheetName = masterMeta.colParoki !== -1
      ? resolveParokiSheetName_(row[masterMeta.colParoki])
      : null;

    if (targetSheetName) {
      const targetSheet = ss.getSheetByName(targetSheetName);

      if (!parokiMap[tid]) {
        targetSheet.appendRow([tid, nama, masterMasuk, masterKonsum]);
        parokiMap[tid] = {
          sheetName: targetSheetName,
          rowNum: targetSheet.getLastRow(),
          masuk: masterMasuk,
          konsum: masterKonsum
        };
      } else {
        const info = parokiMap[tid];
        if (info.sheetName === targetSheetName) {
          finalMasuk = info.masuk || masterMasuk;
          finalKonsum = info.konsum || masterKonsum;

          // Update ke Sheet Paroki jika sebelumnya kosong
          if (!info.masuk && finalMasuk) {
            targetSheet.getRange(info.rowNum, 3).setValue(finalMasuk);
            info.masuk = finalMasuk;
          }
          if (!info.konsum && finalKonsum) {
            targetSheet.getRange(info.rowNum, 4).setValue(finalKonsum);
            info.konsum = finalKonsum;
          }

          // Update balik ke Form Responden Utama jika sebelumnya kosong
          if (!masterMasuk && finalMasuk && masterMeta.colMasuk !== -1) {
            masterSheet.getRange(masterRowNum, masterMeta.colMasuk + 1).setValue(finalMasuk);
          }
          if (!masterKonsum && finalKonsum && masterMeta.colKonsum !== -1) {
            masterSheet.getRange(masterRowNum, masterMeta.colKonsum + 1).setValue(finalKonsum);
          }
        }
      }
    }

    // Set nilai & warna kolom Status berdasarkan ada tidaknya waktu Check in
    const isCheckedIn = Boolean(finalMasuk);
    statusValues[i][0] = isCheckedIn ? STATUS_CHECKED_IN : STATUS_NOT_CHECKED_IN;
    statusBgs[i][0] = isCheckedIn ? COLOR_GREEN_BG : COLOR_RED_BG;
    statusFontColors[i][0] = isCheckedIn ? COLOR_GREEN_TEXT : COLOR_RED_TEXT;
    statusFontWeights[i][0] = 'bold';

    processed++;
  }

  // Terapkan semua perubahan kolom Status sekaligus (batch)
  statusRange.setValues(statusValues);
  statusRange.setBackgrounds(statusBgs);
  statusRange.setFontColors(statusFontColors);
  statusRange.setFontWeights(statusFontWeights);

  SpreadsheetApp.flush();
  return processed;
}

/**
 * Menulis timestamp Check in / Check in Konsum serta memperbarui kolom Status
 * (Hijau = Sudah Check In) pada baris peserta di Form Responden Utama
 */
function updateMasterSheetCheckIn_(ss, cleanId, scanMode, timestampStr) {
  const masterMeta = getMasterSheetMeta_(ss, true);
  if (!masterMeta) return;

  const masterSheet = masterMeta.sheet;
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;

  const ticketRange = masterSheet.getRange(2, masterMeta.colTicket + 1, lastRow - 1, 1).getValues();
  const targetColOneBased = (scanMode === 'konsum')
    ? (masterMeta.colKonsum + 1)
    : (masterMeta.colMasuk + 1);

  for (let i = 0; i < ticketRange.length; i++) {
    const rowId = String(ticketRange[i][0] || '').trim().toUpperCase();
    if (rowId === cleanId) {
      const rowNum = i + 2;
      const cell = masterSheet.getRange(rowNum, targetColOneBased);
      const currentVal = formatCellTime_(cell.getValue());
      if (!currentVal) {
        cell.setValue(timestampStr);
      }

      // Jika mode masuk, langsung ubah kolom Status jadi "Sudah Check In" & warna Hijau
      if (scanMode === 'masuk' && masterMeta.colStatus !== -1) {
        const statusCell = masterSheet.getRange(rowNum, masterMeta.colStatus + 1);
        applyMasterStatusCell_(statusCell, true);
      }
      break;
    }
  }
}

function formatCellTime_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  }
  return String(val).trim();
}

/**
 * Web App Endpoint
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = String(params.action || '').trim().toLowerCase();

  if (action === 'stats') {
    return jsonResponse_(getCheckInStats(params.pin));
  }

  if (action === 'checkin') {
    const mode = String(params.mode || 'masuk').trim().toLowerCase();
    return jsonResponse_(processCheckIn(params.ticketId, params.pin, mode));
  }

  try {
    return HtmlService.createHtmlOutputFromFile('Scanner')
      .setTitle('Gate & Konsumsi Scanner - MISBAR 2027')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return jsonResponse_({
      ok: true,
      requirePin: REQUIRE_PIN,
      message: 'MISBAR 2027 API Aktif.'
    });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyPin_(pin) {
  if (!REQUIRE_PIN) return true;
  return String(pin || '').trim() === GATE_PIN;
}

/**
 * Statistik keseluruhan + rekap per 6 sheet paroki
 */
function getCheckInStats(pin) {
  if (!verifyPin_(pin)) {
    return { ok: false, message: 'PIN Panitia salah.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureParokiSheetsExist_(ss);

  let total = 0;
  let checkedIn = 0;
  let konsumCheckedIn = 0;
  const perParoki = {};

  PAROKI_SHEETS.forEach(function (sheetName) {
    const sh = ss.getSheetByName(sheetName);
    let regCount = 0;
    let masukCount = 0;
    let konsumCount = 0;

    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
      rows.forEach(function (r) {
        const tid = String(r[0] || '').trim();
        if (!tid) return;
        regCount++;
        if (String(r[2] || '').trim() !== '') masukCount++;
        if (String(r[3] || '').trim() !== '') konsumCount++;
      });
    }

    perParoki[sheetName] = {
      registered: regCount,
      masuk: masukCount,
      konsum: konsumCount
    };

    total += regCount;
    checkedIn += masukCount;
    konsumCheckedIn += konsumCount;
  });

  return {
    ok: true,
    requirePin: REQUIRE_PIN,
    stats: {
      total: total,
      checkedIn: checkedIn,
      remaining: Math.max(0, total - checkedIn),
      konsumCheckedIn: konsumCheckedIn,
      perParoki: perParoki
    }
  };
}

/**
 * Check-In terpisah & sinkron ke 2 tempat sekaligus (Sheet Paroki + Form Responden Utama):
 * - mode === 'masuk'  -> Kolom C di Sheet Paroki + Kolom "Check in" & "Status" (Hijau) di Form Responden Utama
 * - mode === 'konsum' -> Kolom D di Sheet Paroki + Kolom "Check in Konsum" di Form Responden Utama
 */
function processCheckIn(ticketId, pin, mode) {
  if (!verifyPin_(pin)) {
    return { ok: false, status: 'INVALID', message: 'PIN Panitia tidak valid.' };
  }

  const cleanId = String(ticketId || '').trim().toUpperCase();
  const scanMode = (String(mode || 'masuk').trim().toLowerCase() === 'konsum') ? 'konsum' : 'masuk';

  if (!cleanId) {
    return { ok: false, status: 'INVALID', mode: scanMode, message: 'ID Tiket kosong.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureParokiSheetsExist_(ss);

    let found = findTicketInParokiSheets_(ss, cleanId);
    if (!found) {
      syncAllSheetsTwoWay_(ss);
      found = findTicketInParokiSheets_(ss, cleanId);
    }

    if (!found) {
      const statsRes = getCheckInStats(pin);
      return {
        ok: false,
        status: 'INVALID',
        mode: scanMode,
        message: 'Tiket "' + cleanId + '" tidak terdaftar di database paroki.',
        stats: statsRes.stats
      };
    }

    const { sheet, sheetName, rowNum, nama, existingMasuk, existingKonsum } = found;
    const nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    if (scanMode === 'konsum') {
      if (existingKonsum !== '') {
        updateMasterSheetCheckIn_(ss, cleanId, 'konsum', existingKonsum);
        SpreadsheetApp.flush();

        const statsRes = getCheckInStats(pin);
        return {
          ok: true,
          status: 'DUPLICATE',
          mode: 'konsum',
          message: 'Konsumsi sudah diambil sebelumnya pada ' + existingKonsum,
          participant: {
            ticketId: cleanId,
            nama: nama,
            paroki: sheetName,
            checkinTime: existingMasuk || '-',
            konsumTime: existingKonsum
          },
          stats: statsRes.stats
        };
      }

      // 1. Tulis ke Kolom D di Sheet Paroki
      sheet.getRange(rowNum, 4).setValue(nowStr);
      // 2. Tulis ke Kolom "Check in Konsum" di Form Responden Utama
      updateMasterSheetCheckIn_(ss, cleanId, 'konsum', nowStr);
      SpreadsheetApp.flush();

      const statsRes = getCheckInStats(pin);
      return {
        ok: true,
        status: 'SUCCESS',
        mode: 'konsum',
        message: 'Check-In Konsumsi tercatat di ' + sheetName + ' & Form Responden!',
        participant: {
          ticketId: cleanId,
          nama: nama,
          paroki: sheetName,
          checkinTime: existingMasuk || '-',
          konsumTime: nowStr
        },
        stats: statsRes.stats
      };
    } else {
      if (existingMasuk !== '') {
        updateMasterSheetCheckIn_(ss, cleanId, 'masuk', existingMasuk);
        SpreadsheetApp.flush();

        const statsRes = getCheckInStats(pin);
        return {
          ok: true,
          status: 'DUPLICATE',
          mode: 'masuk',
          message: 'Tiket sudah dipakai masuk sebelumnya pada ' + existingMasuk,
          participant: {
            ticketId: cleanId,
            nama: nama,
            paroki: sheetName,
            checkinTime: existingMasuk,
            konsumTime: existingKonsum || '-'
          },
          stats: statsRes.stats
        };
      }

      // 1. Tulis ke Kolom C di Sheet Paroki
      sheet.getRange(rowNum, 3).setValue(nowStr);
      // 2. Tulis ke Kolom "Check in" & ubah Status jadi Hijau di Form Responden Utama
      updateMasterSheetCheckIn_(ss, cleanId, 'masuk', nowStr);
      SpreadsheetApp.flush();

      const statsRes = getCheckInStats(pin);
      return {
        ok: true,
        status: 'SUCCESS',
        mode: 'masuk',
        message: 'Check-In Masuk tercatat di ' + sheetName + ' & Form Responden!',
        participant: {
          ticketId: cleanId,
          nama: nama,
          paroki: sheetName,
          checkinTime: nowStr,
          konsumTime: existingKonsum || '-'
        },
        stats: statsRes.stats
      };
    }
  } finally {
    lock.releaseLock();
  }
}

function findTicketInParokiSheets_(ss, cleanId) {
  for (let i = 0; i < PAROKI_SHEETS.length; i++) {
    const sheetName = PAROKI_SHEETS[i];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) continue;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    for (let r = 0; r < data.length; r++) {
      const rowId = String(data[r][0] || '').trim().toUpperCase();
      if (rowId === cleanId) {
        return {
          sheet: sheet,
          sheetName: sheetName,
          rowNum: r + 2,
          nama: String(data[r][1] || '-').trim(),
          existingMasuk: formatCellTime_(data[r][2]),
          existingKonsum: formatCellTime_(data[r][3])
        };
      }
    }
  }
  return null;
}
