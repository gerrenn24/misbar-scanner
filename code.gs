/**
 * File: Code.gs
 * Spreadsheet: cloningan (Responses) - Tab: Form Responses 1
 * Mapping Kolom:
 *   A (1): Timestamp
 *   B (2): Email
 *   C (3): Nama
 *   D (4): No Telepon
 *   E (5): Usia
 *   F (6): Paroki
 *   G (7): Pergi dengan
 *   H (8): Jika kendaraan pribadi (mobil-motor)
 *   I (9): ID Tiket (Otomatis)
 *   J (10): Status Check-in (Otomatis)
 *   K (11): Waktu Check-in (Diisi saat scan di gate)
 */

const CONFIG = {
  SHEET_NAME: 'Form Responses 1',
  COL_EMAIL: 2,          // Kolom B
  COL_NAMA: 3,           // Kolom C
  COL_USIA: 5,           // Kolom E
  COL_PAROKI: 6,         // Kolom F
  COL_TRANSPORT: 7,      // Kolom G
  COL_VEHICLE: 8,        // Kolom H
  COL_TICKET_ID: 9,      // Kolom I
  COL_STATUS: 10,        // Kolom J
  COL_CHECKIN_TIME: 11,  // Kolom K
  EVENT_PREFIX: 'MISBAR27',
  EMAIL_SUBJECT: 'E-Ticket Check-In - Misa Akbar 2027 St. Gabriel',
  PIN_PANITIA: '2027'    // Ubah PIN akses scanner di sini jika diperlukan
};

/**
 * Menangani request POST dari file HTML eksternal (fetch API)
 */
function doPost(e) {
  try {
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = String(payload.action || '').trim();
    const pin = String(payload.pin || '').trim();

    let result;
    if (action === 'stats') {
      result = getCheckInStats(pin);
    } else if (action === 'checkin') {
      result = processCheckIn(payload.ticketId, pin);
    } else {
      result = { ok: false, status: 'INVALID', message: 'Action tidak dikenali.' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      status: 'ERROR',
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Menangani akses langsung ke URL Web App atau request GET
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  if (params.action) {
    let result;
    if (params.action === 'stats') {
      result = getCheckInStats(params.pin);
    } else if (params.action === 'checkin') {
      result = processCheckIn(params.ticketId, params.pin);
    } else {
      result = { ok: false, status: 'INVALID', message: 'Action tidak dikenali.' };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Scanner')
    .setTitle('Gate Scanner - MISBAR 2027')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Mengambil statistik total pendaftar dan jumlah peserta yang sudah check-in
 */
function getCheckInStats(pin) {
  if (String(pin || '').trim() !== CONFIG.PIN_PANITIA) {
    return { ok: false, message: 'PIN Panitia salah.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    return { ok: false, message: 'Sheet "' + CONFIG.SHEET_NAME + '" tidak ditemukan.' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { ok: true, stats: { total: 0, checkedIn: 0, remaining: 0 } };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_CHECKIN_TIME).getValues();
  return {
    ok: true,
    stats: computeStatsFromData_(data)
  };
}

/**
 * Memvalidasi QR Code / ID Tiket dan mencatat waktu check-in ke Google Sheets
 */
function processCheckIn(rawTicketId, pin) {
  if (String(pin || '').trim() !== CONFIG.PIN_PANITIA) {
    return { status: 'UNAUTHORIZED', message: 'PIN Panitia salah.' };
  }

  const ticketId = String(rawTicketId || '').trim().toUpperCase();
  if (!ticketId) {
    return { status: 'INVALID', message: 'Kode QR tidak terbaca atau kosong.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      return { status: 'ERROR', message: 'Tab "' + CONFIG.SHEET_NAME + '" tidak ditemukan.' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: 'INVALID', message: 'Belum ada data peserta terdaftar.' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_CHECKIN_TIME).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowTicketId = String(data[i][CONFIG.COL_TICKET_ID - 1] || '').trim().toUpperCase();

      if (rowTicketId === ticketId) {
        const sheetRow = i + 2;
        const nama = String(data[i][CONFIG.COL_NAMA - 1] || '-').trim();
        const usia = String(data[i][CONFIG.COL_USIA - 1] || '-').trim();
        const paroki = String(data[i][CONFIG.COL_PAROKI - 1] || '-').trim();
        const transport = String(data[i][CONFIG.COL_TRANSPORT - 1] || '').trim();
        const vehicle = String(data[i][CONFIG.COL_VEHICLE - 1] || '').trim();
        const infoKendaraan = formatTransportInfo_(transport, vehicle);

        const currentStatus = String(data[i][CONFIG.COL_STATUS - 1] || '').trim();
        const prevTimeRaw = data[i][CONFIG.COL_CHECKIN_TIME - 1];
        const prevTime = (prevTimeRaw instanceof Date)
          ? Utilities.formatDate(prevTimeRaw, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss')
          : String(prevTimeRaw || '-');

        // Tolak jika tiket sudah pernah dipakai check-in sebelumnya
        if (currentStatus === 'Sudah Check-in') {
          return {
            status: 'DUPLICATE',
            message: 'Tiket sudah dipakai check-in!',
            participant: {
              ticketId: rowTicketId,
              nama: nama,
              usia: usia,
              paroki: paroki,
              kendaraan: infoKendaraan,
              checkinTime: prevTime
            },
            stats: computeStatsFromData_(data)
          };
        }

        // Tandai tiket sebagai Sudah Check-in
        const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
        sheet.getRange(sheetRow, CONFIG.COL_STATUS).setValue('Sudah Check-in');
        sheet.getRange(sheetRow, CONFIG.COL_CHECKIN_TIME).setValue(nowStr);
        SpreadsheetApp.flush();

        data[i][CONFIG.COL_STATUS - 1] = 'Sudah Check-in';

        return {
          status: 'SUCCESS',
          message: 'Check-In Berhasil!',
          participant: {
            ticketId: rowTicketId,
            nama: nama,
            usia: usia,
            paroki: paroki,
            kendaraan: infoKendaraan,
            checkinTime: nowStr
          },
          stats: computeStatsFromData_(data)
        };
      }
    }

    return {
      status: 'INVALID',
      message: 'ID Tiket "' + ticketId + '" tidak terdaftar.'
    };

  } catch (err) {
    return {
      status: 'ERROR',
      message: 'Terjadi kesalahan server: ' + err.message
    };
  } finally {
    lock.releaseLock();
  }
}

function computeStatsFromData_(data) {
  let total = 0;
  let checkedIn = 0;
  for (let i = 0; i < data.length; i++) {
    const tid = String(data[i][CONFIG.COL_TICKET_ID - 1] || '').trim();
    if (tid) {
      total++;
      if (String(data[i][CONFIG.COL_STATUS - 1] || '').trim() === 'Sudah Check-in') {
        checkedIn++;
      }
    }
  }
  return {
    total: total,
    checkedIn: checkedIn,
    remaining: Math.max(0, total - checkedIn)
  };
}

/**
 * Trigger otomatis saat Google Form dikirim:
 * Membuat ID Tiket, generate QR Code, mengirim email E-Ticket HTML, dan mencatat ke Sheets.
 */
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = (e && e.range) ? e.range.getSheet() : ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('Tab sheet "' + CONFIG.SHEET_NAME + '" tidak ditemukan.');
    }

    const row = (e && e.range) ? e.range.getRow() : sheet.getLastRow();
    if (row <= 1) return;

    const existingTicketId = sheet.getRange(row, CONFIG.COL_TICKET_ID).getValue();
    if (existingTicketId) return;

    const email = String(sheet.getRange(row, CONFIG.COL_EMAIL).getValue()).replace(/\s+/g, '');
    const nama = String(sheet.getRange(row, CONFIG.COL_NAMA).getValue()).trim() || 'Teman OMK';
    const usia = String(sheet.getRange(row, CONFIG.COL_USIA).getValue()).trim() || '-';
    const paroki = String(sheet.getRange(row, CONFIG.COL_PAROKI).getValue()).trim() || '-';
    const transport = String(sheet.getRange(row, CONFIG.COL_TRANSPORT).getValue()).trim();
    const vehicle = String(sheet.getRange(row, CONFIG.COL_VEHICLE).getValue()).trim();
    const infoKendaraan = formatTransportInfo_(transport, vehicle);

    if (!email || !email.includes('@')) {
      sheet.getRange(row, CONFIG.COL_STATUS).setValue('Gagal: Email tidak valid');
      return;
    }

    const sequenceNumber = String(row - 1).padStart(4, '0');
    const randomSuffix = generateRandomCode_(4);
    const ticketId = `${CONFIG.EVENT_PREFIX}-${sequenceNumber}-${randomSuffix}`;

    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=280x280&margin=10&data=' + encodeURIComponent(ticketId);
    const qrResponse = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });

    if (qrResponse.getResponseCode() !== 200) {
      throw new Error('Gagal membuat gambar QR Code. HTTP Status: ' + qrResponse.getResponseCode());
    }

    const qrBlob = qrResponse.getBlob().setName(`QR_${ticketId}.png`);

    const plainBody =
      `Halo ${nama}, terima kasih sudah melakukan Registrasi Misa Akbar 2027 di St. Gabriel!\n\n` +
      `Berikut detail E-Ticket Anda:\n` +
      `- ID Tiket: ${ticketId}\n` +
      `- Paroki: ${paroki}\n` +
      `- Usia: ${usia}\n` +
      `- Transportasi: ${infoKendaraan}\n\n` +
      `Silakan tunjukkan QR Code pada email ini saat Check-In tanggal 20 Februari 2027. Sampai jumpa di Misa Akbar 2027!`;

    const htmlBody = buildTicketEmailHtml_({
      nama: nama,
      ticketId: ticketId,
      paroki: paroki,
      usia: usia,
      kendaraan: infoKendaraan,
      qrCid: 'qrCode'
    });

    MailApp.sendEmail({
      to: email,
      subject: `${CONFIG.EMAIL_SUBJECT} [${ticketId}]`,
      body: plainBody,
      htmlBody: htmlBody,
      inlineImages: {
        qrCode: qrBlob
      }
    });

    sheet.getRange(row, CONFIG.COL_TICKET_ID).setValue(ticketId);
    sheet.getRange(row, CONFIG.COL_STATUS).setValue('Belum Check-in');

  } catch (error) {
    Logger.log('Error pada onFormSubmit: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membangun tampilan email E-Ticket HTML responsif untuk Gmail & Apple Mail
 */
function buildTicketEmailHtml_(data) {
  const nama = escapeHtml_(data.nama || 'Peserta');
  const ticketId = escapeHtml_(data.ticketId || '-');
  const paroki = escapeHtml_(data.paroki || '-');
  const usia = escapeHtml_(data.usia || '-');
  const transport = escapeHtml_(data.kendaraan || '-');
  const qrCid = data.qrCid || 'qrCode';

  return `
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>E-Ticket Misa Akbar 2027</title>
</head>
<body style="margin:0; padding:0; background-color:#f1f5f9; font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color:#0f172a;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f1f5f9; padding:24px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:520px; background-color:#ffffff; border-radius:16px; overflow:hidden; border:1px solid #cbd5e1;">
          
          <!-- Banner Header -->
          <tr>
            <td style="background:linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); background-color:#0f172a; padding:28px 24px; text-align:center;">
              <div style="display:inline-block; background-color:rgba(255,255,255,0.14); color:#93c5fd; font-size:11px; font-weight:700; letter-spacing:1.2px; text-transform:uppercase; padding:5px 12px; border-radius:999px; margin-bottom:10px;">
                E-Ticket Resmi Check-In
              </div>
              <h1 style="margin:0; font-size:22px; font-weight:800; color:#ffffff; line-height:1.3;">
                MISA AKBAR 2027
              </h1>
              <p style="margin:6px 0 0 0; font-size:13.5px; color:#cbd5e1;">
                Paroki St. Gabriel Bandung • Sabtu, 20 Februari 2027
              </p>
            </td>
          </tr>

          <!-- Sapaan Pembuka -->
          <tr>
            <td style="padding:24px 24px 12px 24px;">
              <p style="margin:0 0 10px 0; font-size:15px; line-height:1.5; color:#0f172a;">
                Halo <strong>${nama}</strong>,
              </p>
              <p style="margin:0; font-size:14px; line-height:1.6; color:#475569;">
                Terima kasih telah melakukan registrasi <strong>Misa Akbar 2027 di St. Gabriel</strong>. Simpan email ini dan tunjukkan QR Code di bawah kepada petugas di pintu masuk saat melakukan check-in.
              </p>
            </td>
          </tr>

          <!-- Kartu QR Code -->
          <tr>
            <td style="padding:12px 24px;">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f8fafc; border:2px dashed #94a3b8; border-radius:14px;">
                <tr>
                  <td align="center" style="padding:22px 16px;">
                    <div style="font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#64748b; margin-bottom:12px;">
                      Pindai QR Code Ini di Gate
                    </div>
                    <div style="background-color:#ffffff; padding:12px; border-radius:12px; border:1px solid #e2e8f0; display:inline-block;">
                      <img src="cid:${qrCid}" alt="QR Code ${ticketId}" width="200" height="200" style="display:block; width:200px; height:200px; border:0;" />
                    </div>
                    <div style="margin-top:14px; font-family:'Courier New', Courier, monospace; font-size:16px; font-weight:800; letter-spacing:1.5px; color:#0f172a; background-color:#e2e8f0; display:inline-block; padding:6px 14px; border-radius:8px;">
                      ${ticketId}
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Detail Informasi Peserta -->
          <tr>
            <td style="padding:10px 24px 20px 24px;">
              <div style="font-size:11.5px; font-weight:700; text-transform:uppercase; letter-spacing:0.8px; color:#64748b; margin-bottom:8px;">
                Detail Peserta Terdaftar
              </div>
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border:1px solid #e2e8f0; border-radius:10px; overflow:hidden; font-size:13.5px;">
                <tr>
                  <td style="padding:10px 14px; background-color:#f8fafc; color:#64748b; width:38%; border-bottom:1px solid #e2e8f0;">Nama Lengkap</td>
                  <td style="padding:10px 14px; font-weight:700; color:#0f172a; border-bottom:1px solid #e2e8f0;">${nama}</td>
                </tr>
                <tr>
                  <td style="padding:10px 14px; background-color:#f8fafc; color:#64748b; border-bottom:1px solid #e2e8f0;">Asal Paroki</td>
                  <td style="padding:10px 14px; font-weight:600; color:#0f172a; border-bottom:1px solid #e2e8f0;">${paroki}</td>
                </tr>
                <tr>
                  <td style="padding:10px 14px; background-color:#f8fafc; color:#64748b; border-bottom:1px solid #e2e8f0;">Usia</td>
                  <td style="padding:10px 14px; font-weight:600; color:#0f172a; border-bottom:1px solid #e2e8f0;">${usia}</td>
                </tr>
                <tr>
                  <td style="padding:10px 14px; background-color:#f8fafc; color:#64748b;">Transportasi</td>
                  <td style="padding:10px 14px; font-weight:600; color:#0f172a;">${transport}</td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Catatan Penting -->
          <tr>
            <td style="padding:0 24px 24px 24px;">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#eff6ff; border-left:4px solid #2563eb; border-radius:8px;">
                <tr>
                  <td style="padding:12px 14px; font-size:12.5px; line-height:1.55; color:#1e3a8a;">
                    <strong>Catatan Check-In:</strong><br />
                    • Tingkatkan kecerahan layar HP saat QR Code dipindai petugas.<br />
                    • Satu kode tiket hanya berlaku untuk 1 kali masuk.<br />
                    • Disarankan screenshot QR Code ini untuk berjaga-jaga jika sinyal internet di lokasi kurang stabil.
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Footer Email -->
          <tr>
            <td style="background-color:#f8fafc; border-top:1px solid #e2e8f0; padding:16px 24px; text-align:center; font-size:12px; color:#64748b; line-height:1.5;">
              Sampai jumpa di Misa Akbar 2027!<br />
              <strong>Panitia MISBAR • Paroki St. Gabriel Bandung</strong>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

function formatTransportInfo_(transport, vehicle) {
  if (transport && vehicle) {
    return transport + ' (' + vehicle + ')';
  }
  return transport || vehicle || '-';
}

function testKirimBarisTerakhir() {
  onFormSubmit(null);
}

function generateRandomCode_(length) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function escapeHtml_(text) {
  return String(text)
    .replace(/\u0026/g, '\u0026amp;')
    .replace(/\u003C/g, '\u0026lt;')
    .replace(/\u003E/g, '\u0026gt;')
    .replace(/\u0022/g, '\u0026quot;')
    .replace(/\u0027/g, '\u0026#39;');
}