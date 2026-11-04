/**
 * BACK END UCAPAN & KONFIRMASI — Undangan Ratri & Gallang
 * Google Apps Script (Extensions/Ekstensi → Apps Script dari Google Sheet Anda)
 *
 * Kolom tab "Ucapan":
 *   A Waktu | B ID | C Nama | D Kehadiran | E Ucapan | F Diundang ke | G Status
 * Kolom F terisi otomatis dari tautan tamu: &jenis=gereja → Gereja,
 * &jenis=resepsi → Selaras, tanpa jenis → Lengkap.
 * Isi kolom G dengan "sembunyi" untuk menyembunyikan ucapan tertentu dari undangan.
 *
 * Fitur kecepatan & keamanan:
 *  - Daftar ucapan disimpan di CacheService → dibaca dari memori, bukan dari Sheet.
 *  - Setiap kiriman membawa ID unik; ID yang sama tidak akan dicatat dua kali.
 *  - Pengaman tambahan: nama + ucapan yang sama dalam 10 menit dianggap ganda.
 *  - LockService mencegah dua kiriman bersamaan saling tumpang-tindih.
 */

const SHEET_NAME = 'Ucapan';
const CACHE_KEY  = 'wishes_v2';
const CACHE_TTL  = 300;   // detik (5 menit); otomatis diperbarui setiap ada ucapan baru
const MAX_LIST   = 80;    // jumlah ucapan terbaru yang dikirim ke undangan

/* ---------------- Titik masuk ---------------- */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'ping') return json_({ ok: true });
  const data = getList_();
  return json_({ ok: true, items: data.items, stats: data.stats });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Format data tidak valid.' });
  }

  const id     = clean_(body.id, 64);
  const nama   = clean_(body.nama, 60).replace(/\s+/g, ' ');
  const hadir  = /tidak/i.test(String(body.hadir || '')) ? 'Tidak hadir' : 'Hadir';
  const ucapan = clean_(body.ucapan, 500, true);
  const jenis  = String(body.diundang || '').toLowerCase();
  const diundang = jenis === 'gereja' ? 'Gereja' : jenis === 'resepsi' ? 'Selaras' : 'Lengkap';
  if (!id || !nama) return json_({ ok: false, error: 'Nama wajib diisi.' });

  const cache = CacheService.getScriptCache();

  // Jalur tercepat: ID ini baru saja tercatat → langsung jawab "sudah diterima"
  if (cache.get('id_' + id)) return json_({ ok: true, duplicate: true });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return json_({ ok: false, error: 'Server sedang sibuk. Silakan tekan Kirim sekali lagi.' });
  }

  try {
    const sh = sheet_();
    const last = sh.getLastRow();

    if (last > 1) {
      // 1) Cek ID unik di kolom B
      const found = sh.getRange(2, 2, last - 1, 1)
        .createTextFinder(id).matchEntireCell(true).findNext();
      if (found) {
        cache.put('id_' + id, '1', 21600);
        return json_({ ok: true, duplicate: true });
      }
      // 2) Cek isi yang sama dalam 10 menit terakhir (untuk tautan versi lama)
      const n = Math.min(20, last - 1);
      const rows = sh.getRange(last - n + 1, 1, n, 5).getValues();
      const now = Date.now();
      const key = (nama + '|' + ucapan).toLowerCase();
      for (const r of rows) {
        const same = (String(r[2]) + '|' + String(r[4])).toLowerCase() === key;
        if (same && now - new Date(r[0]).getTime() < 10 * 60 * 1000) {
          return json_({ ok: true, duplicate: true });
        }
      }
    }

    const waktu = new Date();
    sh.appendRow([waktu, id, safe_(nama), hadir, safe_(ucapan), diundang, 'tampil']);
    SpreadsheetApp.flush();
    cache.put('id_' + id, '1', 21600);
    cache.remove(CACHE_KEY);

    return json_({
      ok: true,
      item: { id: id, nama: nama, hadir: hadir, ucapan: ucapan, diundang: diundang, waktu: waktu.toISOString() }
    });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Data ---------------- */

function getList_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const sh = sheet_();
  const last = sh.getLastRow();
  let items = [];
  const stats = { hadir: 0, tidak: 0 };
  if (last > 1) {
    // Jumlah Hadir / Tidak Hadir dihitung dari seluruh baris (kolom C, D, G)
    sh.getRange(2, 3, last - 1, 5).getValues().forEach(r => {
      if (!r[0] || String(r[4]).toLowerCase() === 'sembunyi') return;
      if (/tidak/i.test(String(r[1]))) stats.tidak++; else stats.hadir++;
    });
    const n = Math.min(MAX_LIST * 2, last - 1);  // ambil lebih untuk mengimbangi yang disembunyikan
    const rows = sh.getRange(last - n + 1, 1, n, 7).getValues();
    items = rows
      .filter(r => r[2] && String(r[6]).toLowerCase() !== 'sembunyi')
      .map(r => ({
        id: String(r[1]),
        nama: String(r[2]),
        hadir: /tidak/i.test(String(r[3])) ? 'Tidak hadir' : 'Hadir',
        ucapan: String(r[4]),
        waktu: r[0] instanceof Date ? r[0].toISOString() : String(r[0])
      }))
      .reverse()
      .slice(0, MAX_LIST);
  }
  const data = { items: items, stats: stats };
  try { cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL); } catch (e) { /* melebihi 100 KB: lewati cache */ }
  return data;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Waktu', 'ID', 'Nama', 'Kehadiran', 'Ucapan', 'Diundang ke', 'Status']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------------- Pembantu ---------------- */

function clean_(v, max, keepNewline) {
  let s = String(v == null ? '' : v);
  s = keepNewline ? s.replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, ' ')
                  : s.replace(/[\u0000-\u001F\u007F]/g, ' ');
  return s.trim().slice(0, max);
}

// Mencegah teks diperlakukan sebagai rumus di Google Sheet
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Utilitas (jalankan manual dari editor) ---------------- */

// Bila Anda mengubah/menyembunyikan ucapan langsung di Sheet, cache ikut dikosongkan.
function onEdit() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// Membersihkan ucapan ganda yang SUDAH terlanjur tercatat (nama + ucapan sama).
// Pilih fungsi ini di editor Apps Script, lalu tekan "Jalankan". Baris pertama yang dipertahankan.
function hapusUcapanGanda() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 3) return;
  const rows = sh.getRange(2, 1, last - 1, 5).getValues();
  const seen = {};
  const toDelete = [];
  rows.forEach((r, i) => {
    const key = (String(r[2]).trim() + '|' + String(r[4]).trim()).toLowerCase();
    if (seen[key]) toDelete.push(i + 2); else seen[key] = true;
  });
  for (let i = toDelete.length - 1; i >= 0; i--) sh.deleteRow(toDelete[i]);
  CacheService.getScriptCache().remove(CACHE_KEY);
  Logger.log(toDelete.length + ' baris ganda dihapus.');
}
