/** SIMARISA-SAKTI LMS - Backend (Google Apps Script)
 *  Database: Google Sheets (spreadsheet tempat script ini terikat).
 *  Langkah: jalankan setup() sekali, lalu Deploy > Web app.
 */
const SCHEMA = {
  Users: ['id', 'nama', 'username', 'password', 'role'],
  Kelas: ['id', 'nama', 'deskripsi', 'guruId', 'warna'],
  Enroll: ['id', 'kelasId', 'siswaId'],
  Materi: ['id', 'kelasId', 'judul', 'konten', 'link', 'tanggal'],
  Tugas: ['id', 'kelasId', 'judul', 'deskripsi', 'deadline', 'link', 'tutup', 'jenis', 'soal', 'bagikan'], // jenis: uraian | pilgan; bagikan: ya | tidak (tugas lama dianggap dibagikan)
  Kumpul: ['id', 'tugasId', 'siswaId', 'jawaban', 'nilai', 'tanggal', 'link', 'komentar', 'rincian'], // rincian: nilai per soal esai (JSON)
  Selesai: ['id', 'materiId', 'siswaId', 'tanggal'] // penanda materi yang sudah selesai dipelajari siswa
};
const SALT = 'simarisa-sakti';
const db_ = () => SpreadsheetApp.getActiveSpreadsheet();
const ensured_ = {};
function sh_(n) {
  let sh = db_().getSheetByName(n);
  if (!sh && SCHEMA[n]) { sh = db_().insertSheet(n); sh.appendRow(SCHEMA[n]); sh.setFrozenRows(1); }
  if (sh && SCHEMA[n] && !ensured_[n]) { // tambahkan kolom baru yang belum ada (tanpa perlu setup() ulang)
    ensured_[n] = true;
    const h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(x => String(x).trim());
    SCHEMA[n].forEach(k => { if (h.indexOf(k) < 0) { h.push(k); sh.getRange(1, h.length).setValue(k).setFontWeight('bold').setBackground('#6C4DF6').setFontColor('#fff'); } });
  }
  return sh;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Simarisa-sakti LMS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Pintu API untuk frontend yang dihosting di luar Apps Script (GitHub/Cloudflare). Body: {token, a, p} */
function doPost(e) {
  let r;
  try { const b = JSON.parse(e.postData.contents); r = api(b.token, b.a, b.p); }
  catch (err) { r = { ok: false, msg: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

/** Jalankan SEKALI dari editor: membuat sheet + akun admin awal (admin / admin123) */
function setup() {
  const TEXTCOL = { Tugas: ['deadline'], Materi: ['tanggal'], Kumpul: ['tanggal'] }; // agar Sheets tidak mengubah teks tanggal
  Object.keys(SCHEMA).forEach(n => {
    const sh = db_().getSheetByName(n) || db_().insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(SCHEMA[n]);
    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    SCHEMA[n].forEach(k => { if (h.indexOf(k) < 0) { h.push(k); sh.getRange(1, h.length).setValue(k); } }); // migrasi kolom baru
    sh.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#6C4DF6').setFontColor('#fff');
    sh.setFrozenRows(1);
    (TEXTCOL[n] || []).forEach(k => sh.getRange(2, h.indexOf(k) + 1, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@'));
  });
  if (rows_('Users').length === 0)
    add_('Users', { nama: 'Administrator', username: 'admin', password: hash_('admin123'), role: 'admin' });
}

const hash_ = s => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s + SALT));
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function add_(n, o) {
  o.id = Utilities.getUuid().slice(0, 8);
  const sh = sh_(n), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(h.map(k => o[k] === undefined ? '' : o[k]));
  return o;
}
/** Ubah teks tenggat (yyyy-MM-ddTHH:mm) menjadi Date, memakai zona waktu script */
function dl_(s) {
  s = dlStr_(s); if (!s) return null;
  if (/Z$/.test(s)) return new Date(s);
  try { return Utilities.parseDate(s.slice(0, 16), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"); } catch (e) { return null; }
}
/** Seragamkan teks tenggat dari sheet (mis. "2026-11-20 0:00" atau "2026-11-20") menjadi yyyy-MM-ddTHH:mm; format ISO berakhiran Z dibiarkan */
function dlStr_(v) {
  const s = str_(v).trim(); if (!s || /Z$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2}))?/), z = n => ('0' + n).slice(-2);
  if (!m) { const d = new Date(s); return isNaN(d.getTime()) ? '' : d.toISOString(); }
  return m[1] + '-' + z(m[2]) + '-' + z(m[3]) + 'T' + (m[4] === undefined ? '23:59' : z(m[4]) + ':' + m[5]);
}
function find_(n, id) {
  const sh = sh_(n), v = sh.getDataRange().getValues(), c = v[0].indexOf('id');
  for (let i = 1; i < v.length; i++) if (String(v[i][c]) === String(id)) return { sh, row: i + 1, h: v[0] };
  throw new Error('Data tidak ditemukan');
}
function upd_(n, id, patch) {
  const f = find_(n, id);
  Object.keys(patch).forEach(k => { const c = f.h.indexOf(k); if (c > -1) f.sh.getRange(f.row, c + 1).setValue(patch[k]); });
}
function del_(n, id) { const f = find_(n, id); f.sh.deleteRow(f.row); }
/** Jumlah materi di kelas yang belum ditandai selesai oleh siswa */
function belum_(kelasId, siswaId) {
  const done = rows_('Selesai').filter(x => x.siswaId == siswaId).map(x => String(x.materiId));
  return rows_('Materi').filter(m => m.kelasId == kelasId && done.indexOf(String(m.id)) < 0).length;
}
/** ID file Drive gambar soal: hanya karakter aman, selain itu dianggap tidak ada gambar */
const imgId_ = v => /^[A-Za-z0-9_-]{10,80}$/.test(String(v || '')) ? String(v) : '';
/** Soal tersimpan sebagai JSON: [{q, o:[opsi], k:indeksKunci, m:materiId, img:idFileDrive}] */
function parseSoal_(s) {
  try { const a = JSON.parse(String(s || '[]')); return Array.isArray(a) ? a.filter(x => x && x.q).map(x => ({ q: String(x.q), o: Array.isArray(x.o) ? x.o : [], k: x.k, m: x.m || '', img: imgId_(x.img) })) : []; } catch (e) { return []; }
}
/** Jawaban esai/isian per soal tersimpan sebagai JSON {j:[jawaban1, jawaban2, ...]} */
function jwbList_(j) { try { const o = JSON.parse(String(j)); return Array.isArray(o.j) ? o.j.map(String) : []; } catch (e) { return []; } }
/** Nilai per nomor soal esai, tersimpan sebagai JSON array, mis. [20,15,'',20,20] ('' = belum dinilai) */
function rincianList_(j) { try { const a = JSON.parse(String(j)); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function ringkasEsai_(sl, j) {
  const a = jwbList_(j);
  if (!a.length) return String(j || '');
  return sl.map((x, i) => (i + 1) + '. ' + x.q + '\n   Jawaban: ' + (a[i] || '-')).join('\n');
}
const jenis_ = t => t.jenis === 'pilgan' ? 'pilgan' : 'uraian';
const dibagi_ = t => String(t.bagikan) !== 'tidak'; // tugas yang sudah dibagikan ke siswa
function hasil_(j) { try { const o = JSON.parse(String(j)); return { benar: Number(o.b) || 0, total: Number(o.n) || 0 }; } catch (e) { return { benar: 0, total: 0 }; } }
/** Folder Drive tujuan gambar soal (ID folder buatan sendiri) */
const FOLDER_GAMBAR_ID = '1RnTJdgJNqo5hCT3S5NMLVjR57nPmtTDv';
function folderGambar_() { return DriveApp.getFolderById(FOLDER_GAMBAR_ID); }
/** Jalankan SEKALI dari editor (pilih fungsi ini lalu Run) untuk menyetujui izin PENUH Google Drive.
 *  Membuat file uji kecil di folder lalu langsung memindahkannya ke sampah, agar Google meminta izin tulis. */
function izinkanDrive() {
  const f = folderGambar_();
  const t = f.createFile('tes-izin.txt', 'ok');
  t.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  t.setTrashed(true);
  Logger.log('Izin Drive penuh OK. Folder: ' + f.getName());
}
const str_ = d => (d instanceof Date ? d.toISOString() : String(d || ''));

/** Satu pintu API untuk semua aksi dari frontend */
function api(token, a, p) {
  try {
    p = p || {};
    const cache = CacheService.getScriptCache();
    if (a === 'login') {
      const u = rows_('Users').find(x => x.username === String(p.username).trim() && x.password === hash_(p.password));
      if (!u) throw new Error('Username atau password salah');
      const me = { id: u.id, nama: u.nama, username: u.username, role: u.role }, t = Utilities.getUuid();
      cache.put('t_' + t, JSON.stringify(me), 21600);
      return { ok: true, data: { token: t, user: me } };
    }
    if (a === 'register') {
      if (!p.nama || !p.username || String(p.password).length < 6) throw new Error('Lengkapi data (password min. 6 karakter)');
      if (rows_('Users').some(x => x.username === p.username.trim())) throw new Error('Username sudah dipakai');
      add_('Users', { nama: p.nama.trim(), username: p.username.trim(), password: hash_(p.password), role: 'siswa' });
      return { ok: true, data: true };
    }
    const raw = cache.get('t_' + token);
    if (!raw) throw new Error('Sesi berakhir, silakan masuk lagi');
    const me = JSON.parse(raw), staff = me.role !== 'siswa', admin = me.role === 'admin';
    const need = ok => { if (!ok) throw new Error('Anda tidak punya akses untuk aksi ini'); };
    const users = rows_('Users'), nm = id => (users.find(u => u.id == id) || {}).nama || '-';

    switch (a) {
      case 'me': return { ok: true, data: me };

      case 'kelas': {
        const en = rows_('Enroll').filter(e => e.siswaId == me.id).map(e => e.kelasId);
        let k = rows_('Kelas').map(x => ({ ...x, guru: nm(x.guruId), joined: en.includes(x.id) }));
        if (me.role === 'guru') k = k.filter(x => x.guruId == me.id);
        return { ok: true, data: k };
      }
      case 'saveKelas': need(staff);
        if (!p.nama) throw new Error('Nama kelas wajib diisi');
        if (p.id) upd_('Kelas', p.id, { nama: p.nama, deskripsi: p.deskripsi, warna: p.warna });
        else add_('Kelas', { nama: p.nama, deskripsi: p.deskripsi, guruId: me.id, warna: p.warna || 0 });
        return { ok: true, data: true };
      case 'join': need(!staff);
        if (!rows_('Enroll').some(e => e.kelasId == p.id && e.siswaId == me.id)) add_('Enroll', { kelasId: p.id, siswaId: me.id });
        return { ok: true, data: true };

      case 'detail': {
        const kum = rows_('Kumpul');
        const done = rows_('Selesai').filter(x => x.siswaId == me.id).map(x => String(x.materiId));
        const materi = rows_('Materi').filter(m => m.kelasId == p.id).map(m => ({ ...m, tanggal: str_(m.tanggal), selesai: done.indexOf(String(m.id)) > -1 }));
        const terkunci = !staff && materi.some(m => !m.selesai); // tugas terbuka setelah semua materi selesai
        const tugas = rows_('Tugas').filter(t => t.kelasId == p.id && (staff || dibagi_(t))).map(t => {
          const ks = kum.filter(k => k.tugasId == t.id);
          const mine = ks.find(k => k.siswaId == me.id);
          const dd = dl_(t.deadline), jn = jenis_(t), sl = parseSoal_(t.soal), esai = jn === 'uraian' && sl.length > 0;
          const hs = mine && jn === 'pilgan' ? hasil_(mine.jawaban) : null;
          return { ...t, jenis: jn, bagikan: dibagi_(t), nSoal: sl.length, deadline: dlStr_(t.deadline), link: terkunci ? '' : (t.link || ''), adaLink: !!t.link, tutup: t.tutup || 'tidak', total: ks.length, terkunci,
            soal: staff ? sl : ((terkunci || (jn === 'pilgan' && mine)) ? [] : sl.map(x => ({ q: x.q, o: x.o, m: x.m || '', img: x.img || '' }))), // kunci jawaban tidak pernah dikirim ke siswa
            mine: mine ? { jawaban: (jn === 'pilgan' || esai) ? '' : mine.jawaban, jwb: esai ? jwbList_(mine.jawaban) : [], rincian: esai ? rincianList_(mine.rincian) : [], link: mine.link || '', nilai: mine.nilai, komentar: mine.komentar || '', benar: hs ? hs.benar : null, jml: hs ? hs.total : null, terlambat: !!(dd && new Date(str_(mine.tanggal)) > dd) } : null };
        });
        const en = staff ? rows_('Enroll').filter(e => e.kelasId == p.id) : [];
        const anggota = en.map(e => { const u = users.find(x => x.id == e.siswaId) || {}; return { siswaId: e.siswaId, nama: u.nama || '(akun dihapus)', username: u.username || '' }; })
          .sort((a, b) => String(a.nama).localeCompare(String(b.nama)));
        return { ok: true, data: { materi, tugas, anggota } };
      }
      case 'tandai': {
        need(!staff);
        const mt = rows_('Materi').find(m => m.id == p.materiId);
        if (!mt) throw new Error('Materi tidak ditemukan');
        if (!rows_('Enroll').some(e => e.kelasId == mt.kelasId && e.siswaId == me.id)) throw new Error('Anda belum bergabung ke kelas ini');
        const ex = rows_('Selesai').find(x => x.materiId == p.materiId && x.siswaId == me.id);
        if (p.selesai && !ex) add_('Selesai', { materiId: p.materiId, siswaId: me.id, tanggal: new Date().toISOString() });
        if (!p.selesai && ex) del_('Selesai', ex.id);
        return { ok: true, data: true };
      }
      case 'saveMateri': need(staff);
        if (!p.judul) throw new Error('Judul materi wajib diisi');
        if (p.id) upd_('Materi', p.id, { judul: p.judul, konten: p.konten, link: p.link });
        else add_('Materi', { kelasId: p.kelasId, judul: p.judul, konten: p.konten, link: p.link, tanggal: new Date().toISOString() });
        return { ok: true, data: true };
      case 'saveTugas': need(staff);
        if (!p.judul) throw new Error('Judul tugas wajib diisi');
        if (p.link && !/^https?:\/\//i.test(p.link)) throw new Error('Tautan harus diawali http:// atau https://');
        const jn = p.jenis === 'pilgan' ? 'pilgan' : 'uraian';
        let soalJson = '';
        if (jn === 'pilgan') {
          const arr = Array.isArray(p.soal) ? p.soal : [];
          if (!arr.length) throw new Error('Tambahkan minimal 1 soal pilihan ganda');
          if (arr.length > 30) throw new Error('Maksimal 30 soal per tugas');
          const clean = arr.map((x, i) => {
            const o = (Array.isArray(x.o) ? x.o : []).map(v => String(v || '').trim().slice(0, 300));
            const k = Number(x.k), q = String(x.q || '').trim().slice(0, 1000);
            if (!q) throw new Error('Soal ' + (i + 1) + ': pertanyaan kosong');
            if (o.length < 2 || o.some(v => !v)) throw new Error('Soal ' + (i + 1) + ': minimal 2 pilihan, semua harus terisi');
            if (!(k >= 0 && k < o.length) || k % 1 !== 0) throw new Error('Soal ' + (i + 1) + ': kunci jawaban belum dipilih');
            return { q, o, k, m: String(x.m || ''), img: imgId_(x.img) };
          });
          soalJson = JSON.stringify(clean);
          if (soalJson.length > 45000) throw new Error('Soal terlalu panjang, kurangi jumlah atau panjang soal');
        } else { // uraian: soal esai/isian bersifat opsional
          const arr = Array.isArray(p.soal) ? p.soal : [];
          if (arr.length > 20) throw new Error('Maksimal 20 soal per tugas uraian');
          const cl = arr.map((x, i) => {
            const q = String(x.q || '').trim().slice(0, 1500);
            if (!q) throw new Error('Soal ' + (i + 1) + ': pertanyaan kosong');
            return { q, m: String(x.m || ''), img: imgId_(x.img) };
          });
          soalJson = cl.length ? JSON.stringify(cl) : '';
        }
        const tf = { judul: p.judul, deskripsi: p.deskripsi, deadline: p.deadline, link: p.link || '', tutup: p.tutup === 'ya' ? 'ya' : 'tidak', jenis: jn, soal: soalJson, bagikan: p.bagikan === 'tidak' ? 'tidak' : 'ya' };
        if (p.id) upd_('Tugas', p.id, tf);
        else add_('Tugas', { ...tf, kelasId: p.kelasId });
        return { ok: true, data: true };

      case 'uploadGambar': { // simpan gambar soal ke Google Drive (folder khusus, bisa dilihat siapa pun yang punya tautan)
        need(staff);
        const b64 = String(p.data || '');
        if (!b64) throw new Error('Data gambar kosong');
        if (b64.length > 4000000) throw new Error('Ukuran gambar terlalu besar (maks. sekitar 3 MB)');
        const bytes = Utilities.base64Decode(b64);
        const blob = Utilities.newBlob(bytes, 'image/jpeg', 'soal_' + Date.now() + '.jpg');
        const fi = folderGambar_().createFile(blob);
        fi.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return { ok: true, data: { id: fi.getId() } };
      }
      case 'rekap': { // rekap nilai satu kelas: siswa x tugas yang dibagikan
        need(staff);
        const kl = rows_('Kelas').find(k => k.id == p.kelasId);
        if (!kl) throw new Error('Kelas tidak ditemukan');
        if (me.role === 'guru' && kl.guruId != me.id) throw new Error('Anda tidak punya akses untuk aksi ini');
        const tg = rows_('Tugas').filter(t => t.kelasId == p.kelasId && dibagi_(t)), km = {};
        rows_('Kumpul').forEach(k => { km[k.tugasId + '|' + k.siswaId] = k; });
        const siswa = rows_('Enroll').filter(e => e.kelasId == p.kelasId).map(e => {
          const u = users.find(x => x.id == e.siswaId) || {}, h = {};
          tg.forEach(t => { const k = km[t.id + '|' + e.siswaId]; if (k) h[t.id] = { v: k.nilai === '' || k.nilai == null ? '' : Number(k.nilai) }; });
          return { id: e.siswaId, nama: u.nama || '(akun dihapus)', username: u.username || '', h };
        }).sort((a, b) => String(a.nama).localeCompare(String(b.nama)));
        return { ok: true, data: { kelas: kl.nama, tugas: tg.map(t => ({ id: t.id, judul: t.judul, jenis: jenis_(t) })), siswa } };
      }
      case 'siswaKelas': { // daftar siswa yang belum masuk kelas ini (untuk ditambahkan guru/admin)
        need(staff);
        const sudah = rows_('Enroll').filter(e => e.kelasId == p.kelasId).map(e => String(e.siswaId));
        return { ok: true, data: users.filter(u => u.role === 'siswa' && sudah.indexOf(String(u.id)) < 0).map(u => ({ id: u.id, nama: u.nama, username: u.username }))
          .sort((a, b) => String(a.nama).localeCompare(String(b.nama))) };
      }
      case 'tambahSiswa': {
        need(staff);
        if (!rows_('Kelas').some(k => k.id == p.kelasId)) throw new Error('Kelas tidak ditemukan');
        const ids = Array.isArray(p.siswaIds) ? p.siswaIds : [], sudah = rows_('Enroll').filter(e => e.kelasId == p.kelasId).map(e => String(e.siswaId));
        let n = 0;
        ids.forEach(id => {
          const u = users.find(x => x.id == id);
          if (u && u.role === 'siswa' && sudah.indexOf(String(id)) < 0) { add_('Enroll', { kelasId: p.kelasId, siswaId: id }); sudah.push(String(id)); n++; }
        });
        return { ok: true, data: n };
      }
      case 'keluarkan': {
        need(staff);
        const en = rows_('Enroll').find(e => e.kelasId == p.kelasId && e.siswaId == p.siswaId);
        if (!en) throw new Error('Siswa tidak ada di kelas ini');
        del_('Enroll', en.id); // nilai dan jawaban siswa tetap tersimpan di sheet Kumpul
        return { ok: true, data: true };
      }
      case 'bagikan': need(staff);
        upd_('Tugas', p.id, { bagikan: p.bagikan ? 'ya' : 'tidak' });
        return { ok: true, data: true };
      case 'submit': {
        need(!staff);
        const tg = rows_('Tugas').find(t => t.id == p.tugasId);
        if (!tg) throw new Error('Tugas tidak ditemukan');
        if (!rows_('Enroll').some(e => e.kelasId == tg.kelasId && e.siswaId == me.id)) throw new Error('Anda belum bergabung ke kelas ini');
        if (!dibagi_(tg)) throw new Error('Tugas ini belum dibagikan oleh guru');
        if (belum_(tg.kelasId, me.id) > 0) throw new Error('Tandai semua materi di kelas ini sebagai selesai terlebih dahulu');
        const dd = dl_(tg.deadline);
        if (tg.tutup === 'ya' && dd && new Date() > dd) throw new Error('Pengumpulan sudah ditutup karena melewati tenggat');
        const ex = rows_('Kumpul').find(k => k.tugasId == p.tugasId && k.siswaId == me.id), now = new Date().toISOString();
        if (jenis_(tg) === 'pilgan') { // dinilai otomatis dari kunci jawaban (kunci hanya ada di server)
          if (ex) throw new Error('Pilihan ganda hanya bisa dikerjakan satu kali');
          const sl = parseSoal_(tg.soal), jb = Array.isArray(p.jawab) ? p.jawab.map(Number) : [];
          if (!sl.length) throw new Error('Soal belum tersedia');
          if (jb.length !== sl.length || jb.some((v, i) => !(v >= 0 && v < sl[i].o.length) || v % 1 !== 0)) throw new Error('Jawab semua soal terlebih dahulu');
          const benar = sl.filter((x, i) => Number(x.k) === jb[i]).length, nilai = Math.round(benar / sl.length * 100);
          add_('Kumpul', { tugasId: p.tugasId, siswaId: me.id, jawaban: JSON.stringify({ a: jb, b: benar, n: sl.length }), link: '', nilai, tanggal: now });
          return { ok: true, data: { nilai, benar, total: sl.length } };
        }
        const sq = parseSoal_(tg.soal), ln = String(p.link || '').trim();
        let jw;
        if (sq.length) { // tugas esai/isian dengan daftar soal: satu jawaban per soal
          const ja = Array.isArray(p.jawabSoal) ? p.jawabSoal.map(v => String(v || '').trim().slice(0, 2000)) : [];
          if (ja.length !== sq.length || ja.some(v => !v)) throw new Error('Jawab semua soal terlebih dahulu');
          jw = JSON.stringify({ j: ja });
        } else {
          jw = String(p.jawaban || '').trim();
          if (!jw && !ln) throw new Error('Isi jawaban atau tautan terlebih dahulu');
        }
        if (ln && !/^https?:\/\//i.test(ln)) throw new Error('Tautan harus diawali http:// atau https://');
        if (ex) upd_('Kumpul', ex.id, { jawaban: jw, link: ln, tanggal: now });
        else add_('Kumpul', { tugasId: p.tugasId, siswaId: me.id, jawaban: jw, link: ln, nilai: '', tanggal: now });
        return { ok: true, data: true };
      }
      case 'subs': {
        need(staff);
        const tg = rows_('Tugas').find(t => t.id == p.tugasId);
        if (!tg) throw new Error('Tugas tidak ditemukan');
        const dd = dl_(tg.deadline), kum = rows_('Kumpul').filter(k => k.tugasId == tg.id), sqs = parseSoal_(tg.soal);
        const data = rows_('Enroll').filter(e => e.kelasId == tg.kelasId).map(e => {
          const k = kum.find(x => x.siswaId == e.siswaId);
          return { id: k ? k.id : '', nama: nm(e.siswaId), sudah: !!k, jawaban: k ? (jenis_(tg) === 'pilgan' ? 'Pilihan ganda: benar ' + hasil_(k.jawaban).benar + ' dari ' + hasil_(k.jawaban).total + ' soal' : (sqs.length ? ringkasEsai_(sqs, k.jawaban) : k.jawaban)) : '', link: k ? k.link || '' : '', nilai: k ? k.nilai : '',
            komentar: k ? k.komentar || '' : '', jwb: k && sqs.length ? jwbList_(k.jawaban) : [], sk: k ? rincianList_(k.rincian) : [], terlambat: !!(k && dd && new Date(str_(k.tanggal)) > dd) };
        }).sort((a, b) => (a.sudah - b.sudah) || a.nama.localeCompare(b.nama)); // yang belum mengumpulkan tampil paling atas
        return { ok: true, data: { soal: sqs.map(x => ({ q: x.q, img: x.img || '' })), rows: data } };
      }
      case 'grade': {
        need(staff);
        const patch = { komentar: String(p.komentar || '').slice(0, 500), nilai: '' };
        if (Array.isArray(p.rincian)) { // esai: nilai per nomor, nilai akhir dijumlahkan otomatis
          const kk = rows_('Kumpul').find(k => k.id == p.id);
          if (!kk) throw new Error('Data tidak ditemukan');
          const tt = rows_('Tugas').find(t => t.id == kk.tugasId), n = tt ? parseSoal_(tt.soal).length : 0;
          if (!n || p.rincian.length !== n) throw new Error('Rincian nilai tidak sesuai jumlah soal');
          const bobot = 100 / n; let sum = 0, any = false;
          const rin = p.rincian.map((v, i) => {
            if (v === '' || v == null) return '';
            let x = Number(v);
            if (isNaN(x) || x < 0 || x > bobot + 0.01) throw new Error('Soal ' + (i + 1) + ': nilai harus 0-' + Math.round(bobot * 100) / 100);
            if (Math.abs(x - bobot) < 0.01) x = bobot; // bobot desimal (mis. 33,33) dihitung penuh agar total tepat 100
            any = true; sum += x;
            return Math.round(x * 100) / 100;
          });
          patch.rincian = JSON.stringify(rin);
          patch.nilai = any ? Math.min(100, Math.round(sum * 100) / 100) : '';
        } else if (p.nilai !== '' && p.nilai != null) {
          const n = Number(p.nilai);
          if (isNaN(n) || n < 0 || n > 100) throw new Error('Nilai harus 0-100');
          patch.nilai = n;
        }
        upd_('Kumpul', p.id, patch);
        return { ok: true, data: true };
      }

      case 'del': {
        const okSheet = { Kelas: staff, Materi: staff, Tugas: staff, Users: admin };
        need(okSheet[p.sheet]);
        if (p.sheet === 'Users' && p.id == me.id) throw new Error('Tidak bisa menghapus akun sendiri');
        del_(p.sheet, p.id);
        return { ok: true, data: true };
      }

      case 'users': need(admin);
        return { ok: true, data: users.map(u => ({ id: u.id, nama: u.nama, username: u.username, role: u.role })) };
      case 'saveUser': need(admin);
        if (!p.nama || !p.username) throw new Error('Nama dan username wajib diisi');
        if (p.id) {
          const patch = { nama: p.nama, username: p.username, role: p.role };
          if (p.password) patch.password = hash_(p.password);
          upd_('Users', p.id, patch);
        } else {
          if (String(p.password).length < 6) throw new Error('Password min. 6 karakter');
          if (users.some(u => u.username === p.username)) throw new Error('Username sudah dipakai');
          add_('Users', { nama: p.nama, username: p.username, password: hash_(p.password), role: p.role });
        }
        return { ok: true, data: true };

      case 'dash': {
        const ks = rows_('Kelas'), en = rows_('Enroll'), tg = rows_('Tugas'), kum = rows_('Kumpul');
        let my = ks;
        if (me.role === 'guru') my = ks.filter(k => k.guruId == me.id);
        if (me.role === 'siswa') my = ks.filter(k => en.some(e => e.kelasId == k.id && e.siswaId == me.id));
        const ids = my.map(k => k.id), myT = tg.filter(t => ids.includes(t.kelasId) && (me.role !== 'siswa' || dibagi_(t)));
        let a1, a2;
        const dn = me.role === 'siswa' ? kum.filter(k => k.siswaId == me.id).map(k => k.tugasId) : [];
        if (me.role === 'siswa') {
          const done = kum.filter(k => k.siswaId == me.id).map(k => k.tugasId);
          a1 = ['Tugas belum dikumpulkan', myT.filter(t => !done.includes(t.id)).length];
          const nl = kum.filter(k => k.siswaId == me.id && k.nilai !== '').map(k => Number(k.nilai));
          a2 = ['Rata-rata nilai', nl.length ? Math.round(nl.reduce((x, y) => x + y, 0) / nl.length) : '-'];
        } else {
          const tids = myT.map(t => t.id), sub = kum.filter(k => tids.includes(k.tugasId));
          a1 = ['Perlu dinilai', sub.filter(k => k.nilai === '').length];
          a2 = [admin ? 'Total pengguna' : 'Siswa terdaftar', admin ? users.length : new Set(en.filter(e => ids.includes(e.kelasId)).map(e => e.siswaId)).size];
        }
        // ---- data grafik dashboard ----
        const tz = Session.getScriptTimeZone(), fd = (d, f) => Utilities.formatDate(d, tz, f), hari = [];
        for (let i = 13; i >= 0; i--) { const d = new Date(Date.now() - i * 864e5); hari.push({ k: fd(d, 'yyyy-MM-dd'), l: fd(d, 'dd/MM') }); }
        const perHari = list => { const m = {}; list.forEach(r => { const t = new Date(str_(r.tanggal)); if (!isNaN(t.getTime())) { const k = fd(t, 'yyyy-MM-dd'); m[k] = (m[k] || 0) + 1; } }); return hari.map(h => m[h.k] || 0); };
        const rata1 = a => Math.round(a.reduce((x, y) => x + y, 0) / a.length * 10) / 10, mtAll = rows_('Materi'), tidAll = myT.map(t => t.id);
        const charts = { lab: hari.map(h => h.l) };
        if (me.role === 'siswa') {
          const mine = kum.filter(k => k.siswaId == me.id && tidAll.includes(k.tugasId)), sel = rows_('Selesai').filter(x => x.siswaId == me.id);
          const dinilai = mine.filter(k => k.nilai !== '' && k.nilai != null);
          charts.nilai = dinilai.sort((a, b) => str_(a.tanggal) < str_(b.tanggal) ? -1 : 1).slice(-8).map(k => ({ j: (myT.find(t => t.id == k.tugasId) || {}).judul || '-', v: Number(k.nilai) }));
          charts.status = [dinilai.length, mine.length - dinilai.length, Math.max(0, myT.length - mine.length)];
          charts.progres = my.map(k => { const ms = mtAll.filter(m => m.kelasId == k.id); return { k: k.nama, s: ms.filter(m => sel.some(x => x.materiId == m.id)).length, t: ms.length }; }).filter(x => x.t > 0);
          charts.akt = [{ n: 'Pengumpulan tugas', d: perHari(mine) }, { n: 'Materi selesai', d: perHari(sel.filter(x => mtAll.some(m => m.id == x.materiId && ids.includes(m.kelasId)))) }];
        } else {
          const sb = kum.filter(k => tidAll.includes(k.tugasId)), gr = sb.filter(k => k.nilai !== '' && k.nilai != null);
          charts.kelas = my.map(k => { const kt = myT.filter(t => t.kelasId == k.id).map(t => t.id), v = gr.filter(x => kt.includes(x.tugasId)).map(x => Number(x.nilai)); return { k: k.nama, v: v.length ? rata1(v) : 0, n: v.length }; }).filter(x => x.n > 0);
          charts.sebaran = [0, 0, 0, 0, 0];
          gr.forEach(x => { const v = Number(x.nilai); charts.sebaran[v < 60 ? 0 : v < 70 ? 1 : v < 80 ? 2 : v < 90 ? 3 : 4]++; });
          const target = myT.reduce((a, t) => a + en.filter(e => e.kelasId == t.kelasId).length, 0);
          charts.status = [gr.length, sb.length - gr.length, Math.max(0, target - sb.length)];
          charts.target = my.map(k => ({ k: k.nama, t: myT.filter(t => t.kelasId == k.id).length, s: en.filter(e => e.kelasId == k.id).length })).filter(x => x.t > 0 && x.s > 0); // rincian target per kelas
          charts.akt = [{ n: 'Pengumpulan tugas', d: perHari(sb) }];
          if (admin) charts.peran = ['admin', 'guru', 'siswa'].map(r => users.filter(u => u.role === r).length);
        }
        return { ok: true, data: { charts, kelas: my.length, materi: rows_('Materi').filter(m => ids.includes(m.kelasId)).length, tugas: myT.length, a1, a2,
          upcoming: myT.filter(t => !dn.includes(t.id))
            .map(t => { const d = dl_(t.deadline); return { judul: t.judul, deadline: dlStr_(t.deadline), ms: d ? d.getTime() : 0, kelas: (my.find(k => k.id == t.kelasId) || {}).nama }; })
            .filter(t => t.ms && (me.role === 'siswa' || t.ms >= Date.now())).sort((x, y) => x.ms - y.ms).slice(0, 6) } };
      }
    }
    throw new Error('Aksi tidak dikenal: ' + a);
  } catch (e) { return { ok: false, msg: e.message }; }
}
