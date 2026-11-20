
/* =========================================================
   SISTEM MANAJEMEN BARANG MILIK NEGARA (BMN)
   BACKEND GOOGLE APPS SCRIPT
   ========================================================= */


/* =========================================================
   1. KONFIGURASI DATABASE
   ========================================================= */


/* =========================================================
   1. KONFIGURASI DATABASE V2
   ========================================================= */

// ID DATABASE KHUSUS SISTEM BMN V2
const DATABASE_ID = "1yTweXmy9VFM9Fys7eTcZZnKNRP_4Z4rJt8nPVNBopqM";

// Mengakses database V2 secara eksplisit
function getDatabase() {
  return SpreadsheetApp.openById(DATABASE_ID);
}



function setupDatabase() {
  const ss = getDatabase();

  const struktur = {
    Master_BMN: [
      "ID_Aset", "Jenis_BMN", "Kode_BMN", "Nama_Aset",
      "Kategori", "Merk", "Tipe_Model", "Nomor_Seri",
      "Tahun_Perolehan", "Kondisi", "Lokasi",
      "Status_Aset", "Status_Penggunaan", "Keterangan",
      "Foto_Aset", "Tanggal_Input", "Terakhir_Diperbarui",
      "Dihapus"
    ],

    Data_Peminjaman: [
      "ID_Peminjaman", "Tanggal_Pengajuan", "Kategori_Pengguna",
      "Nama_Peminjam", "NIP", "Jabatan", "Instansi", "No_HP",
      "ID_Aset", "Kode_BMN", "Nama_Aset", "Keperluan",
      "Tanggal_Pinjam", "Rencana_Kembali", "Status_Pengajuan",
      "Admin_Pemeriksa", "Waktu_Persetujuan", "Waktu_Penyerahan",
      "Kondisi_Saat_Diserahkan", "Tanggal_Dikembalikan",
      "Kondisi_Saat_Dikembalikan", "Petugas_Penerima", "Catatan"
    ],

    Data_Pengguna: [
      "ID_Pengguna", "Nama", "NIP", "Jabatan",
      "Instansi", "No_HP", "Kategori", "Tanggal_Terdaftar"
    ],

    Data_Admin: [
      "ID_Admin", "Nama_Admin", "Username", "Password_Hash",
      "Salt", "Role", "Status", "Terakhir_Login"
    ],

    Log_Aktivitas: [
      "ID_Log", "Waktu", "Username", "Aktivitas",
      "Modul", "ID_Referensi", "Keterangan"
    ],

    Pengaturan: ["Parameter", "Nilai"],

    Referensi: ["Jenis_Referensi", "Kode", "Nama"],

    Data_Penggunaan_Barang: [
      "ID_Penggunaan", "Waktu_Input", "Jenis_Form", "Jenis_ATK",
      "Nama_Barang", "Nama_Pemakai", "Tanggal", "Jumlah",
      "Pernyataan", "Status"
    ],

    Referensi_Barang: [
      "Jenis_Form", "Jenis_ATK", "Nama_Barang", "Satuan", "Aktif"
    ],

    Data_Usulan_ATK_Operasional: [
      "ID_Usulan", "Waktu_Input", "Nama_Pengusul", "Nama_Barang",
      "Jumlah", "Spesifikasi", "Perkiraan_Harga_Satuan",
      "Alasan_Kebutuhan", "Pernyataan", "Status"
    ],

    Master_Pemeliharaan: [
      "ID_Aset_Pemeliharaan", "Nama_Barang", "Kategori", "Merk",
      "Tipe_Model", "Nomor_Seri", "Lokasi", "Kondisi",
      "Status_Aktif", "Keterangan"
    ],

    Data_Pemindahan_Barang: ["ID_Pemindahan","Waktu_Input","Nama_Penanggung_Jawab_Ruang","NIP","ID_Aset","Nama_Barang_Merk","Jumlah","Kondisi","Ruang_Asal","Ruang_Tujuan","Alasan_Pemindahan","Pernyataan","Status"],
    Data_Pelaporan_Barang_Rusak: ["ID_Pelaporan_Rusak","Waktu_Input","Tanggal","Nama_Pelapor","ID_Aset","Nama_Barang_Merk","Tingkat_Kerusakan","Detail_Kerusakan","Foto_URL","Lokasi_Barang","Pernyataan","Status"],
    Data_Pelaporan_Alat_Gelas_Pecah: ["ID_Pelaporan_Pecah","Waktu_Input","Nama_Pelapor","Tanggal","Nama_Barang","Jumlah_Pecah","Pernyataan","Status"],
    Data_Pemeliharaan_Aset: [
      "ID_Pemeliharaan", "Waktu_Input", "ID_Aset_Pemeliharaan", "Nama_Barang",
      "Tanggal_Pemeliharaan", "Jenis_Pemeliharaan", "Temuan",
      "Hasil_Pemeliharaan", "Petugas_Pemeliharaan",
      "Dokumentasi_URL", "Pernyataan", "Status"
    ]
  };

  Object.entries(struktur).forEach(([nama, header]) => {
    let sheet = ss.getSheetByName(nama);

    if (!sheet) {
      sheet = ss.insertSheet(nama);
    }

    // Hanya membuat header jika sheet benar-benar kosong.
    // Data lama tidak akan ditimpa.
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, header.length)
        .setValues([header]);

      sheet.getRange(1, 1, 1, header.length)
        .setBackground("#173B5E")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");

      sheet.setFrozenRows(1);
    }
  });

  // Migrasi aman: perbarui nama kolom ID pada data pemeliharaan lama
  // tanpa menghapus atau mengubah isi data yang sudah tersimpan.
  const dataPemeliharaanLama = ss.getSheetByName("Data_Pemeliharaan_Aset");
  if (dataPemeliharaanLama && dataPemeliharaanLama.getLastColumn() >= 3) {
    const headerId = String(dataPemeliharaanLama.getRange(1, 3).getDisplayValue() || "").trim();
    if (headerId === "ID_Aset") {
      dataPemeliharaanLama.getRange(1, 3).setValue("ID_Aset_Pemeliharaan");
    }
  }

  const ref = ss.getSheetByName("Referensi");

  if (ref.getLastRow() === 1) {
    const dataReferensi = [
      ["Kategori", "ELEKTRONIK", "Elektronik"],
      ["Kategori", "KENDARAAN", "Kendaraan"],
      ["Kategori", "PERALATAN", "Peralatan Kantor"],
      ["Kondisi", "BAIK", "Baik"],
      ["Kondisi", "RUSAK_RINGAN", "Rusak Ringan"],
      ["Kondisi", "RUSAK_BERAT", "Rusak Berat"],
      ["Status_Aset", "AKTIF", "Aktif"],
      ["Status_Aset", "NONAKTIF", "Nonaktif"],
      ["Status_Penggunaan", "TERSEDIA", "Tersedia"],
      ["Status_Penggunaan", "DIPINJAM", "Dipinjam"],
      ["Status_Penggunaan", "PEMELIHARAAN", "Pemeliharaan"]
    ];

    ref.getRange(2, 1, dataReferensi.length, 3)
      .setValues(dataReferensi);
  }

  console.log("Database Sistem BMN V2 berhasil diperiksa: " + ss.getName());
return "Database Sistem BMN V2 berhasil diperiksa.";
}


/* =========================================================
   2. AUTENTIKASI ADMIN
   ========================================================= */

// Hash password: 5.000 iterasi SHA-256.
// Pertahankan pola ini agar sesuai dengan akun admin
// yang sudah dibuat menggunakan versi sebelumnya.

function buatHashPassword(password, salt) {
  let hash = String(password) + ":" + String(salt);

  for (let i = 0; i < 5000; i++) {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      hash,
      Utilities.Charset.UTF_8
    );

    hash = Utilities.base64Encode(digest);
  }

  return hash;
}

function setupAdminPertama() {
  const props = PropertiesService.getScriptProperties();

  // Mengambil data admin dari Script Properties
  const username = props.getProperty("INITIAL_ADMIN_USERNAME");
  const password = props.getProperty("INITIAL_ADMIN_PASSWORD");
  const nama = props.getProperty("INITIAL_ADMIN_NAME");

  // Memastikan seluruh data tersedia
  if (!username || !password || !nama) {
    throw new Error(
      "Data admin belum lengkap. Periksa Script Properties."
    );
  }

  if (password.length < 12) {
    throw new Error("Password admin minimal 12 karakter.");
  }

  const ss = getDatabase();
  const sheet = ss.getSheetByName("Data_Admin");

  if (!sheet) {
    throw new Error("Sheet Data_Admin tidak ditemukan. Jalankan setupDatabase terlebih dahulu.");
  }

  // Membaca nama kolom pada baris pertama
  const jumlahKolom = sheet.getLastColumn();
  const header = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getValues()[0]
    .map(h => String(h).trim());

  const kolom = {};
  header.forEach((namaKolom, index) => {
    kolom[namaKolom] = index;
  });

  // Memastikan kolom yang dibutuhkan tersedia
  const wajib = [
    "ID_Admin",
    "Nama_Admin",
    "Username",
    "Password_Hash",
    "Salt",
    "Role",
    "Status",
    "Terakhir_Login"
  ];

  wajib.forEach(namaKolom => {
    if (kolom[namaKolom] === undefined) {
      throw new Error("Kolom " + namaKolom + " tidak ditemukan di Data_Admin.");
    }
  });

  // Memeriksa apakah username sudah terdaftar
  if (sheet.getLastRow() >= 2) {
    const data = sheet
      .getRange(2, kolom["Username"] + 1, sheet.getLastRow() - 1, 1)
      .getValues();

    const usernameAda = data.some(row =>
      String(row[0]).trim().toLowerCase() === username.trim().toLowerCase()
    );

    if (usernameAda) {
      throw new Error("Username admin tersebut sudah terdaftar.");
    }
  }

  // Membuat salt dan hash password
  const salt = Utilities.getUuid();
  const passwordHash = buatHashPassword(password, salt);

  // Menyiapkan data sesuai urutan kolom pada sheet
  const barisBaru = new Array(header.length).fill("");

  barisBaru[kolom["ID_Admin"]] = Utilities.getUuid();
  barisBaru[kolom["Nama_Admin"]] = nama;
  barisBaru[kolom["Username"]] = username.trim().toLowerCase();
  barisBaru[kolom["Password_Hash"]] = passwordHash;
  barisBaru[kolom["Salt"]] = salt;
  barisBaru[kolom["Role"]] = "ADMIN";
  barisBaru[kolom["Status"]] = "AKTIF";
  barisBaru[kolom["Terakhir_Login"]] = "";

  // Menyimpan akun admin ke Google Sheets
  sheet.appendRow(barisBaru);

  // Menghapus Script Properties setelah akun berhasil disimpan
  props.deleteProperty("INITIAL_ADMIN_USERNAME");
  props.deleteProperty("INITIAL_ADMIN_PASSWORD");
  props.deleteProperty("INITIAL_ADMIN_NAME");

  return "Admin pertama berhasil dibuat. Silakan login ke aplikasi.";
}


function loginAdmin(username, password) {
  username = String(username || "").trim().toLowerCase();
  password = String(password || "");

  if (!username || !password) {
    throw new Error("Username dan password wajib diisi.");
  }

  const cache = CacheService.getScriptCache();
  const attemptKey = "login_attempt_" + username;
  const attempts = Number(cache.get(attemptKey) || 0);

  if (attempts >= 5) {
    throw new Error(
      "Terlalu banyak percobaan login. Silakan coba kembali setelah 15 menit."
    );
  }

  const ss = getDatabase();
  const sheet = ss.getSheetByName("Data_Admin");

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error("Data admin belum tersedia.");
  }

  const data = sheet.getDataRange().getValues();
  const header = data[0].map(String);

  const col = {};
  header.forEach((nama, index) => col[nama] = index);

  const indexUsername = col.Username;
  const indexPassword = col.Password_Hash;
  const indexSalt = col.Salt;
  const indexRole = col.Role;
  const indexStatus = col.Status;

  if (
    indexUsername === undefined ||
    indexPassword === undefined ||
    indexSalt === undefined ||
    indexRole === undefined ||
    indexStatus === undefined
  ) {
    throw new Error("Struktur Data_Admin tidak sesuai.");
  }

  let admin = null;
  let nomorBaris = -1;

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][indexUsername]).trim().toLowerCase() === username
    ) {
      admin = data[i];
      nomorBaris = i + 1;
      break;
    }
  }

  if (!admin) {
    cache.put(attemptKey, String(attempts + 1), 900);
    throw new Error("Username atau password salah.");
  }

  if (
    String(admin[indexStatus]).toUpperCase() !== "AKTIF" ||
    String(admin[indexRole]).toUpperCase() !== "ADMIN"
  ) {
    throw new Error("Akun admin tidak aktif atau tidak memiliki hak akses.");
  }

  const hashInput = buatHashPassword(
    password,
    admin[indexSalt]
  );

  if (hashInput !== String(admin[indexPassword])) {
    cache.put(attemptKey, String(attempts + 1), 900);
    throw new Error("Username atau password salah.");
  }

  cache.remove(attemptKey);

  const token = Utilities.getUuid() + Utilities.getUuid();

  const sesi = {
    id: admin[col.ID_Admin],
    nama: admin[col.Nama_Admin],
    username: admin[indexUsername],
    role: admin[indexRole]
  };

  cache.put(
    "session_" + token,
    JSON.stringify(sesi),
    21600
  );

  if (col.Terakhir_Login !== undefined) {
    sheet.getRange(nomorBaris, col.Terakhir_Login + 1)
      .setValue(new Date());
  }

  return {
    berhasil: true,
    token: token,
    nama: sesi.nama,
    username: sesi.username
  };
}


function verifikasiAdmin(token) {
  if (!token) return null;

  const cache = CacheService.getScriptCache();
  const sesi = cache.get("session_" + token);

  if (!sesi) return null;

  const adminSesi = JSON.parse(sesi);

  const sheet = getDatabase().getSheetByName("Data_Admin");

  if (!sheet || sheet.getLastRow() < 2) return null;

  const data = sheet.getDataRange().getValues();
  const header = data[0].map(String);

  const idxID = header.indexOf("ID_Admin");
  const idxRole = header.indexOf("Role");
  const idxStatus = header.indexOf("Status");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxID]) === String(adminSesi.id)) {
      if (
        String(data[i][idxRole]).toUpperCase() !== "ADMIN" ||
        String(data[i][idxStatus]).toUpperCase() !== "AKTIF"
      ) {
        cache.remove("session_" + token);
        return null;
      }

      return {
        id: adminSesi.id,
        nama: data[i][header.indexOf("Nama_Admin")],
        username: data[i][header.indexOf("Username")],
        role: data[i][idxRole]
      };
    }
  }

  return null;
}


function logoutAdmin(token) {
  if (token) {
    CacheService.getScriptCache()
      .remove("session_" + token);
  }

  return { berhasil: true };
}


/* =========================================================
   3. HALAMAN APLIKASI
   ========================================================= */

function doGet(e) {
  if (e && e.parameter && e.parameter.bmnHealth === "1") {
    let health = { ok: false, deployment: "BMN-V2", database: false, masterBMN: false };
    try {
      const ss = getDatabase();
      health.database = !!ss;
      const sh = ss.getSheetByName("Master_BMN");
      health.masterBMN = !!sh;
      if (sh) {
        health.lastRow = sh.getLastRow();
        health.headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0] : [];
        const required = ["ID_Aset", "Kode_BMN", "Nama_Aset", "Jenis_BMN", "Status_Aset", "Status_Penggunaan", "Dihapus"];
        health.missingHeaders = required.filter(h => !health.headers.includes(h));
      }
      health.ok = health.database && health.masterBMN && (!health.missingHeaders || health.missingHeaders.length === 0);
    } catch (err) {
      health.error = err && err.message ? err.message : String(err);
    }
    return ContentService.createTextOutput(JSON.stringify(health)).setMimeType(ContentService.MimeType.JSON);
  }

  const halaman = e && e.parameter ? e.parameter.page : "";

  if (halaman === "peminjaman") {
    return HtmlService.createHtmlOutputFromFile("PublicPeminjaman")
      .setTitle("Formulir Peminjaman BMN")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Sistem Manajemen BMN")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


/* =========================================================
   API UNTUK HALAMAN FRONTEND GITHUB PAGES
   ========================================================= */

// Mengirim hasil API kembali ke halaman GitHub melalui iframe tersembunyi.
function bmnApiResponse_(payload) {
  const json = JSON.stringify(payload).replace(/</g, "\\u003c");
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>window.top.postMessage(' + json + ', "*");<\/script>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Endpoint POST. Hanya fungsi yang tercantum di daftar ini yang dapat dipanggil.
function doPost(e) {
  let request = {};
  try {
    const raw = e && e.parameter ? e.parameter.bmnApi : "";
    if (!raw) throw new Error("Permintaan API BMN tidak ditemukan.");
    request = JSON.parse(raw);

    const api = {
      loginAdmin: loginAdmin,
      verifikasiAdmin: verifikasiAdmin,
      logoutAdmin: logoutAdmin,
      getDashboardData: getDashboardData,
      getMasterBMN: getMasterBMN,
      simpanAset: simpanAset,
      ubahStatusAset: ubahStatusAset,
      hapusAset: hapusAset,
      getReferensiBMN: getReferensiBMN,
      getDaftarPeminjamanAdmin: getDaftarPeminjamanAdmin,
      prosesPersetujuanPeminjaman: prosesPersetujuanPeminjaman,
      getDaftarPengembalianAdmin: getDaftarPengembalianAdmin,
      prosesPengembalianPeminjaman: prosesPengembalianPeminjaman,
      getLaporanPeminjamanBulanan: getLaporanPeminjamanBulanan,
      ambilAsetTersediaPublik: ambilAsetTersediaPublik,
      ajukanPeminjamanPublik: ajukanPeminjamanPublik,
      getReferensiFormPublik: getReferensiFormPublik,
      kirimPenggunaanBarang: kirimPenggunaanBarang,
      kirimUsulanATKOperasional: kirimUsulanATKOperasional,
      kirimPemeliharaanAset: kirimPemeliharaanAset,
      getDataPenggunaanBarangAdmin: getDataPenggunaanBarangAdmin,
      getDataUsulanATKAdmin: getDataUsulanATKAdmin,
      getDataPemeliharaanAdmin: getDataPemeliharaanAdmin,
      kirimPemindahanBarang: kirimPemindahanBarang,
      kirimPelaporanBarangRusak: kirimPelaporanBarangRusak,
      kirimPelaporanAlatGelasPecah: kirimPelaporanAlatGelasPecah,
      getDataPemindahanBarangAdmin: getDataPemindahanBarangAdmin,
      getDataPelaporanBarangRusakAdmin: getDataPelaporanBarangRusakAdmin,
      getDataPelaporanAlatGelasAdmin: getDataPelaporanAlatGelasAdmin,
      getRingkasanFormPengelolaanAdmin: getRingkasanFormPengelolaanAdmin
    };

    const method = String(request.method || "");
    if (!Object.prototype.hasOwnProperty.call(api, method)) {
      throw new Error("Fungsi API tidak diizinkan: " + method);
    }
    const args = Array.isArray(request.args) ? request.args : [];
    const result = api[method].apply(null, args);
    return bmnApiResponse_({
      __bmnApiResponse: true,
      id: String(request.id || ""),
      ok: true,
      result: result
    });
  } catch (error) {
    return bmnApiResponse_({
      __bmnApiResponse: true,
      id: String(request.id || ""),
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}


/* =========================================================
   4. DASHBOARD
   ========================================================= */

function getDashboardData(token) {
  wajibAdminBMN_(token);

  const ss = getDatabase();
  const sheet = ss.getSheetByName("Master_BMN");

  const hasil = {
    totalAset: 0,
    asetAktif: 0,
    tersedia: 0,
    dipinjam: 0,
    pemeliharaan: 0,
    nonaktif: 0,
    totalPeminjaman: 0
  };

  if (sheet && sheet.getLastRow() >= 2) {
    const values = sheet.getDataRange().getValues();
    const header = values[0].map(String);

    const idx = {};
    header.forEach((h, i) => idx[h] = i);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const deleted = String(row[idx.Dihapus]).toLowerCase();

      if (["true", "ya", "1"].includes(deleted)) continue;

      hasil.totalAset++;

      const status = String(row[idx.Status_Aset]).toUpperCase();
      const penggunaan = String(row[idx.Status_Penggunaan]).toUpperCase();

      if (status === "AKTIF") hasil.asetAktif++;
      if (status === "NONAKTIF") hasil.nonaktif++;

      if (penggunaan === "TERSEDIA") hasil.tersedia++;
      if (penggunaan === "DIPINJAM") hasil.dipinjam++;
      if (penggunaan === "PEMELIHARAAN") hasil.pemeliharaan++;
    }
  }

  const pinjam = ss.getSheetByName("Data_Peminjaman");

  if (pinjam && pinjam.getLastRow() >= 2) {
    hasil.totalPeminjaman = pinjam.getLastRow() - 1;
  }

  return hasil;
}


/* =========================================================
   5. FUNGSI BANTU MASTER BMN
   ========================================================= */

function wajibAdminBMN_(token) {
  const admin = verifikasiAdmin(token);

  if (!admin) {
    throw new Error(
      "Sesi admin tidak valid atau sudah berakhir. Silakan login kembali."
    );
  }

  return admin;
}


function ambilHeaderBMN_() {
  const sheet = getDatabase().getSheetByName("Master_BMN");

  if (!sheet) {
    throw new Error("Sheet Master_BMN tidak ditemukan.");
  }

  const jumlahKolom = sheet.getLastColumn();

  if (jumlahKolom === 0) {
    throw new Error("Header Master_BMN belum tersedia.");
  }

  const header = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getValues()[0]
    .map(h => String(h).trim());

  const posisi = {};

  header.forEach((nama, index) => {
    posisi[nama] = index;
  });

  return { sheet, header, posisi };
}


function catatAktivitasBMN_(admin, aktivitas, idAset, keterangan) {
  const sheet = getDatabase().getSheetByName("Log_Aktivitas");

  if (!sheet) return;

  sheet.appendRow([
    Utilities.getUuid(),
    new Date(),
    admin.username || "",
    aktivitas,
    "Master BMN",
    idAset || "",
    keterangan || ""
  ]);
}


/* =========================================================
   6. MEMBACA DATA MASTER BMN
   ========================================================= */

function getMasterBMN(token, filter) {
  wajibAdminBMN_(token);

  const data = ambilHeaderBMN_();
  const sheet = data.sheet;
  const posisi = data.posisi;
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];
  if (posisi.ID_Aset === undefined) {
    throw new Error("Kolom ID_Aset tidak ditemukan. Pastikan header Master_BMN sesuai struktur.");
  }

  // Kunci proses agar dua admin yang membuka daftar bersamaan tidak
  // mendapatkan ID berurutan yang sama.
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const jumlahBaris = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, jumlahBaris, data.header.length).getValues();

    // Cari nomor ASET terbesar hanya dari kolom ID_Aset, BUKAN Kode_BMN.
    let nomorTerbesar = 0;
    let lebarAngka = 4;
    values.forEach(row => {
      const id = String(row[posisi.ID_Aset] || "").trim();
      const cocok = id.match(/^ASET(\d+)$/i);
      if (cocok) {
        const nomor = parseInt(cocok[1], 10);
        if (Number.isFinite(nomor) && nomor > nomorTerbesar) nomorTerbesar = nomor;
        lebarAngka = Math.max(lebarAngka, cocok[1].length);
      }
    });

    // Beri ID berurutan kepada data manual yang kolom ID_Aset-nya masih kosong.
    // Jangan pernah mengisi ID_Aset dengan Kode_BMN.
    values.forEach((row, i) => {
      const adaIsi = ["Jenis_BMN", "Kode_BMN", "Nama_Aset"].some(nama => {
        const idx = posisi[nama];
        return idx !== undefined && String(row[idx] || "").trim() !== "";
      });
      const idLama = String(row[posisi.ID_Aset] || "").trim();
      if (adaIsi && !idLama) {
        nomorTerbesar++;
        const idBaru = "ASET" + String(nomorTerbesar).padStart(lebarAngka, "0");
        sheet.getRange(i + 2, posisi.ID_Aset + 1).setValue(idBaru);
        row[posisi.ID_Aset] = idBaru;
      }
    });
    SpreadsheetApp.flush();

    filter = filter || {};
    const kataKunci = String(filter.kataKunci || "").toLowerCase().trim();
    const jenisFilter = String(filter.jenisBMN || "").toLowerCase().trim();
    const statusFilter = String(filter.statusAset || "").toUpperCase().trim();
    const hasil = [];

    values.forEach(row => {
      const aset = {};
      data.header.forEach((nama, index) => {
        let value = row[index];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        aset[nama] = value;
      });

      const isiData = [aset.Jenis_BMN, aset.Kode_BMN, aset.Nama_Aset].some(v => String(v || "").trim());
      if (!isiData) return;

      const deleted = String(aset.Dihapus).toLowerCase();
      if (["true", "ya", "1"].includes(deleted)) return;
      if (jenisFilter && String(aset.Jenis_BMN || "").toLowerCase() !== jenisFilter) return;
      if (statusFilter && String(aset.Status_Aset || "").toUpperCase() !== statusFilter) return;

      if (kataKunci) {
        const teks = [aset.Jenis_BMN, aset.Kode_BMN, aset.Nama_Aset, aset.Kategori,
          aset.Merk, aset.Tipe_Model, aset.Nomor_Seri, aset.Lokasi].join(" ").toLowerCase();
        if (!teks.includes(kataKunci)) return;
      }
      hasil.push(aset);
    });

    return hasil;
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   7. MENAMBAHKAN DAN MENGEDIT ASET
   ========================================================= */

// Membuat ID berurutan berdasarkan nomor terbesar pada kolom ID_Aset.
// Contoh: ASET0534 menjadi ASET0535. ID lama tidak diubah.
function buatIdAsetBerikutnya_(sheet, posisi) {
  const lastRow = sheet.getLastRow();
  let nomorTerbesar = 0;
  let lebarAngka = 4;

  if (lastRow >= 2) {
    // PENTING: hanya membaca kolom ID_Aset. Kolom Kode_BMN tidak digunakan.
    const daftarId = sheet.getRange(2, posisi.ID_Aset + 1, lastRow - 1, 1).getDisplayValues();
    daftarId.forEach(function(baris) {
      const id = String(baris[0] || "").trim().toUpperCase();
      const cocok = id.match(/^ASET(\d+)$/);
      if (!cocok) return;
      const nomor = parseInt(cocok[1], 10);
      if (Number.isFinite(nomor) && nomor > nomorTerbesar) nomorTerbesar = nomor;
      lebarAngka = Math.max(lebarAngka, cocok[1].length);
    });
  }

  return "ASET" + String(nomorTerbesar + 1).padStart(lebarAngka, "0");
}

function simpanAset(token, dataAset) {
  const admin = wajibAdminBMN_(token);
  if (!dataAset || typeof dataAset !== "object") throw new Error("Data aset tidak valid.");

  const ss = getDatabase();
  const sheet = ss.getSheetByName("Master_BMN");
  if (!sheet) throw new Error("Sheet Master_BMN tidak ditemukan pada spreadsheet: " + ss.getName());

  const jumlahKolom = sheet.getLastColumn();
  if (jumlahKolom < 1) throw new Error("Header Master_BMN belum tersedia.");
  const header = sheet.getRange(1, 1, 1, jumlahKolom).getDisplayValues()[0].map(h => String(h).trim());
  const posisi = {};
  header.forEach((nama, i) => { if (nama) posisi[nama] = i; });

  ["ID_Aset", "Jenis_BMN", "Kode_BMN", "Nama_Aset"].forEach(nama => {
    if (posisi[nama] === undefined) throw new Error("Kolom '" + nama + "' tidak ditemukan di Master_BMN.");
  });
  ["Jenis_BMN", "Kode_BMN", "Nama_Aset"].forEach(nama => {
    if (!String(dataAset[nama] || "").trim()) throw new Error(nama.replace(/_/g, " ") + " wajib diisi.");
  });

  const jenisBMNDiizinkan = [
    "ALAT ANGKUTAN BERMOTOR",
    "MESIN PERALATAN NON TIK",
    "MESIN PERALATAN KHUSUS TIK",
    "ASET TETAP LAINNYA"
  ];
  if (!jenisBMNDiizinkan.includes(String(dataAset.Jenis_BMN).trim().toUpperCase())) {
    throw new Error("Jenis BMN harus dipilih dari daftar yang tersedia.");
  }

  const lokasiDiizinkan = [
    "Pos Satpam",
    "Teras Kantor",
    "Ruang Lobby",
    "Lorong",
    "Ruang Staff Depan",
    "Ruang Rapat",
    "Ruang Kepala Balai",
    "Ruang Staff Belakang",
    "Ruang AAS",
    "Ruang Instrumen",
    "Ruang Timbang",
    "Ruang Preparasi",
    "Ruang Gas",
    "Ruang Sampel",
    "Ruang Media",
    "Ruang Dekstruksi",
    "Ruang Inkubasi",
    "Ruang Cemaran Jamur",
    "Ruang Cemaran Bakteri",
    "Gudang Reagen",
    "Ruang Arsip"
  ];
  const lokasiAset = String(dataAset.Lokasi || "").trim();
  if (lokasiAset && !lokasiDiizinkan.includes(lokasiAset)) {
    throw new Error("Lokasi harus dipilih dari daftar lokasi yang tersedia.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const idAset = String(dataAset.ID_Aset || "").trim();
    const lastRow = sheet.getLastRow();
    let nomorBaris = 0;
    let barisLama = null;

    // ID_Aset yang diterima dari formulir harus sudah ada untuk operasi edit.
    if (idAset) {
      if (lastRow >= 2) {
        const ids = sheet.getRange(2, posisi.ID_Aset + 1, lastRow - 1, 1).getDisplayValues();
        const cocok = [];
        ids.forEach((r, i) => { if (String(r[0]).trim() === idAset) cocok.push(i + 2); });
        if (cocok.length > 1) throw new Error("ID_Aset duplikat ditemukan di beberapa baris. Periksa kolom ID_Aset.");
        if (cocok.length === 1) {
          nomorBaris = cocok[0];
          barisLama = sheet.getRange(nomorBaris, 1, 1, header.length).getValues()[0];
        }
      }
      if (!nomorBaris) throw new Error("ID aset untuk diedit tidak ditemukan. Data tidak ditambahkan sebagai aset baru. Muat ulang halaman.");
    }

    const sekarang = new Date();
    // Edit wajib memakai ID yang sudah ditemukan. Tambah memakai urutan ASET####.
    const idFinal = idAset || buatIdAsetBerikutnya_(sheet, posisi);
    const barisBaru = barisLama ? barisLama.slice() : new Array(header.length).fill("");

    header.forEach((nama, i) => {
      if (Object.prototype.hasOwnProperty.call(dataAset, nama)) barisBaru[i] = dataAset[nama];
    });
    barisBaru[posisi.ID_Aset] = idFinal;

    if (!barisLama) {
      nomorBaris = Math.max(sheet.getLastRow() + 1, 2);
      if (posisi.Tanggal_Input !== undefined) barisBaru[posisi.Tanggal_Input] = sekarang;
      if (posisi.Dihapus !== undefined) barisBaru[posisi.Dihapus] = false;
    }
    if (posisi.Status_Aset !== undefined && !String(barisBaru[posisi.Status_Aset] || "").trim()) barisBaru[posisi.Status_Aset] = "AKTIF";
    if (posisi.Status_Penggunaan !== undefined && !String(barisBaru[posisi.Status_Penggunaan] || "").trim()) barisBaru[posisi.Status_Penggunaan] = "TERSEDIA";
    if (posisi.Terakhir_Diperbarui !== undefined) barisBaru[posisi.Terakhir_Diperbarui] = sekarang;

    sheet.getRange(nomorBaris, 1, 1, header.length).setValues([barisBaru]);
    SpreadsheetApp.flush();

    const idTersimpan = String(sheet.getRange(nomorBaris, posisi.ID_Aset + 1).getDisplayValue()).trim();
    if (idTersimpan !== idFinal) throw new Error("Verifikasi gagal: ID aset tidak cocok setelah penyimpanan.");

    try {
      catatAktivitasBMN_(admin, barisLama ? "Mengubah data aset" : "Menambahkan aset", idFinal,
        (barisLama ? "Mengubah " : "Menambahkan ") + dataAset.Nama_Aset);
    } catch (logError) {
      console.error("Pencatatan log gagal: " + logError.message);
    }

    return {
      berhasil: true,
      pesan: barisLama ? "Aset berhasil diperbarui." : "Aset berhasil ditambahkan.",
      idAset: idFinal,
      nomorBaris: nomorBaris,
      operasi: barisLama ? "EDIT" : "TAMBAH",
      namaSpreadsheet: ss.getName(),
      namaSheet: sheet.getName()
    };
  } finally {
    lock.releaseLock();
  }
}
/* =========================================================
   8. MENGUBAH STATUS ASET
   ========================================================= */

function ubahStatusAset(token, idAset, statusBaru) {
  const admin = wajibAdminBMN_(token);

  statusBaru = String(statusBaru || "").toUpperCase();

  if (!["AKTIF", "NONAKTIF"].includes(statusBaru)) {
    throw new Error("Status aset tidak valid.");
  }

  const data = ambilHeaderBMN_();
  const sheet = data.sheet;
  const posisi = data.posisi;
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) throw new Error("Data aset belum tersedia.");

  const ids = sheet.getRange(
    2, posisi.ID_Aset + 1, lastRow - 1, 1
  ).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idAset)) {
      const row = i + 2;

      sheet.getRange(row, posisi.Status_Aset + 1)
        .setValue(statusBaru);

      sheet.getRange(row, posisi.Terakhir_Diperbarui + 1)
        .setValue(new Date());

      catatAktivitasBMN_(
        admin, "Mengubah status aset", idAset,
        "Status diubah menjadi " + statusBaru
      );

      return {
        berhasil: true,
        pesan: "Status aset berhasil diubah."
      };
    }
  }

  throw new Error("Aset tidak ditemukan.");
}


/* =========================================================
   9. MENGHAPUS ASET (SOFT DELETE)
   ========================================================= */

function hapusAset(token, idAset) {
  const admin = wajibAdminBMN_(token);
  const id = String(idAset || "").trim();
  if (!id) throw new Error("ID aset kosong. Muat ulang daftar aset lalu coba kembali.");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = ambilHeaderBMN_();
    const sheet = data.sheet;
    const posisi = data.posisi;
    if (posisi.ID_Aset === undefined) throw new Error("Kolom ID_Aset tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Data aset belum tersedia.");

    const ids = sheet.getRange(2, posisi.ID_Aset + 1, lastRow - 1, 1).getDisplayValues();
    const barisCocok = [];
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === id) barisCocok.push(i + 2);
    }
    if (barisCocok.length === 0) throw new Error("Aset dengan ID " + id + " tidak ditemukan di Master_BMN.");
    if (barisCocok.length > 1) throw new Error("ID aset duplikat ditemukan. Penghapusan dibatalkan agar tidak salah menghapus data.");

    const rowTarget = barisCocok[0];
    const namaAset = posisi.Nama_Aset === undefined
      ? ""
      : String(sheet.getRange(rowTarget, posisi.Nama_Aset + 1).getDisplayValue());

    // Penghapusan fisik: menghapus seluruh baris sheet, bukan mencentang kolom Dihapus.
    sheet.deleteRow(rowTarget);
    SpreadsheetApp.flush();

    // Verifikasi bahwa ID tersebut benar-benar sudah tidak ada pada sheet.
    const lastRowSesudah = sheet.getLastRow();
    if (lastRowSesudah >= 2) {
      const idsSesudah = sheet.getRange(2, posisi.ID_Aset + 1, lastRowSesudah - 1, 1).getDisplayValues();
      const masihAda = idsSesudah.some(function(baris) {
        return String(baris[0] || "").trim() === id;
      });
      if (masihAda) throw new Error("Baris belum berhasil terhapus dari spreadsheet. Silakan periksa izin dan sheet Master_BMN.");
    }

    try {
      catatAktivitasBMN_(admin, "Menghapus aset permanen", id,
        "Baris " + rowTarget + " dihapus dari Master_BMN: " + namaAset);
    } catch (logError) {
      console.error("Aset sudah dihapus, tetapi pencatatan log gagal: " + logError.message);
    }

    return {
      berhasil: true,
      pesan: "Aset berhasil dihapus.",
      idAset: id,
      barisTerhapus: rowTarget,
      operasi: "HAPUS_PERMANEN"
    };
  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
   10. MENGAMBIL REFERENSI FORMULIR
   ========================================================= */

function getReferensiBMN(token) {
  wajibAdminBMN_(token);

  const sheet = getDatabase().getSheetByName("Referensi");
  // Hapus hanya referensi kategori TANAH dan BANGUNAN dari daftar pilihan.
  // Data aset pada Master_BMN tidak disentuh.
  if (sheet && sheet.getLastRow() >= 2) {
    const referensi = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
    for (let i = referensi.length - 1; i >= 0; i--) {
      const jenis = String(referensi[i][0] || "").trim().toUpperCase();
      const kode = String(referensi[i][1] || "").trim().toUpperCase();
      if (jenis === "KATEGORI" && ["TANAH", "BANGUNAN"].includes(kode)) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  const hasil = {
    kategori: [],
    kondisi: [],
    statusAset: [],
    statusPenggunaan: []  };

  if (!sheet || sheet.getLastRow() < 2) return hasil;

  const values = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 3
  ).getValues();

  values.forEach(row => {
    const jenis = String(row[0]).trim();
    const kode = String(row[1]).trim();
    const nama = String(row[2]).trim();

    if (!kode || !nama) return;

    if (jenis === "Kategori") {
      if (!["TANAH", "BANGUNAN"].includes(kode.toUpperCase())) {
        hasil.kategori.push({ kode, nama });
      }
    } else if (jenis === "Kondisi") {
      hasil.kondisi.push({ kode, nama });
    } else if (jenis === "Status_Aset") {
      hasil.statusAset.push({ kode, nama });
    } else if (jenis === "Status_Penggunaan") {
      hasil.statusPenggunaan.push({ kode, nama });
    }
  });

  return hasil;
}



/* =========================================================
   MODUL FORMULIR PEMINJAMAN PUBLIK
   Diakses melalui URL aplikasi ?page=peminjaman
   ========================================================= */

function ambilAsetTersediaPublik() {
  const sheet = getDatabase().getSheetByName("Master_BMN");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(String);
  const idx = {};
  header.forEach((h, i) => idx[h] = i);
  ["ID_Aset", "Kode_BMN", "Nama_Aset", "Jenis_BMN", "Status_Aset", "Status_Penggunaan", "Dihapus"].forEach(h => {
    if (idx[h] === undefined) throw new Error("Kolom " + h + " tidak ditemukan di Master_BMN.");
  });
  return values.slice(1).filter(row => {
    const deleted = String(row[idx.Dihapus] || "").trim().toLowerCase();
    const statusAset = String(row[idx.Status_Aset] || "").trim().toUpperCase();
    const statusPenggunaan = String(row[idx.Status_Penggunaan] || "").trim().toUpperCase();
    // Data lama yang belum memiliki status penggunaan diperlakukan sebagai tersedia,
    // selama aset aktif dan tidak ditandai telah dihapus.
    return !["true", "ya", "1"].includes(deleted) &&
      statusAset === "AKTIF" &&
      (!statusPenggunaan || statusPenggunaan === "TERSEDIA");
  }).map(row => ({
    id: String(row[idx.ID_Aset] || ""), kode: String(row[idx.Kode_BMN] || ""),
    nama: String(row[idx.Nama_Aset] || ""), jenis: String(row[idx.Jenis_BMN] || "")
  })).filter(a => a.id && a.nama);
}

function ajukanPeminjamanPublik(data) {
  if (!data || typeof data !== "object") throw new Error("Data pengajuan tidak valid.");
  const kategori = String(data.kategori || "").trim();
  const nama = String(data.nama || "").trim();
  const nip = String(data.nip || "").trim();
  const nik = String(data.nik || "").trim();
  const jabatan = String(data.jabatan || "").trim();
  const instansi = String(data.instansi || "").trim();
  const noHp = String(data.noHp || "").trim();
  const idAset = String(data.idAset || "").trim();
  const keperluan = String(data.keperluan || "").trim();
  const tanggalPinjam = String(data.tanggalPinjam || "").trim();
  const rencanaKembali = String(data.rencanaKembali || "").trim();
  if (!["Pegawai Internal", "Pegawai Eksternal"].includes(kategori)) throw new Error("Kategori peminjam tidak valid.");
  if (!nama || !nip || !jabatan || !instansi || !noHp || !idAset || !keperluan || !tanggalPinjam || !rencanaKembali)
    throw new Error("Mohon lengkapi seluruh data wajib.");
  if (kategori === "Pegawai Eksternal" && !/^\d{16}$/.test(nik)) throw new Error("NIK pegawai eksternal harus terdiri dari 16 digit.");
  if (kategori === "Pegawai Internal" && !/^[0-9A-Za-z./-]+$/.test(nip)) throw new Error("Format NIP tidak valid.");
  const tPinjam = new Date(tanggalPinjam + "T00:00:00");
  const tKembali = new Date(rencanaKembali + "T00:00:00");
  if (isNaN(tPinjam.getTime()) || isNaN(tKembali.getTime()) || tKembali < tPinjam)
    throw new Error("Tanggal rencana pengembalian tidak boleh sebelum tanggal peminjaman.");
  const ss = getDatabase();
  const asetTersedia = ambilAsetTersediaPublik();
  const aset = asetTersedia.find(a => a.id === idAset);
  if (!aset) throw new Error("Aset tidak tersedia. Silakan muat ulang daftar aset.");
  const sheet = ss.getSheetByName("Data_Peminjaman");
  if (!sheet) throw new Error("Sheet Data_Peminjaman tidak ditemukan.");
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  if (!headers.includes("NIK")) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("NIK");
    headers.push("NIK");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!ambilAsetTersediaPublik().some(a => a.id === idAset)) throw new Error("Aset baru saja tidak tersedia. Silakan pilih aset lain.");
    const idPengajuan = "PJM" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + String(Math.floor(Math.random() * 900) + 100);
    const record = {
      ID_Peminjaman: idPengajuan, Tanggal_Pengajuan: new Date(), Kategori_Pengguna: kategori,
      Nama_Peminjam: nama, NIP: nip, Jabatan: jabatan, Instansi: instansi, No_HP: noHp,
      ID_Aset: aset.id, Kode_BMN: aset.kode, Nama_Aset: aset.nama, Keperluan: keperluan,
      Tanggal_Pinjam: tPinjam, Rencana_Kembali: tKembali,
      Status_Pengajuan: "MENUNGGU PERSETUJUAN", NIK: kategori === "Pegawai Eksternal" ? nik : ""
    };
    sheet.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(record, h) ? record[h] : ""));
    return { berhasil: true, pesan: "Pengajuan peminjaman berhasil dikirim.", id: idPengajuan };
  } finally { lock.releaseLock(); }
}


/* =========================================================
   MODUL ADMIN: PERSETUJUAN, PENGEMBALIAN, DAN LAPORAN
   ========================================================= */

function ambilDataPeminjamanAdmin_(token) {
  wajibAdminBMN_(token);
  const sh = getDatabase().getSheetByName('Data_Peminjaman');
  if (!sh || sh.getLastRow() < 2) return { sheet: sh, headers: [], rows: [] };
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  const rows = values.slice(1).map((r, n) => {
    const o = { _row: n + 2 };
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  }).filter(o => String(o.ID_Peminjaman || '').trim());
  return { sheet: sh, headers, rows, idx };
}

function nilaiTanggalBMN_(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const d = new Date(v);
  return isNaN(d.getTime()) ? String(v) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatPeminjamanAdmin_(o) {
  return {
    id: String(o.ID_Peminjaman || ''), tanggalPengajuan: nilaiTanggalBMN_(o.Tanggal_Pengajuan),
    kategori: String(o.Kategori_Pengguna || ''), nama: String(o.Nama_Peminjam || ''),
    nip: String(o.NIP || ''), nik: String(o.NIK || ''), jabatan: String(o.Jabatan || ''),
    instansi: String(o.Instansi || ''), noHp: String(o.No_HP || ''),
    idAset: String(o.ID_Aset || ''), kode: String(o.Kode_BMN || ''), namaAset: String(o.Nama_Aset || ''),
    keperluan: String(o.Keperluan || ''), tanggalPinjam: nilaiTanggalBMN_(o.Tanggal_Pinjam),
    rencanaKembali: nilaiTanggalBMN_(o.Rencana_Kembali), status: String(o.Status_Pengajuan || ''),
    admin: String(o.Admin_Pemeriksa || ''), waktuPersetujuan: nilaiTanggalBMN_(o.Waktu_Persetujuan),
    tanggalDikembalikan: nilaiTanggalBMN_(o.Tanggal_Dikembalikan),
    kondisiDiserahkan: String(o.Kondisi_Saat_Diserahkan || ''),
    kondisiDikembalikan: String(o.Kondisi_Saat_Dikembalikan || ''),
    petugasPenerima: String(o.Petugas_Penerima || ''), catatan: String(o.Catatan || '')
  };
}

function getDaftarPeminjamanAdmin(token) {
  const data = ambilDataPeminjamanAdmin_(token);
  return data.rows.map(formatPeminjamanAdmin_).sort((a,b) => b.tanggalPengajuan.localeCompare(a.tanggalPengajuan));
}

function prosesPersetujuanPeminjaman(token, idPeminjaman, keputusan, catatan, kondisiDiserahkan) {
  const admin = wajibAdminBMN_(token);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const data = ambilDataPeminjamanAdmin_(token);
    const rec = data.rows.find(r => String(r.ID_Peminjaman) === String(idPeminjaman));
    if (!rec) throw new Error('Pengajuan peminjaman tidak ditemukan.');
    if (String(rec.Status_Pengajuan).toUpperCase() !== 'MENUNGGU PERSETUJUAN') throw new Error('Pengajuan ini sudah diproses sebelumnya.');
    const keputusanNorm = String(keputusan || '').toUpperCase();
    if (!['DISETUJUI','DITOLAK'].includes(keputusanNorm)) throw new Error('Keputusan tidak valid.');
    const sh = data.sheet, idx = data.idx, row = rec._row;
    const set = (h,v) => { if (idx[h] !== undefined) sh.getRange(row, idx[h]+1).setValue(v); };
    let asetSh = null, asetPos = -1, asetIx = {};
    if (keputusanNorm === 'DISETUJUI') {
      asetSh = getDatabase().getSheetByName('Master_BMN');
      if (!asetSh || asetSh.getLastRow() < 2) throw new Error('Data aset tidak ditemukan.');
      const vals = asetSh.getDataRange().getValues(), h = vals[0].map(String);
      h.forEach((x,i)=>asetIx[x]=i);
      asetPos = vals.findIndex((r,i)=>i>0 && String(r[asetIx.ID_Aset])===String(rec.ID_Aset));
      if (asetPos < 0) throw new Error('Data aset tidak ditemukan.');
      if (String(vals[asetPos][asetIx.Status_Penggunaan]).toUpperCase() !== 'TERSEDIA') throw new Error('Aset sudah tidak tersedia. Pengajuan belum diubah.');
    }
    set('Status_Pengajuan', keputusanNorm);
    set('Admin_Pemeriksa', admin.username || admin.nama || 'ADMIN');
    set('Waktu_Persetujuan', new Date());
    set('Catatan', String(catatan || '').trim());
    if (keputusanNorm === 'DISETUJUI') {
      set('Kondisi_Saat_Diserahkan', String(kondisiDiserahkan || 'BAIK'));
      asetSh.getRange(asetPos+1, asetIx.Status_Penggunaan+1).setValue('DIPINJAM');
      if (asetIx.Terakhir_Diperbarui !== undefined) asetSh.getRange(asetPos+1,asetIx.Terakhir_Diperbarui+1).setValue(new Date());
    }
    SpreadsheetApp.flush();
    return {berhasil:true,pesan: keputusanNorm === 'DISETUJUI' ? 'Pengajuan disetujui dan aset ditandai dipinjam.' : 'Pengajuan ditolak.'};
  } finally { lock.releaseLock(); }
}

function getDaftarPengembalianAdmin(token) {
  return getDaftarPeminjamanAdmin(token).filter(r => r.status.toUpperCase() === 'DISETUJUI');
}

function prosesPengembalianPeminjaman(token, idPeminjaman, tanggalKembali, kondisiKembali, petugas, catatan) {
  const admin = wajibAdminBMN_(token);
  if (!tanggalKembali) throw new Error('Tanggal pengembalian wajib diisi.');
  const kondisi = String(kondisiKembali || '').trim().toUpperCase();
  if (!['BAIK','RUSAK_RINGAN','RUSAK_BERAT'].includes(kondisi)) throw new Error('Kondisi barang tidak valid.');

  const cocok = String(tanggalKembali).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!cocok) throw new Error('Format tanggal pengembalian harus YYYY-MM-DD.');
  const tanggal = new Date(Number(cocok[1]), Number(cocok[2])-1, Number(cocok[3]), 12, 0, 0);
  if (isNaN(tanggal.getTime()) || tanggal.getFullYear() !== Number(cocok[1]) || (tanggal.getMonth()+1) !== Number(cocok[2]) || tanggal.getDate() !== Number(cocok[3])) {
    throw new Error('Tanggal pengembalian tidak valid.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = ambilDataPeminjamanAdmin_(token);
    const rec = data.rows.find(r => String(r.ID_Peminjaman).trim() === String(idPeminjaman).trim());
    if (!rec) throw new Error('Data peminjaman tidak ditemukan. Muat ulang daftar pengembalian.');
    if (String(rec.Status_Pengajuan || '').trim().toUpperCase() !== 'DISETUJUI') {
      throw new Error('Peminjaman ini tidak berstatus DISETUJUI atau sudah pernah dikembalikan.');
    }

    const asetSh = getDatabase().getSheetByName('Master_BMN');
    if (!asetSh || asetSh.getLastRow() < 2) throw new Error('Sheet Master_BMN atau data aset tidak ditemukan.');
    const asetValues = asetSh.getDataRange().getValues();
    const asetHeaders = asetValues[0].map(h => String(h).trim());
    const asetIx = {};
    asetHeaders.forEach((h,i) => asetIx[h] = i);
    ['ID_Aset','Status_Penggunaan'].forEach(h => { if (asetIx[h] === undefined) throw new Error('Kolom ' + h + ' tidak ditemukan pada Master_BMN.'); });
    const asetPos = asetValues.findIndex((r,i) => i > 0 && String(r[asetIx.ID_Aset]).trim() === String(rec.ID_Aset).trim());
    if (asetPos < 0) throw new Error('Aset dengan ID ' + rec.ID_Aset + ' tidak ditemukan pada Master_BMN.');
    if (String(asetValues[asetPos][asetIx.Status_Penggunaan] || '').trim().toUpperCase() !== 'DIPINJAM') {
      throw new Error('Status aset bukan DIPINJAM. Periksa data aset sebelum mencatat pengembalian.');
    }

    const sh = data.sheet;
    const idx = data.idx || {};
    const wajib = ['Status_Pengajuan','Tanggal_Dikembalikan','Kondisi_Saat_Dikembalikan','Petugas_Penerima'];
    wajib.forEach(h => { if (idx[h] === undefined) throw new Error('Kolom ' + h + ' tidak ditemukan pada Data_Peminjaman.'); });

    // Tulis seluruh informasi pengembalian ke baris transaksi yang sama.
    const row = rec._row;
    sh.getRange(row, idx.Tanggal_Dikembalikan+1).setValue(tanggal);
    sh.getRange(row, idx.Kondisi_Saat_Dikembalikan+1).setValue(kondisi);
    sh.getRange(row, idx.Petugas_Penerima+1).setValue(String(petugas || admin.username || admin.nama || 'ADMIN').trim());
    sh.getRange(row, idx.Status_Pengajuan+1).setValue('DIKEMBALIKAN');
    if (idx.Catatan !== undefined && String(catatan || '').trim()) sh.getRange(row, idx.Catatan+1).setValue(String(catatan).trim());

    // Kembalikan status aset sesuai kondisi fisik saat diterima.
    const barisAset = asetPos + 1; // indeks array termasuk header sama dengan nomor baris sheet - 1
    asetSh.getRange(barisAset, asetIx.Status_Penggunaan+1).setValue(kondisi === 'BAIK' ? 'TERSEDIA' : 'PEMELIHARAAN');
    if (asetIx.Kondisi !== undefined) asetSh.getRange(barisAset, asetIx.Kondisi+1).setValue(kondisi);
    if (asetIx.Terakhir_Diperbarui !== undefined) asetSh.getRange(barisAset, asetIx.Terakhir_Diperbarui+1).setValue(new Date());

    SpreadsheetApp.flush();

    // Verifikasi hasil tulis agar UI tidak menyatakan berhasil jika data belum tersimpan.
    const hasilPinjam = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const tanggalTersimpan = hasilPinjam[idx.Tanggal_Dikembalikan];
    const statusTersimpan = String(hasilPinjam[idx.Status_Pengajuan] || '').toUpperCase();
    const kondisiTersimpan = String(hasilPinjam[idx.Kondisi_Saat_Dikembalikan] || '').toUpperCase();
    const statusAsetTersimpan = String(asetSh.getRange(barisAset, asetIx.Status_Penggunaan+1).getValue() || '').toUpperCase();
    if (!(tanggalTersimpan instanceof Date) || statusTersimpan !== 'DIKEMBALIKAN' || kondisiTersimpan !== kondisi || statusAsetTersimpan !== (kondisi === 'BAIK' ? 'TERSEDIA' : 'PEMELIHARAAN')) {
      throw new Error('Data pengembalian belum tersimpan sempurna. Periksa kembali sheet Data_Peminjaman dan Master_BMN.');
    }

    return {berhasil:true, pesan:'Pengembalian berhasil dicatat. Tanggal, kondisi, dan status aset telah diperbarui.'};
  } finally {
    lock.releaseLock();
  }
}

function getLaporanPeminjamanBulanan(token, bulan, tahun) {
  wajibAdminBMN_(token);
  const b = Number(String(bulan == null ? "" : bulan).trim());
  const y = Number(String(tahun == null ? "" : tahun).trim());
  if (!Number.isInteger(b) || b < 1 || b > 12 || !Number.isInteger(y) || y < 2000 || y > 2100) {
    throw new Error('Periode laporan tidak valid. Pilih bulan dan isi tahun antara 2000–2100.');
  }
  const semua = getDaftarPeminjamanAdmin(token);
  const rows = semua.filter(r => {
    if (!r.tanggalPinjam) return false;
    const p = r.tanggalPinjam.split('-'); return Number(p[0])===y && Number(p[1])===b;
  });
  const status = {menunggu:0, disetujui:0, ditolak:0, dikembalikan:0};
  rows.forEach(r=>{
    const s=r.status.toUpperCase();
    if(s==='MENUNGGU PERSETUJUAN') status.menunggu++;
    else if(s==='DISETUJUI') status.disetujui++;
    else if(s==='DITOLAK') status.ditolak++;
    else if(s==='DIKEMBALIKAN') status.dikembalikan++;
  });
  return {bulan:b,tahun:y,total:rows.length,status:status,rows:rows.sort((a,b)=>a.tanggalPinjam.localeCompare(b.tanggalPinjam))};
}
function tesKoneksiDatabaseV2() {
  const ss = getDatabase();

  const hasil = {
    namaDatabase: ss.getName(),
    idDatabase: ss.getId(),
    daftarSheet: ss.getSheets().map(sheet => sheet.getName()),
    status: "KONEKSI BERHASIL"
  };

  console.log(JSON.stringify(hasil, null, 2));

  return hasil;
}

/* =========================================================
   MODUL FORMULIR PENGELOLAAN BMN V2
   ========================================================= */

const JENIS_FORM_PENGGUNAAN_ = [
  "ATK", "REAGEN DAN MEDIA", "BAKU PEMBANDING",
  "BAKU MIKROBA", "SUKU CADANG", "ALAT GELAS"
];
const PERNYATAAN_PENGGUNAAN_ = "Saya bersedia menggunakan barang tersebut sesuai dengan kebutuhan dan ketentuan, mencatat setiap pemakaian dengan benar, serta bertanggung jawab atas penggunaan yang tidak sesuai prosedur.";
const PERNYATAAN_USULAN_ = "Saya menyatakan bahwa usulan pengadaan ini disampaikan dengan sebenar-benarnya dan dapat dipertanggungjawabkan.";
const PERNYATAAN_PEMELIHARAAN_ = "Saya bersedia menjaga dan memelihara aset yang menjadi tanggung jawab saya dengan melakukan pemeliharaan sesuai ketentuan serta melaporkan apabila terjadi kerusakan.";
const PETUGAS_PEMELIHARAAN_ = [
  "Anwar", "Syaiful Abdul Gani", "Viktorimus Sesarius Porkas", "Teknisi"
];

function sheetBMN_(nama, header) {
  const ss = getDatabase();
  let sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setBackground("#173B5E").setFontColor("#FFFFFF").setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function tanggalISOValidBMN_(v) {
  const s = String(v || "").trim();
  if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(s)) return false;
  const d = new Date(s + "T12:00:00");
  return !isNaN(d.getTime()) &&
    Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") === s;
}

function getReferensiFormPublik() {
  const ss = getDatabase();
  const ref = ss.getSheetByName("Referensi_Barang");
  const daftarBarang = [];
  const jenisATK = [];
  if (ref && ref.getLastRow() >= 2) {
    const values = ref.getRange(2, 1, ref.getLastRow() - 1, 5).getDisplayValues();
    values.forEach(r => {
      const jenis = String(r[0] || "").trim().toUpperCase();
      const atk = String(r[1] || "").trim();
      const barang = String(r[2] || "").trim();
      const aktif = String(r[4] || "YA").trim().toUpperCase();
      if (aktif === "TIDAK" || aktif === "FALSE" || aktif === "0") return;
      if (jenis === "ATK" && atk && !jenisATK.includes(atk)) jenisATK.push(atk);
      if (jenis && barang) daftarBarang.push({ jenisForm: jenis, jenisATK: atk, namaBarang: barang, satuan: r[3] || "" });
    });
  }

  const asetSh = ss.getSheetByName("Master_Pemeliharaan");
  const daftarPemeliharaan = [];
  if (asetSh && asetSh.getLastRow() >= 2) {
    const v = asetSh.getDataRange().getDisplayValues();
    const h = v[0].map(x => String(x).trim());
    const ix = {};
    h.forEach((x,i) => ix[x] = i);
    const idxId = ix.ID_Aset_Pemeliharaan;
    const idxNama = ix.Nama_Barang;
    const idxAktif = ix.Status_Aktif;
    v.slice(1).forEach(r => {
      const id = idxId === undefined ? "" : String(r[idxId] || "").trim();
      const nama = idxNama === undefined ? "" : String(r[idxNama] || "").trim();
      const status = idxAktif === undefined ? "AKTIF" : String(r[idxAktif] || "").trim().toUpperCase();
      if (id && nama && (!status || ["AKTIF","YA","TRUE","1"].includes(status))) {
        daftarPemeliharaan.push({ id: id, nama: nama });
      }
    });
  }
  const master = ss.getSheetByName("Master_BMN");
  const daftarAset = [], daftarLokasi = [];
  if (master && master.getLastRow() >= 2) {
    const values = master.getDataRange().getDisplayValues();
    const h = values[0].map(x => String(x).trim()), ix = {};
    h.forEach((x,i) => ix[x] = i);
    values.slice(1).forEach(r => {
      const id = String(r[ix.ID_Aset] || "").trim();
      const nama = String(r[ix.Nama_Aset] || "").trim();
      const lokasi = String(r[ix.Lokasi] || "").trim();
      const status = String(r[ix.Status_Aset] || "").trim().toUpperCase();
      const dihapus = String(r[ix.Dihapus] || "").trim().toUpperCase();
      if (!id || !nama || ["YA","TRUE","1","DIHAPUS"].includes(dihapus)) return;
      const merk = ix.Merk === undefined ? "" : String(r[ix.Merk] || "").trim();
      daftarAset.push({id:id,nama:nama,merk:merk,lokasi:lokasi,label:[nama,merk].filter(Boolean).join(" / ")});
      if (lokasi && !["NONAKTIF","TIDAK AKTIF"].includes(status) && !daftarLokasi.includes(lokasi)) daftarLokasi.push(lokasi);
    });
  }
  const daftarAlatGelas = [...new Set(daftarBarang.filter(x => x.jenisForm === "ALAT GELAS").map(x => x.namaBarang))].sort();
  return {
    jenisForm: JENIS_FORM_PENGGUNAAN_, jenisATK: jenisATK.sort(), daftarBarang: daftarBarang,
    daftarPemeliharaan: daftarPemeliharaan, daftarAset: daftarAset, daftarLokasi: daftarLokasi.sort(),
    daftarAlatGelas: daftarAlatGelas,
    jenisPemeliharaan: ["RUTIN", "PERBAIKAN RINGAN", "DARURAT"], petugasPemeliharaan: PETUGAS_PEMELIHARAAN_
  };
}

function kirimPenggunaanBarang(data) {
  if (!data || typeof data !== "object") throw new Error("Data formulir tidak valid.");
  const jenis = String(data.jenisForm || "").trim().toUpperCase();
  const jenisATK = String(data.jenisATK || "").trim();
  const barang = String(data.namaBarang || "").trim();
  const pemakai = String(data.namaPemakai || "").trim();
  const tanggal = String(data.tanggal || "").trim();
  const jumlah = Number(data.jumlah);
  if (!JENIS_FORM_PENGGUNAAN_.includes(jenis)) throw new Error("Jenis formulir tidak valid.");
  if (jenis === "ATK" && !jenisATK) throw new Error("Jenis ATK wajib dipilih.");
  if (!barang || !pemakai || !tanggalISOValidBMN_(tanggal) || !(jumlah > 0)) {
    throw new Error("Mohon lengkapi nama barang, nama pemakai, tanggal, dan jumlah dengan benar.");
  }
  const ref = getReferensiFormPublik().daftarBarang;
  const cocok = ref.some(x => x.jenisForm === jenis && x.namaBarang === barang &&
    (jenis !== "ATK" || x.jenisATK === jenisATK));
  if (!cocok) throw new Error("Nama barang tidak sesuai dengan referensi database. Muat ulang formulir.");
  const sh = sheetBMN_("Data_Penggunaan_Barang", [
    "ID_Penggunaan","Waktu_Input","Jenis_Form","Jenis_ATK","Nama_Barang",
    "Nama_Pemakai","Tanggal","Jumlah","Pernyataan","Status"
  ]);
  const id = "PGN" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + Math.floor(100 + Math.random()*900);
  sh.appendRow([id,new Date(),jenis,jenis === "ATK" ? jenisATK : "",barang,pemakai,new Date(tanggal+"T12:00:00"),jumlah,PERNYATAAN_PENGGUNAAN_,"BARU"]);
  return {berhasil:true,pesan:"Formulir penggunaan barang berhasil dikirim.",id:id};
}

function kirimUsulanATKOperasional(data) {
  if (!data || typeof data !== "object") throw new Error("Data formulir tidak valid.");
  const nama = String(data.namaPengusul || "").trim();
  const barang = String(data.namaBarang || "").trim();
  const jumlah = Number(data.jumlah);
  const spesifikasi = String(data.spesifikasi || "").trim();
  const harga = Number(data.perkiraanHargaSatuan);
  const alasan = String(data.alasanKebutuhan || "").trim();
  if (!nama || !barang || !(jumlah > 0) || !spesifikasi || !(harga >= 0) || !alasan) {
    throw new Error("Mohon lengkapi seluruh data usulan dengan benar.");
  }
  const sh = sheetBMN_("Data_Usulan_ATK_Operasional", [
    "ID_Usulan","Waktu_Input","Nama_Pengusul","Nama_Barang","Jumlah",
    "Spesifikasi","Perkiraan_Harga_Satuan","Alasan_Kebutuhan","Pernyataan","Status"
  ]);
  const id = "USL" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + Math.floor(100 + Math.random()*900);
  sh.appendRow([id,new Date(),nama,barang,jumlah,spesifikasi,harga,alasan,PERNYATAAN_USULAN_,"BARU"]);
  return {berhasil:true,pesan:"Usulan ATK dan operasional berhasil dikirim.",id:id};
}

function simpanDokumentasiPemeliharaan_(foto) {
  if (!foto || typeof foto !== "object") throw new Error("Dokumentasi foto wajib diunggah.");
  const mime = String(foto.mimeType || "").toLowerCase();
  if (!["image/jpeg","image/png","image/webp"].includes(mime)) throw new Error("Foto harus berformat JPG, PNG, atau WEBP.");
  const base64 = String(foto.base64 || "").replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/, "");
  if (!base64 || base64.length > 7 * 1024 * 1024) throw new Error("Ukuran foto terlalu besar. Gunakan foto maksimal sekitar 5 MB.");
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 5 * 1024 * 1024) throw new Error("Ukuran foto maksimal 5 MB.");
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty("FOLDER_DOKUMENTASI_PEMELIHARAAN");
  let folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder("Dokumentasi Pemeliharaan BMN V2");
    props.setProperty("FOLDER_DOKUMENTASI_PEMELIHARAAN", folder.getId());
  }
  const nama = "PEMELIHARAAN_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + "." + (mime === "image/png" ? "png" : mime === "image/webp" ? "webp" : "jpg");
  const file = folder.createFile(Utilities.newBlob(bytes, mime, nama));
  return file.getUrl();
}

function kirimPemeliharaanAset(data) {
  if (!data || typeof data !== "object") throw new Error("Data formulir tidak valid.");
  const idAset = String(data.idAsetPemeliharaan || "").trim();
  const tanggal = String(data.tanggalPemeliharaan || "").trim();
  const jenis = String(data.jenisPemeliharaan || "").trim().toUpperCase();
  const temuan = String(data.temuan || "").trim();
  const hasil = String(data.hasilPemeliharaan || "").trim();
  const petugas = String(data.petugasPemeliharaan || "").trim();
  if (!idAset || !tanggalISOValidBMN_(tanggal) ||
      !["RUTIN","PERBAIKAN RINGAN","DARURAT"].includes(jenis) ||
      !temuan || !hasil || !PETUGAS_PEMELIHARAAN_.includes(petugas)) {
    throw new Error("Mohon lengkapi seluruh data pemeliharaan dengan benar.");
  }
  const aset = getReferensiFormPublik().daftarPemeliharaan.find(x => x.id === idAset);
  if (!aset) throw new Error("Aset tidak ditemukan atau tidak aktif.");
  const urlFoto = simpanDokumentasiPemeliharaan_(data.dokumentasi);
  const sh = sheetBMN_("Data_Pemeliharaan_Aset", [
    "ID_Pemeliharaan","Waktu_Input","ID_Aset_Pemeliharaan","Nama_Barang",
    "Tanggal_Pemeliharaan","Jenis_Pemeliharaan","Temuan",
    "Hasil_Pemeliharaan","Petugas_Pemeliharaan","Dokumentasi_URL","Pernyataan","Status"
  ]);
  const id = "PMH" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + Math.floor(100 + Math.random()*900);
  sh.appendRow([id,new Date(),idAset,aset.nama,new Date(tanggal+"T12:00:00"),jenis,temuan,hasil,petugas,urlFoto,PERNYATAAN_PEMELIHARAAN_,"BARU"]);
  return {berhasil:true,pesan:"Laporan pemeliharaan berhasil dikirim.",id:id,dokumentasi:urlFoto};
}


function asetPilihanFormBMN_(id) {
  const ref = getReferensiFormPublik();
  const aset = ref.daftarAset.find(x => x.id === String(id || "").trim());
  if (!aset) throw new Error("Barang tidak ditemukan pada Master BMN.");
  return aset;
}

function kirimPemindahanBarang(data) {
  if (!data || typeof data !== "object") throw new Error("Data formulir tidak valid.");
  const nama=String(data.namaPenanggungJawab||"").trim(), nip=String(data.nip||"").trim();
  const id=String(data.idAset||"").trim(), jumlah=Number(data.jumlah);
  const kondisi=String(data.kondisi||"").trim().toUpperCase();
  const asal=String(data.ruangAsal||"").trim(), tujuan=String(data.ruangTujuan||"").trim();
  const alasan=String(data.alasanPemindahan||"").trim().toUpperCase();
  const lainnya=String(data.alasanLainnya||"").trim();
  if(!nama||!nip||!id||!(jumlah>0)||!["BAIK","RUSAK RINGAN","RUSAK BERAT"].includes(kondisi)||!asal||!tujuan||asal===tujuan||!["PENATAAN RUANGAN","KEBUTUHAN OPERASIONAL","LAINNYA"].includes(alasan)||(alasan==="LAINNYA"&&!lainnya)) throw new Error("Mohon lengkapi seluruh data pemindahan dengan benar.");
  const aset=asetPilihanFormBMN_(id);
  if(aset.lokasi&&asal!==aset.lokasi) throw new Error("Ruang asal harus sesuai lokasi barang pada Master BMN ("+aset.lokasi+").");
  const sh=sheetBMN_("Data_Pemindahan_Barang",["ID_Pemindahan","Waktu_Input","Nama_Penanggung_Jawab_Ruang","NIP","ID_Aset","Nama_Barang_Merk","Jumlah","Kondisi","Ruang_Asal","Ruang_Tujuan","Alasan_Pemindahan","Pernyataan","Status"]);
  const nomor="PDM"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMddHHmmss")+Math.floor(100+Math.random()*900);
  const pernyataan="Dengan ini menyatakan bahwa barang tersebut benar dipindahkan sesuai data di atas dan akan digunakan serta dipelihara dengan penuh tanggung jawab.";
  sh.appendRow([nomor,new Date(),nama,nip,id,aset.label,jumlah,kondisi,asal,tujuan,alasan==="LAINNYA"?lainnya:alasan,pernyataan,"BARU"]);
  return {berhasil:true,pesan:"Formulir pemindahan barang berhasil dikirim.",id:nomor};
}

function kirimPelaporanBarangRusak(data) {
  if(!data||typeof data!=="object") throw new Error("Data formulir tidak valid.");
  const tanggal=String(data.tanggal||"").trim(), pelapor=String(data.namaPelapor||"").trim();
  const id=String(data.idAset||"").trim(), tingkat=String(data.tingkatKerusakan||"").trim().toUpperCase();
  const detail=String(data.detailKerusakan||"").trim(), lokasi=String(data.lokasiBarang||"").trim();
  if(!tanggalISOValidBMN_(tanggal)||!pelapor||!id||!["RUSAK RINGAN","RUSAK BERAT"].includes(tingkat)||!detail||!lokasi) throw new Error("Mohon lengkapi seluruh data pelaporan kerusakan.");
  const aset=asetPilihanFormBMN_(id);
  if(aset.lokasi&&lokasi!==aset.lokasi) throw new Error("Lokasi barang harus sesuai Master BMN ("+aset.lokasi+").");
  const foto=simpanDokumentasiPemeliharaan_(data.foto);
  const sh=sheetBMN_("Data_Pelaporan_Barang_Rusak",["ID_Pelaporan_Rusak","Waktu_Input","Tanggal","Nama_Pelapor","ID_Aset","Nama_Barang_Merk","Tingkat_Kerusakan","Detail_Kerusakan","Foto_URL","Lokasi_Barang","Pernyataan","Status"]);
  const nomor="RSK"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMddHHmmss")+Math.floor(100+Math.random()*900);
  const pernyataan="Saya menyatakan bahwa pengaduan ini disampaikan dengan sebenar-benarnya dan dapat dipertanggungjawabkan.";
  sh.appendRow([nomor,new Date(),new Date(tanggal+"T12:00:00"),pelapor,id,aset.label,tingkat,detail,foto,lokasi,pernyataan,"BARU"]);
  return {berhasil:true,pesan:"Laporan barang rusak berhasil dikirim.",id:nomor,foto:foto};
}

function kirimPelaporanAlatGelasPecah(data) {
  if(!data||typeof data!=="object") throw new Error("Data formulir tidak valid.");
  const pelapor=String(data.namaPelapor||"").trim(), tanggal=String(data.tanggal||"").trim();
  const barang=String(data.namaBarang||"").trim(), jumlah=Number(data.jumlahPecah);
  if(!pelapor||!tanggalISOValidBMN_(tanggal)||!barang||!(jumlah>0)) throw new Error("Mohon lengkapi data pelaporan alat gelas pecah.");
  if(!getReferensiFormPublik().daftarBarang.some(x=>x.jenisForm==="ALAT GELAS"&&x.namaBarang===barang)) throw new Error("Nama alat gelas tidak ditemukan pada referensi ALAT GELAS.");
  const sh=sheetBMN_("Data_Pelaporan_Alat_Gelas_Pecah",["ID_Pelaporan_Pecah","Waktu_Input","Nama_Pelapor","Tanggal","Nama_Barang","Jumlah_Pecah","Pernyataan","Status"]);
  const nomor="GLS"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMddHHmmss")+Math.floor(100+Math.random()*900);
  const pernyataan="Saya bersedia melaporkan setiap kejadian alat gelas yang pecah secara jujur dan bertanggung jawab atas kerusakan yang terjadi akibat penggunaan yang tidak sesuai dengan prosedur.";
  sh.appendRow([nomor,new Date(),pelapor,new Date(tanggal+"T12:00:00"),barang,jumlah,pernyataan,"BARU"]);
  return {berhasil:true,pesan:"Laporan alat gelas pecah berhasil dikirim.",id:nomor};
}

function getDataPemindahanBarangAdmin(token){return dataSheetKeObjekBMN_("Data_Pemindahan_Barang",token);}
function getDataPelaporanBarangRusakAdmin(token){return dataSheetKeObjekBMN_("Data_Pelaporan_Barang_Rusak",token);}
function getDataPelaporanAlatGelasAdmin(token){return dataSheetKeObjekBMN_("Data_Pelaporan_Alat_Gelas_Pecah",token);}

function dataSheetKeObjekBMN_(namaSheet, token) {
  wajibAdminBMN_(token);
  const sh = getDatabase().getSheetByName(namaSheet);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== "")).map(r => {
    const o = {};
    headers.forEach((h,i) => {
      const v = r[i];
      o[h] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : v;
    });
    return o;
  }).reverse();
}

function getDataPenggunaanBarangAdmin(token) {
  return dataSheetKeObjekBMN_("Data_Penggunaan_Barang",token);
}
function getDataUsulanATKAdmin(token) {
  return dataSheetKeObjekBMN_("Data_Usulan_ATK_Operasional",token);
}
function getDataPemeliharaanAdmin(token) {
  return dataSheetKeObjekBMN_("Data_Pemeliharaan_Aset",token);
}
function getRingkasanFormPengelolaanAdmin(token) {
  wajibAdminBMN_(token);
  const ss = getDatabase();
  const count = name => {
    const sh = ss.getSheetByName(name);
    return sh && sh.getLastRow() > 1 ? sh.getLastRow()-1 : 0;
  };
  return {
    penggunaanBarang:count("Data_Penggunaan_Barang"),
    usulanATKOperasional:count("Data_Usulan_ATK_Operasional"),
    pemeliharaanAset:count("Data_Pemeliharaan_Aset")
  };
}
