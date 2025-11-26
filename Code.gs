////////////////////////////////////////KODE ISI FORMULIR////////////////////////////////////////////
const FORM_ID = 'COPY_ID_GOOGLE_FORMULIR_DI_SINI'; 

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Formulir') // Nama menu utama
      .addItem('Buat Sheet','buatSistemUjian')
      .addItem('Buka Draft Formulir', 'openDraftGoogleFormLink') 
      .addItem('Update Formulir', 'showConfirmationDialog') 
      .addItem('Tautkan/Hapus Respon Google Form', 'hapusSemuaRespons') 
      .addToUi();
}
// Setelah Menekan tombol Buat Sheet, Menu nya hilang
function onOpen1() {
  SpreadsheetApp.getUi()
      .createMenu('Formulir') // Nama menu utama
      .addItem('Buka Draft Formulir', 'openDraftGoogleFormLink') 
      .addItem('Update Formulir', 'showConfirmationDialog') 
      .addItem('Tautkan/Hapus Respon Google Form', 'hapusSemuaRespons') 
      .addToUi();
}

function buatSistemUjian() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Sheet "Input Soal" ---
  var sheetInput = createOrGetSheet(ss, "Input Soal");
  sheetInput.clear();
  sheetInput.getRange("A1:H1").setValues([
    ["No", "Jenis Soal", "Pertanyaan", "Jawaban Yang Benar", "Pilihan 1", "Pilihan 2", "Pilihan 3", "Pilihan 4"]
  ]);
  // Isi Nomor 1-40 dan "PG"
  var dataPG = [];
  for (var i = 1; i <= 40; i++) {
    dataPG.push([i, "PG"]);
  }
  sheetInput.getRange(2, 1, 40, 2).setValues(dataPG);
  // Isi Nomor 1-5 dan "Esai"
  var dataEsai = [];
  for (var j = 1; j <= 5; j++) {
    dataEsai.push([j, "Esai"]);
  }
  sheetInput.getRange(42, 1, 5, 2).setValues(dataEsai);
  
  //Hapus Kolom dan Baris (Sisakan 60 baris, 8 Kolom A-H)
  hapusBarisKolomKosong(sheetInput, 46, 8);

  //Proteksi Sheet Input Soal
  var pInput = sheetInput.protect().setDescription("Proteksi Sheet");
  var rangeInputBoleh = sheetInput.getRange("C2:H46"); // Area Pertanyaan boleh diedit
  pInput.setUnprotectedRanges([rangeInputBoleh]);
  aturProteksiHanyaSaya(pInput); // <--- Set Hanya Anda

  // --- 2. Sheet "Data Peserta" ---
  var sheetPeserta = createOrGetSheet(ss, "Data Peserta");
  sheetPeserta.clear();
  // Set Rumus A1. Note: Kita escape tanda kutip dua (") dengan backslash (\)
  sheetPeserta.getRange("B1").setValue("Isi Link Sumber >>");
  sheetPeserta.getRange("B2").setValue("Isi Kelas >>");
  sheetPeserta.getRange("A1").setValue("PILIH NAMA")
  sheetPeserta.getRange("A2").setFormula('=IMPORTRANGE(C1;"Data Peserta "&C2&"!A1:A")');
  
  hapusBarisKolomKosong(sheetPeserta, 1000, 3);

  var pPeserta = sheetPeserta.protect().setDescription("Proteksi Sheet");
  var rangePesertaBoleh = sheetPeserta.getRange("B1"); // Area Pertanyaan boleh diedit
  var rangeC1C2 = sheetPeserta.getRange("C1:C2");
  pPeserta.setUnprotectedRanges([rangePesertaBoleh, rangeC1C2]);
  aturProteksiHanyaSaya(pPeserta); // <--- Set Hanya Anda

  
  // --- 3. Sheet "Soal" ---
  var sheetSoal = createOrGetSheet(ss, "Soal");
  sheetSoal.clear();
  sheetSoal.getRange("A1").setValue("DD");
  sheetSoal.getRange("B1").setValue("PILIH NAMA");
  sheetSoal.getRange("C1").setValue("Copy di Sini");
  sheetSoal.getRange("A2").setFormula("=QUERY(ARRAYFORMULA(TO_TEXT('Input Soal'!B2:H)); \"Select *\"; -1)");
  sheetSoal.getRange("H2").setFormula("=arrayformula(IF(ISBLANK(B2:B);\"\";IF(A2:A=\"\";\"\";1)))");
  sheetSoal.getRange("I2").setFormula("=arrayformula(if(ISBLANK(B2:B);\"\";if(A2:A=\"\";\"\";\"Soal \"&'Input Soal'!B2:B&\" \"&'Input Soal'!A2:A)))");

  hapusBarisKolomKosong(sheetSoal, 46, 9);

  var pSoal = sheetSoal.protect().setDescription("Proteksi Sheet");
  aturProteksiHanyaSaya(pSoal); // <--- Set Hanya Anda

  // --- 4. Sheet "Rekap" ---
  var sheetRekap = createOrGetSheet(ss, "Rekap");
  sheetRekap.clear();
  
  sheetRekap.getRange("A1").setValue("SHEET RESPON");
  sheetRekap.getRange("A2").setValue("Form Responses"); // Biasanya default Google Form adalah 'Form Responses 1', sesuaikan jika perlu
  
  // Rumus B1 (Query Form Responses)
  sheetRekap.getRange("B1").setFormula('=QUERY(INDIRECT(A2&"!A:C");"select *";-1)');
  
  // Rumus E1 (Nomor Peserta)
  var rumusE1 = '=iferror(ARRAY_CONSTRAIN(ARRAYFORMULA(if(row(B:B)=1;"NOMOR PESERTA";RIGHT(D:D; LEN(D:D) - FIND("-"; D:D; FIND("-"; D:D)+1)))); 1500; 1);"")';
  sheetRekap.getRange("E1").setFormula(rumusE1);
  
  // Rumus F1 (Nama Murid)
  var rumusF1 = '=iferror(ARRAY_CONSTRAIN(ARRAYFORMULA(if(row(B:B)=1;"NAMA MURID";MID(D:D; FIND("-"; D:D)+1; FIND("-"; D:D; FIND("-"; D:D)+1) - FIND("-"; D:D) - 1))); 1500; 1);"")';
  sheetRekap.getRange("F1").setFormula(rumusF1);
  
  // Rumus G1 (Kelas)
  var rumusG1 = '=iferror(ARRAY_CONSTRAIN(ARRAYFORMULA(if(row(B:B)=1;"KELAS";left(D:D; FIND("-"; D:D)-1))); 1500; 1);"")';
  sheetRekap.getRange("G1").setFormula(rumusG1);
  
  // Rumus H1 (Nilai PG)
  var rumusH1 = '=ARRAY_CONSTRAIN(ARRAYFORMULA(if(row(B:B)=1;"NILAI PG";IF(D1:D=""; "";(VALUE(INDEX(SPLIT(C:C; "/");;1)) /VALUE(INDEX(SPLIT(C:C; "/");;2))) * 100 ))); 800; 1)';
  sheetRekap.getRange("H1").setFormula(rumusH1);
  
  // Rumus I1 (Query Jawaban Esai - Asumsi kolom AR:AV di sheet respon)
  sheetRekap.getRange("I1").setFormula('=QUERY(INDIRECT(A2&"!AR:AV");"select *";-1)');
  
  hapusBarisKolomKosong(sheetRekap, 500, 13);

  var pRekap = sheetRekap.protect().setDescription("Proteksi Sheet");
  aturProteksiHanyaSaya(pRekap); // <--- Set Hanya Anda

  // --- 5. Sheet "Koreksi" ---
  var sheetKoreksi = createOrGetSheet(ss, "Koreksi");
  sheetKoreksi.clear();
  
  // Rumus A1 (Query dari Rekap)
  sheetKoreksi.getRange("A1").setFormula('=QUERY(Rekap!E1:H;"select * Where Col1 is not null";-1)');
  sheetKoreksi.getRange("E1").setValue("NILAI ESAI");
  
  // Rumus F1 (Nilai Akhir)
  var rumusF1Nilai = '=IFERROR(ARRAYFORMULA(if(row(B:B)=1;"NILAI AKHIR";IF(A:A="";IFERROR(1/0);IF(E:E="";D:D;D:D*50%+E:E*50%))));"")';
  sheetKoreksi.getRange("F1").setFormula(rumusF1Nilai);
  
  hapusBarisKolomKosong(sheetKoreksi, 1000, 6);

  var pKoreksi = sheetKoreksi.protect().setDescription("Proteksi Sheet");
  var rangeKoreksiBoleh = sheetKoreksi.getRange("E2:E"); // Area Pertanyaan boleh diedit
  pKoreksi.setUnprotectedRanges([rangeKoreksiBoleh]);
  aturProteksiHanyaSaya(pKoreksi); // <--- Set Hanya Anda

// --- 6. Sheet "Nilai" ---
  var sheetDownloadNilai = createOrGetSheet(ss, "Nilai");
  sheetDownloadNilai.clear();
  sheetDownloadNilai.getRange("A1").setFormula('=QUERY(Koreksi!A1:F;"Select * where Col1 is not null order by Col3, Col2";-1)');

  hapusBarisKolomKosong(sheetDownloadNilai, 1000, 6);

  var pDownloadNilai = sheetDownloadNilai.protect().setDescription("Proteksi Sheet");
  aturProteksiHanyaSaya(pDownloadNilai); // <--- Set Hanya Anda

  onOpen1();
  SpreadsheetApp.getUi().alert("Selesai! Semua sheet dan rumus telah dibuat.");
}

// Fungsi pembantu untuk mengecek apakah sheet sudah ada atau belum
function createOrGetSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    return sheet;
  } else {
    return ss.insertSheet(name);
  }
}

// Fungsi baru untuk mengatur izin "Hanya Anda"
function aturProteksiHanyaSaya(protection) {
  // 1. Pastikan script mengenali Anda (pemilik) sebagai editor
  var me = Session.getEffectiveUser();
  protection.addEditor(me);

  // 2. Hapus semua editor lain yang ada di list (teman/kolega)
  protection.removeEditors(protection.getEditors());

  // 3. Matikan edit domain jika menggunakan akun sekolah/kantor
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

// Menghapus Baris dan Kolom yang tidak digunakan
function hapusBarisKolomKosong(sheet, barisDiinginkan, kolomDiinginkan) {
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  
  // Hapus Baris Berlebih
  if (maxRows > barisDiinginkan) {
    sheet.deleteRows(barisDiinginkan + 1, maxRows - barisDiinginkan);
  }
  
  // Hapus Kolom Berlebih
  if (maxCols > kolomDiinginkan) {
    sheet.deleteColumns(kolomDiinginkan + 1, maxCols - kolomDiinginkan);
  }
}

function showConfirmationDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Konfirmasi Update GForm',
     'Yakin ingin HAPUS SEMUA pertanyaan di GForm (' + FORM_ID + ') ?',
     ui.ButtonSet.YES_NO
  );
  if (result == ui.Button.YES) {
    updateFormWithMixedTypes();
    populateGoogleForms();
    openGoogleFormLink();
    ui.alert('Update GForm Selesai!', 'Formulir telah berhasil diperbarui.', ui.ButtonSet.OK);
  } else {
    Logger.log('Update dibatalkan.');
  }
}

function updateFormWithMixedTypes() {
  Logger.log("-- Mulai Update Form --");
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Soal');

  if (!sheet) {
    Logger.log("ERROR: Sheet bernama 'Soal' tidak ditemukan! Harap periksa nama sheet.");
    return;
  }
  var numberRows = sheet.getLastRow();
  if (numberRows <= 1) {
    Logger.log("ERROR: Sheet tidak memiliki data yang cukup (minimal 2 baris data).");
    return;
  }
  
  // MEMBUKA 9 KOLOM (KOLOM A-I)
  var allData = sheet.getRange(1, 1, numberRows, 9).getValues();
  var totalRows = allData.length;

  // MEMBUKA DAN MEMBERSIHKAN FORMULIR
  try {
    var form = FormApp.openById(FORM_ID);

    // Menghapus semua item/pertanyaan yang sudah ada
    var items = form.getItems();
    for (var i = items.length - 1; i >= 0; i--) {
      form.deleteItem(items[i]);
    }
    Logger.log("Semua item lama berhasil dihapus.");

    // Mengatur formulir menjadi mode Kuis
    form.setIsQuiz(true);

  } catch (e) {
    Logger.log("ERROR KRITIS saat membuka/menghapus item Form: " + e.toString() + ". Pastikan ID dan izin benar.");
    return;
   }
  
  // --- 3. ITERASI DAN MENAMBAHKAN ITEM BARU DENGAN SECTION ---
  for (var i = 0; i < totalRows; i++) {
    var row = allData[i];
    var questionType = row[0].toString().toUpperCase().trim(); // Kolom A: Jenis Soal

    // Bersihkan dan validasi Judul Pertanyaan
    var cleanTitle = row[1] ? row[1].toString().trim() : ""; // Kolom B: Judul Pertanyaan

    if (cleanTitle === "") {
      Logger.log("Baris " + (i + 1) + " dilewati karena judul pertanyaan kosong.");
      continue; 
    }
    row[1] = cleanTitle;

    // Panggil fungsi pembantu untuk menambahkan pertanyaan berdasarkan jenis
    switch (questionType) {
      case 'PG':
        addMultipleChoiceItem(form, row);
        break;
      case 'DD':
        addDropdownItem(form, row);
        break;
      case 'IS':
        addShortAnswerItem(form, row);
        break;
      case 'ESAI':
        addParagraphItem(form, row);
        break;
      default:
        Logger.log("Jenis pertanyaan tidak dikenal di baris " + (i + 1) + ": " + questionType + ". Dilewati.");
    }

    // Judul Section diambil dari Kolom I (Indeks 8) dari baris berikutnya
    if (i < totalRows - 1) { 
      var nextSectionTitle = allData[i+1][8].toString().trim(); // Kolom I (Indeks 8)
      // Fallback jika kolom Judul Section kosong
      if (nextSectionTitle === "") {
        nextSectionTitle = "Lanjut ke Pertanyaan " + (i + 2);
      }
      form.addPageBreakItem()
        .setTitle(nextSectionTitle); 
    }
   }
  Logger.log("--- Update Formulir Selesai ---");
  // CATATAN: Pesan sukses sekarang ditampilkan oleh showConfirmationDialog() 
}

//Poin Pertanyaan diambil dari Kolom H (Indeks 7)
//Menambahkan item Pilihan Ganda (Multiple Choice). (PG)
function addMultipleChoiceItem(form, row) {
  var questionTitle = row[1];
  var myAnswers = row[2];
  var myGuesses = row.slice(2, 7);
  var questionPoint = parseInt(row[7], 10) || 1; // Kolom H (Indeks 7)
  var shuffledOptions = shuffleArray(myGuesses);
  var addItem = form.addMultipleChoiceItem();
  var choices = createChoices(addItem, shuffledOptions, myAnswers);
  addItem.setTitle(questionTitle)
      .setPoints(questionPoint)
      .setChoices(choices);
}

//Menambahkan item Dropdown (List). (DD)
function addDropdownItem(form, row) {
  var questionTitle = row[1];
  var myAnswers = row[2];
  var myGuesses = row.slice(2, 7);
  var shuffledOptions = shuffleArray(myGuesses);
  var addItem = form.addListItem();
  var choices = createChoices(addItem, shuffledOptions, myAnswers);
  addItem.setTitle(questionTitle)
}

//Menambahkan item Isian Singkat (Short Answer/Text). (IS)
function addShortAnswerItem(form, row) {
  var questionTitle = row[1];
  var correctAnswer = row[2].toString().trim();
  var addItem = form.addTextItem();
      addItem.setTitle(questionTitle)
  if (correctAnswer !== "") {
      addItem.setValidation(
      FormApp.createTextValidation()
        .requireTextIsEqualTo(correctAnswer)
        .build()
      );
      var feedback = FormApp.createFeedback().setText('Jawaban yang benar adalah: ' + correctAnswer).build();
    addItem.setCorrectFeedback(feedback)
    .setIncorrectFeedback(feedback);
  }
}

//Menambahkan item Paragraf (Paragraph/Esai). (ESAI)
function addParagraphItem(form, row) {
  var questionTitle = row[1];
  // Item esai tidak secara otomatis diberi poin di sini (Poin default 0)
  var addItem = form.addParagraphTextItem();
  addItem.setTitle(questionTitle);
}

//Fungsi untuk membuat array Choices untuk PG atau DD.
function createChoices(item, shuffledOptions, myAnswers) {
  var choices = [];
  var correctIndex = shuffledOptions.indexOf(myAnswers);
  for (var j = 0; j < shuffledOptions.length; j++) {
    var isCorrect = (j === correctIndex);
    var optionValue = shuffledOptions[j].toString().trim();
      if (optionValue !== "") {
        choices.push(
        item.createChoice(optionValue, isCorrect)
        );
      }
    }
    return choices;
}

//Mengacak elemen dalam array (Algoritma Fisher-Yates).
function shuffleArray(array) {
  var i, j, temp;
  for (i = array.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

/**
Mengambil data dari Spreadsheet dan mengisinya sebagai pilihan (choices)
untuk item-item di Google Form yang memiliki judul yang sama.
*/
const populateGoogleForms = () => {
// 1. Dapatkan Spreadsheet yang aktif
const ss = SpreadsheetApp.getActiveSpreadsheet();
// 2. Buka Google Form menggunakan FORM_ID Menggunakan FormApp.openById() lebih ringkas dan sesuai dengan penggunaan konstanta ID.
const form = FormApp.openById(FORM_ID);
// 3. Ambil semua data dari sheet 'data_peserta'
const sheet = ss.getSheetByName('Data Peserta');
  if (!sheet) {
    throw new Error('Sheet "Data Peserta" tidak ditemukan!');
  }
// [heads, ...data] mendestrukturisasi, 'heads' adalah baris header, 'data' adalah baris data
const [heads, ...data] = sheet.getDataRange().getDisplayValues();
// 4. Siapkan objek 'choices' untuk menyimpan pilihan untuk setiap judul kolom
const choices = {};
  heads.forEach((title, i) => {
    // Ambil semua nilai dari kolom ke-i (kecuali header) dan filter nilai kosong
    choices[title] = data.map((d) => d[i]).filter((e) => e);
  });
// 5. Perbarui item Form dengan pilihan yang sesuai
  form.getItems()
    .map((item) => ({
      item,
      // Cari pilihan berdasarkan Judul Item Form yang harus sama dengan Judul Kolom Spreadsheet
      values: choices[item.getTitle()],
    }))
    .filter(({ values }) => values && values.length > 0) // Hanya proses item yang memiliki data pilihan
    .forEach(({ item, values }) => {
      // Perbarui pilihan berdasarkan Tipe Item
      switch (item.getType()) {
        case FormApp.ItemType.CHECKBOX:
          item.asCheckboxItem().setChoiceValues(values);
          break;
        case FormApp.ItemType.LIST: // Tipe Dropdown
          item.asListItem().setChoiceValues(values);
          break;
        case FormApp.ItemType.MULTIPLE_CHOICE:
          item.asMultipleChoiceItem().setChoiceValues(values);
          break;
        default:
          // Abaikan tipe item lainnya
          Logger.log(`Tipe item "${item.getTitle()}" (${item.getType()}) tidak didukung untuk pembaruan pilihan.`);
      }
    });
  // 6. Beri notifikasi sukses
  ss.toast('Google Form Successfully Updated!');
};

//Membuka tautan draf (edit) Google Form di tab baru.
function openDraftGoogleFormLink() {
  // Membuat URL Draf/Edit dengan FORM_ID
  const draftUrl = `https://docs.google.com/forms/d/${FORM_ID}/edit`;
  var html = HtmlService.createHtmlOutput('<script>window.open("' + draftUrl + '", "_blank");</script>')
      .setWidth(100)
      .setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Membuka Tautan');
}

//Membuka tautan live (viewform) Google Form di tab baru.
function openGoogleFormLink() {
  // Menggunakan FORM_ID untuk mendapatkan objek Form.
  // Dari objek Form, kita bisa mendapatkan URL yang sudah dipublikasikan (viewform).
  try {
    const form = FormApp.openById(FORM_ID);
    const publishedUrl = form.getPublishedUrl(); 
    var html = HtmlService.createHtmlOutput('<script>window.open("' + publishedUrl + '", "_blank");</script>')
        .setWidth(100)
        .setHeight(1);
    SpreadsheetApp.getUi().showModalDialog(html, 'Membuka Tautan');
  } catch (e) {
    // Penanganan kesalahan jika FormApp.openById gagal
    SpreadsheetApp.getUi().alert('Gagal membuka Formulir.', 'Pastikan FORM_ID benar dan Anda memiliki akses yang diperlukan. ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getForm() {
  try {
    const form = FormApp.openById(FORM_ID);
    return form;
  } catch (e) {
    Logger.log("Error membuka formulir: " + e.toString());
    throw new Error("Gagal membuka formulir. Pastikan ID formulir benar dan Anda memiliki izin akses.");
  }
}

function hapusSemuaRespons() {
  const form = getForm();  
  // Metode untuk menghapus semua tanggapan
  form.deleteAllResponses(); 
  Logger.log(`Semua respons (${form.getResponses().length}) dari formulir "${form.getTitle()}" telah dihapus.`);
  unlinkAndRelinkForm();
}

/**
 Membatalkan tautan (unlink) dan menautkan kembali (re-link) formulir yang terhubung
 ke Google Sheet saat ini.
 */
function unlinkAndRelinkForm() {
  // 1. Dapatkan Spreadsheet aktif
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 2. Dapatkan objek Form yang terhubung
  // Ini mengasumsikan hanya ada satu formulir yang terhubung ke Sheet ini.
  const form = ss.getFormUrl() ? FormApp.openByUrl(ss.getFormUrl()) : null;
  if (form) {
    try {
      // --- Langkah 1: Batalkan Tautan (Unlink) ---  
      // Ambil ID dari Spreadsheet saat ini sebelum membatalkan tautan
      const spreadsheetId = ss.getId(); 
      Logger.log(`Membatalkan tautan formulir '${form.getTitle()}' (ID: ${form.getId()}) dari Spreadsheet (ID: ${spreadsheetId})...`);
      // Memanggil metode untuk membatalkan tautan Sheet saat ini.
      form.removeDestination(); 
      /**SpreadsheetApp.getUi().alert('Status', 'Formulir berhasil dibatalkan tautannya (Unlinked).', SpreadsheetApp.getUi().ButtonSet.OK);*/
      Logger.log('Pembatalan tautan berhasil.');
      // --- Langkah 2: Tautkan Kembali (Relink) ---
      Logger.log(`Menautkan kembali formulir ke Spreadsheet (ID: ${spreadsheetId})...`);
      // Menautkan kembali formulir ke Spreadsheet saat ini sebagai tujuan baru.
      // Data respons formulir akan masuk ke Sheet baru yang dibuat di Spreadsheet ini
      // (biasanya bernama "Form Responses 1").
      form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
      SpreadsheetApp.getUi().alert('Selesai', 'Tautan berhasi diperbaruhi, Ganti nama Sheet menjadi "Form Responses"', SpreadsheetApp.getUi().ButtonSet.OK);
      Logger.log('Penautan kembali berhasil.');
      hapusSheetBerdasarkanNama();
      hapusDataSheetKoreksiKolomE();
    } catch (e) {
      Logger.log(`Terjadi kesalahan: ${e.toString()}`);
      SpreadsheetApp.getUi().alert('Kesalahan', `Tidak dapat menyelesaikan operasi: ${e.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } else {
    SpreadsheetApp.getUi().alert('Peringatan', 'Tidak ada formulir yang terhubung, akan dihubungkan terlebih dahulu.', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Tidak ada formulir yang terhubung.');
    tautkanFormulirKeSheetAktif();
    SpreadsheetApp.getUi().alert('Peringatan', 'Ganti nama sheet menjadi "Form Responses"', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Fungsi ini menautkan Google Formulir ke SPREADSHEET YANG AKTIF
 * (spreadsheet tempat skrip ini dijalankan).
 */
function tautkanFormulirKeSheetAktif() {
  try {
    // 1. Membuka Formulir berdasarkan ID
    const form = FormApp.openById(FORM_ID);
    
    // 2. Mendapatkan Spreadsheet yang sedang aktif saat ini
    // Skrip ini HARUS dijalankan dari dalam file Sheet yang dituju
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 3. Menetapkan Spreadsheet ini sebagai tujuan respons
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    
    // 4. (Opsional) Mencatat log sukses ke konsol
    Logger.log('Formulir "' + form.getTitle() + '" telah berhasil ditautkan ke Spreadsheet aktif: ' + spreadsheet.getName());
    Logger.log('URL Spreadsheet: ' + spreadsheet.getUrl());
    
  } catch (e) {
    // Menangani jika ID Formulir salah atau ada masalah izin
    Logger.log('Error: ' + e.message);
    if (e.message.includes("Cannot call SpreadsheetApp.getActiveSpreadsheet()")) {
        Logger.log('PENTING: Skrip ini harus dijalankan dari dalam file Google Sheet.');
    }
  }
}

function hapusSheetBerdasarkanNama() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const namaSheet = "Form Responses"; // Ganti dengan nama yang ingin dihapus
  
  // 1. Cari sheet berdasarkan nama
  const sheetTarget = spreadsheet.getSheetByName(namaSheet);

  // 2. Cek apakah sheet tersebut ada?
  if (sheetTarget) {
    spreadsheet.deleteSheet(sheetTarget);
  } else {
    Logger.log("Sheet dengan nama '" + namaSheet + "' tidak ditemukan.");
  }
}

function hapusDataSheetKoreksiKolomE() {
  // 1. Ambil sheet bernama "Koreksi"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Koreksi');
  
  if (sheet) {
    // 2. Ambil range E2 sampai E paling bawah
    // clearContent() = hanya hapus teks/angka, format tetap
    sheet.getRange("E2:E").clearContent();
    
    Logger.log("Data E2:E di sheet Koreksi berhasil dihapus.");
  } else {
    Logger.log("Error: Sheet 'Koreksi' tidak ditemukan.");
  }
}
////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////KODE ADMIN//////////////////////////////////

// Code.gs

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aplikasi Koreksi Esai')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi 1: Mengambil data dari sheet "Koreksi" untuk ditampilkan di tabel
function getDataKoreksi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Koreksi");
  
  // Cek apakah sheet ada?
  if (!sheet) {
    throw new Error("Sheet 'Koreksi' tidak ditemukan! Periksa nama tab di bawah.");
  }

  var lastRow = sheet.getLastRow();
  
  // Jika baris data kurang dari 2 (hanya header atau kosong), kembalikan array kosong
  if (lastRow < 2) {
    return []; 
  }
  
  // Ambil data dari baris 2 sampai baris terakhir, kolom 1 s/d 6
  // Format: getRange(row, column, numRows, numColumns)
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues(); 
  // Catatan: Saya ganti getValues() jadi getDisplayValues() agar format angka/tanggal sesuai tampilan sheet
  
  return data;
}
// Fungsi 2: Menyimpan perubahan Nilai Esai
function simpanSemuaNilai(dataNilai) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Koreksi");
  var dataSheet = sheet.getDataRange().getValues(); // Ambil semua data untuk pencocokan
  
  // dataNilai adalah array objek: [{id: 'P001', nilai: 80}, ...]
  
  // Loop data yang dikirim dari HTML
  dataNilai.forEach(function(item) {
    // Cari baris yang cocok dengan Nomor Peserta (ID)
    for (var i = 1; i < dataSheet.length; i++) { // Mulai i=1 karena i=0 adalah header
      if (dataSheet[i][0] == item.id) { // Kolom A (indeks 0) adalah Nomor Peserta
        // Update Kolom E (indeks 4 adalah kolom ke-5/NILAI ESAI)
        // i + 1 karena baris di getRange mulai dari 1
        sheet.getRange(i + 1, 5).setValue(item.nilai); 
        break; 
      }
    }
  });
  SpreadsheetApp.flush();
  return "Berhasil disimpan!";
}

// Fungsi 3: Mengambil Soal dan Jawaban Esai dari sheet "Rekap"
function getDetailEsai(nomorPeserta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Rekap");
  
  // 1. Ambil Soal (Header I1 - M1)
  var soalRange = sheet.getRange("I1:M1").getValues()[0];
  
  // 2. Cari Jawaban berdasarkan Nomor Peserta
  var data = sheet.getDataRange().getValues();
  var jawabanSiswa = [];
  
  for (var i = 1; i < data.length; i++) {
    // Asumsi Kolom A di sheet Rekap adalah Nomor Peserta
    if (data[i][4] == nomorPeserta) {
      // Ambil jawaban dari kolom I (indeks 8) sampai M (indeks 12)
      // slice mengambil dari indeks awal sampai (sebelum) indeks akhir
      jawabanSiswa = data[i].slice(8, 13); 
      break;
    }
  }
  
  if (jawabanSiswa.length === 0) {
    return null; // Data tidak ditemukan
  }
  
  return {
    soal: soalRange,
    jawaban: jawabanSiswa
  };
}

function dapatkanLinkDownload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai"); // Pastikan nama sheet sesuai (Koreksi/Nilai)
  
  if (!sheet) return null;

  var ssId = ss.getId();
  var sheetId = sheet.getSheetId();
  
  // Membuat URL Download Format Excel (.xlsx)
  var url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx&gid=" + sheetId;
  
  return url;
}
