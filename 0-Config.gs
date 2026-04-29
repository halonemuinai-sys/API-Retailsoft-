// Config object to map API endpoints to specific sheet names
var SHEET_CONFIG = {
  "Catalog": "1. Catalog",
  "Detail": "2. Detail",
  "Stock": "3. Stock Location",
  "Summary": "4. Cat Summary",
  "Location": "5. Cat Location",
  "Sales": "6. Daily Sales",
  "MonthlySummary": "6A. Monthly Summary"
};

var CONFIG_CRM = {
  "PROFILING_SS_ID": "17dIze7RwnA4nqxCVRbTeDIDlwCmW1OvgXQ9zMOM-ovM",
  "T_SHEET_NAME": "Traffic",
  "APP_TITLE": "ARES CRM"
};

var BASE_URL = "http://139.99.102.231:8189/api";

// 🔴 SILAKAN ISI TOKEN ANDA DI SINI
// Masukkan token asli Anda (misalnya hasil copy-paste dari aplikasi postman/auth Anda)
// Contoh: "Bearer eyJhbGciOiJIUzI1NiIsInR..."
var API_TOKEN = "Bearer MASUKKAN_TOKEN_ANDA_DISINI";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Tarik Data Bvlgari')
      .addItem('1. Install / Perbaiki Struktur Sheet', 'setupSemuaSheet')
      .addSeparator()
      .addItem('2. Tarik Data: Catalog', 'fetchCatalogProduct')
      .addItem('3. Tarik Data: Detail', 'fetchProductDetail')
      .addItem('4. Tarik Data: Stock', 'fetchStockLocation')
      .addItem('5. Tarik Data: Category Summary', 'fetchCategorySummary')
      .addItem('6. Tarik Data: Category Location', 'fetchCategoryLocation')
      .addItem('7. Tarik Data: Daily Sales', 'fetchDailySales')
      .addItem('8. Update: Monthly Summary (Sync YTD)', 'buildMonthlyIndex')
      .addSeparator()
      .addItem('Cek Error API (Test)', 'testerAPI')
      .addToUi();
}

function setupSemuaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var configs = [
    { nama: SHEET_CONFIG.Catalog, headers: ["Keyword (Input A)", "Nama Item", "Item Code", "Item SKU", "Harga"] },
    { nama: SHEET_CONFIG.Detail, headers: ["Item Code (Input A)", "Nama Item", "Deskripsi Lengkap", "Harga", "Kategori"] },
    { nama: SHEET_CONFIG.Stock, headers: ["Item SKU (Input A)", "Item Name (Input B)", "Daftar Lokasi & Qty", "Total Semua Cabang"] },
    { nama: SHEET_CONFIG.Summary, headers: ["Category ID (Input A)", "Category Name", "Total Qty", "Total Value"] },
    { nama: SHEET_CONFIG.Location, headers: ["Category ID (Input A)", "Daftar Lokasi & Qty", "Total Qty"] },
    { nama: SHEET_CONFIG.Sales, headers: ["Transaction Date", "Time", "Salesman", "Customer Name", "Phone No", "Transaction No", "Location", "SAP Code", "Case No", "Catalogue Code", "Description", "Collection", "Qty", "Price", "Discount", "Sub Total Discount", "Tax", "Sub Total Tax", "Net Sales"] },
    { nama: SHEET_CONFIG.MonthlySummary, headers: ["Period", "Location", "Total Net Sales", "Total Qty", "Total Transactions"] }
  ];

  for (var i = 0; i < configs.length; i++) {
    var conf = configs[i];
    var sheet = ss.getSheetByName(conf.nama);
    
    if (!sheet) {
      sheet = ss.insertSheet(conf.nama);
    }
    
    // Warnai header dan bekukan baris pertama
    sheet.getRange(1, 1, 1, conf.headers.length).setValues([conf.headers])
         .setFontWeight("bold")
         .setBackground("#D9EAD3");
    
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200); 
  }
  
  ui.alert("Setup Selesai!", "5 Tab Sheet beserta judul kolomnya sudah berhasil dibuat.", ui.ButtonSet.OK);
}

