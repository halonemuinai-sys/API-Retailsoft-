function fetchDailySales(startDate, endDate) {
  var ui = null;
  
  var sheet = getTargetSheet(SHEET_CONFIG.Sales);
  if (!sheet) return;
  
  // Jika dipanggil manual tanpa parameter
  if (!startDate || !endDate) {
    try { ui = SpreadsheetApp.getUi(); } catch(e) {}
    
    var promptBulan = ui.prompt("Tarik Data Bulanan", "Masukkan Bulan (Angka 1-12, contoh: 4 untuk April):", ui.ButtonSet.OK_CANCEL);
    if (promptBulan.getSelectedButton() !== ui.Button.OK || !promptBulan.getResponseText()) return;
    var bulan = parseInt(promptBulan.getResponseText().trim());
    
    var promptTahun = ui.prompt("Tarik Data Bulanan", "Masukkan Tahun (contoh: 2026):", ui.ButtonSet.OK_CANCEL);
    if (promptTahun.getSelectedButton() !== ui.Button.OK || !promptTahun.getResponseText()) return;
    var tahun = parseInt(promptTahun.getResponseText().trim());
    
    if (isNaN(bulan) || isNaN(tahun) || bulan < 1 || bulan > 12) {
      ui.alert("Error", "Input bulan atau tahun tidak valid.", ui.ButtonSet.OK);
      return;
    }

    // Hitung hari pertama dan terakhir di bulan tersebut
    var tglMulai = new Date(tahun, bulan - 1, 1);
    var tglAkhir = new Date(tahun, bulan, 0); 
    
    startDate = Utilities.formatDate(tglMulai, Session.getScriptTimeZone(), "yyyy-MM-dd");
    endDate = Utilities.formatDate(tglAkhir, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  
  var apiUrlSales = "http://139.99.102.231:8089/demo/dailysalestransaction?startdate=" + encodeURIComponent(startDate) + "&enddate=" + encodeURIComponent(endDate);
  
  try {
    var options = {
      "method": "get",
      "muteHttpExceptions": true,
      "headers": { "Authorization": API_TOKEN }
    };
    var response = UrlFetchApp.fetch(apiUrlSales, options);
    var res = JSON.parse(response.getContentText());
    
    if (res && Array.isArray(res) && res.length > 0) {
      // Ambil data transaksi yang sudah ada di Sheet untuk menghindari duplikat
      var lastRow = sheet.getLastRow();
      var existingTx = {};
      if (lastRow > 1) {
        var txColumn = sheet.getRange(2, 6, lastRow - 1, 1).getValues(); // Transaction No ada di Kolom F (6)
        txColumn.forEach(function(r) { existingTx[r[0]] = true; });
      }
      
      var newRows = [];
      for (var i = 0; i < res.length; i++) {
        var itm = res[i];
        if (existingTx[itm.transactionNo]) continue;
        
        var qtyRaw = parseInt(itm.qty);
        var qtyForCalc = (isNaN(qtyRaw) || qtyRaw === 0) ? 1 : qtyRaw;
        var unitPrice = parseFloat(itm.price) || 0;
        var gross = qtyForCalc * unitPrice;
        var discount = parseFloat(itm.subTotalDiscount) || 0;
        var taxAmount = parseFloat(itm.subTotaltax || itm.subTotalTax) || 0;
        var taxRate = parseFloat(itm.tax) || 0;
        var collCode = (itm.collectionCode || "").toUpperCase();
        var grossAfterDiscount = gross - discount;

        // Opsi 1 Agregasi/Cleansing: Jika subTotaltax 0 (biasanya di PFM) tapi harga sudah termasuk pajak
        if (taxAmount === 0 && (collCode === "PFM" || taxRate > 1)) {
           var divisor = taxRate > 1 ? taxRate : 1.11; // Default PPN 11% (1.11) jika tidak ada
           var calculatedNet = grossAfterDiscount / divisor;
           taxAmount = grossAfterDiscount - calculatedNet;
           itm.subTotaltax = taxAmount; // Timpa data mentah agar tersimpan di Excel
        }

        var netSales = grossAfterDiscount - taxAmount;

        newRows.push([
          itm.transactionDate || "",
          itm.transactionTime || "",
          itm.salesman || "",
          itm.customerName || "",
          itm.phoneNo || "",
          itm.transactionNo || "",
          itm.location || "",
          itm.sapCode || "",
          itm.caseNo || "",
          itm.catalogueCode || "",
          itm.description || "",
          itm.collectionCode || "",
          itm.qty || 0,
          itm.price || 0,
          itm.discount || "",
          itm.subTotalDiscount || 0,
          itm.tax || 0,
          itm.subTotaltax || 0,
          netSales
        ]);
      }
      
      if (newRows.length > 0) {
        sheet.getRange(lastRow + 1, 1, newRows.length, 19).setValues(newRows);
        buildMonthlyIndexSilent(); // Gunakan versi silent agar aman di Web App
        if (ui) ui.alert("Berhasil", "Ditambahkan " + newRows.length + " data transaksi baru.", ui.ButtonSet.OK);
      } else {
        if (ui) ui.alert("Info", "Semua data sudah ada di internal.", ui.ButtonSet.OK);
      }
    } else {
      if (ui) ui.alert("Info", "Tidak ada data baru.", ui.ButtonSet.OK);
    }
  } catch (e) {
    if (ui) ui.alert("Error", e.toString(), ui.ButtonSet.OK);
  }
}

// Fungsi Tarik Data Tahunan (Bacthing per Bulan agar tidak timeout)
function fetchYearlySales() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Sales);
  if (!sheet) return;

  var promptTahun = ui.prompt("Tarik Data Tahunan", "Masukkan Tahun (contoh: 2026):", ui.ButtonSet.OK_CANCEL);
  if (promptTahun.getSelectedButton() !== ui.Button.OK || !promptTahun.getResponseText()) return;
  var tahun = parseInt(promptTahun.getResponseText().trim());
  
  if (isNaN(tahun)) {
    ui.alert("Error", "Tahun tidak valid.", ui.ButtonSet.OK);
    return;
  }

  ui.alert("Info", "Sistem akan menarik data selama 1 Tahun penuh secara bertahap (Januari - Desember). Proses ini mungkin memakan waktu 10-30 detik. Silakan klik OK dan tunggu hingga ada notifikasi Selesai.", ui.ButtonSet.OK);

  var lastRow = sheet.getLastRow();
  var existingTx = {};
  if (lastRow > 1) {
    var txColumn = sheet.getRange(2, 6, lastRow - 1, 1).getValues(); 
    txColumn.forEach(function(r) { existingTx[r[0]] = true; });
  }

  var allNewRows = [];
  
  for (var m = 1; m <= 12; m++) {
    var startD = new Date(tahun, m - 1, 1);
    var endD = new Date(tahun, m, 0); 
    
    var startDate = Utilities.formatDate(startD, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var endDate = Utilities.formatDate(endD, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    var apiUrlSales = "http://139.99.102.231:8089/demo/dailysalestransaction?startdate=" + encodeURIComponent(startDate) + "&enddate=" + encodeURIComponent(endDate);
    
    try {
      var options = { "method": "get", "muteHttpExceptions": true, "headers": { "Authorization": API_TOKEN } };
      var response = UrlFetchApp.fetch(apiUrlSales, options);
      var res = JSON.parse(response.getContentText());
      
      if (res && Array.isArray(res)) {
        for (var i = 0; i < res.length; i++) {
          var itm = res[i];
          if (existingTx[itm.transactionNo]) continue;
          
          var qtyRaw = parseInt(itm.qty);
          var qtyForCalc = (isNaN(qtyRaw) || qtyRaw === 0) ? 1 : qtyRaw;
          var unitPrice = parseFloat(itm.price) || 0;
          var gross = qtyForCalc * unitPrice;
          var discount = parseFloat(itm.subTotalDiscount) || 0;
          var taxAmount = parseFloat(itm.subTotaltax || itm.subTotalTax) || 0;
          var taxRate = parseFloat(itm.tax) || 0;
          var collCode = (itm.collectionCode || "").toUpperCase();
          var grossAfterDiscount = gross - discount;

          if (taxAmount === 0 && (collCode === "PFM" || taxRate > 1)) {
             var divisor = taxRate > 1 ? taxRate : 1.11; 
             var calculatedNet = grossAfterDiscount / divisor;
             taxAmount = grossAfterDiscount - calculatedNet;
             itm.subTotaltax = taxAmount; 
          }

          var netSales = grossAfterDiscount - taxAmount;

          allNewRows.push([
            itm.transactionDate || "",
            itm.transactionTime || "",
            itm.salesman || "",
            itm.customerName || "",
            itm.phoneNo || "",
            itm.transactionNo || "",
            itm.location || "",
            itm.sapCode || "",
            itm.caseNo || "",
            itm.catalogueCode || "",
            itm.description || "",
            itm.collectionCode || "",
            itm.qty || 0,
            itm.price || 0,
            itm.discount || "",
            itm.subTotalDiscount || 0,
            itm.tax || 0,
            itm.subTotaltax || 0,
            netSales
          ]);
        }
      }
    } catch(e) {
      // Abaikan jika 1 bulan error agar bulan lain tetap jalan
    }
  }

  if (allNewRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, allNewRows.length, 19).setValues(allNewRows);
    buildMonthlyIndexSilent(); // Gunakan versi silent
    if (ui) ui.alert("Selesai", "Berhasil menarik dan menyimpan " + allNewRows.length + " transaksi baru untuk tahun " + tahun, ui.ButtonSet.OK);
  } else {
    if (ui) ui.alert("Info", "Tidak ada data baru untuk tahun " + tahun + " atau semua data sudah tersinkronisasi.", ui.ButtonSet.OK);
  }
}

function resetDataMart() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Sales);
  if (!sheet) return;
  
  var confirm = ui.alert("Peringatan", "Fungsi ini akan MENGHAPUS seluruh data di Sheet '6. Daily Sales' dan mereset Header.\n\nApakah Anda yakin?", ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  
  sheet.clear();
  var headers = [
    "Date", "Time", "Salesman", "Customer Name", "Phone No", "Transaction No", 
    "Location", "SAP Code", "Case No", "Catalogue Code", "Description", 
    "Collection", "Qty", "Price", "Discount", "Sub Total Discount", "Tax", "Sub Total Tax", "Net Sales"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
  ui.alert("Berhasil", "Struktur Data Mart sudah direset. Silakan jalankan kembali fungsi penarikan data Anda.", ui.ButtonSet.OK);
}

// ========================================================
// ENGINE INDEXING (AGGREGATION)
// ========================================================
function buildMonthlyIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSales = ss.getSheetByName(SHEET_CONFIG.Sales);
  var sheetSummary = ss.getSheetByName(SHEET_CONFIG.MonthlySummary);
  
  if (!sheetSales || !sheetSummary) return;

  var data = sheetSales.getDataRange().getValues();
  if (data.length < 2) return; // Belum ada data

  var rows = data.slice(1);
  var summaryObj = {}; // Key: "YYYY-MM|Location"

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var dateVal = row[0];
    if (!dateVal) continue;
    
    var rowDate = new Date(dateVal);
    if (isNaN(rowDate.getTime())) continue;

    var year = rowDate.getFullYear();
    var month = ("0" + (rowDate.getMonth() + 1)).slice(-2);
    var period = "'" + year + "-" + month;
    var location = row[6] ? String(row[6]) : "Unknown";
    
    var collCode = row[11] ? String(row[11]).toUpperCase() : "";
    var qty = parseInt(row[12]) || 0;
    var netSales = parseFloat(row[18]) || 0; // Kolom S
    var txNo = row[5] ? String(row[5]) : "";

    var key = period + "|" + location;
    
    if (!summaryObj[key]) {
      summaryObj[key] = {
        period: period,
        location: location,
        netSales: 0,
        qty: 0,
        txSet: {}
      };
    }

    if (collCode.indexOf("PACK") > -1 || collCode.indexOf("SVC") > -1 || collCode.indexOf("DPS") > -1) {
      // Abaikan
    } else {
      summaryObj[key].netSales += netSales;
      summaryObj[key].qty += qty;
    }

    if (txNo) summaryObj[key].txSet[txNo] = true;
  }

  // Convert to Array
  var summaryRows = [];
  for (var k in summaryObj) {
    var item = summaryObj[k];
    var txCount = Object.keys(item.txSet).length;
    summaryRows.push([
      item.period,
      item.location,
      item.netSales,
      item.qty,
      txCount
    ]);
  }

  // Sort descending by Period
  summaryRows.sort(function(a, b) {
    if (a[0] > b[0]) return -1;
    if (a[0] < b[0]) return 1;
    return 0;
  });

  // Write to Sheet
  sheetSummary.clearContents();
  // Rewrite Header
  var headers = ["Period", "Location", "Total Net Sales", "Total Qty", "Total Transactions"];
  sheetSummary.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  if (summaryRows.length > 0) {
    sheetSummary.getRange(2, 1, summaryRows.length, headers.length).setValues(summaryRows);
  }
  
  // Keamanan UI Context
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert("Sync Selesai", "Data Monthly Summary (6A) berhasil diperbarui.", ui.ButtonSet.OK);
    }
  } catch(e) {
    // Abaikan jika dijalankan dari Web App
  }
}

/**
 * Versi silent tanpa UI Alert untuk dipanggil dari Web App
 */
function buildMonthlyIndexSilent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSales = ss.getSheetByName(SHEET_CONFIG.Sales);
  var sheetSummary = ss.getSheetByName(SHEET_CONFIG.MonthlySummary);
  
  if (!sheetSales || !sheetSummary) return;

  var data = sheetSales.getDataRange().getValues();
  if (data.length < 2) return; 

  var rows = data.slice(1);
  var summaryObj = {}; 

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var dateVal = row[0];
    if (!dateVal) continue;
    
    var rowDate = new Date(dateVal);
    if (isNaN(rowDate.getTime())) continue;

    var year = rowDate.getFullYear();
    var month = ("0" + (rowDate.getMonth() + 1)).slice(-2);
    var period = "'" + year + "-" + month;
    var location = row[6] ? String(row[6]) : "Unknown";
    
    var collCode = row[11] ? String(row[11]).toUpperCase() : "";
    var qty = parseInt(row[12]) || 0;
    var netSales = parseFloat(row[18]) || 0; 
    var txNo = row[5] ? String(row[5]) : "";

    var key = period + "|" + location;
    
    if (!summaryObj[key]) {
      summaryObj[key] = {
        period: period,
        location: location,
        netSales: 0,
        qty: 0,
        txSet: {}
      };
    }

    if (collCode.indexOf("PACK") > -1 || collCode.indexOf("SVC") > -1 || collCode.indexOf("DPS") > -1) {
    } else {
      summaryObj[key].netSales += netSales;
      summaryObj[key].qty += qty;
    }

    if (txNo) summaryObj[key].txSet[txNo] = true;
  }

  var summaryRows = [];
  for (var k in summaryObj) {
    var item = summaryObj[k];
    var txCount = Object.keys(item.txSet).length;
    summaryRows.push([
      item.period,
      item.location,
      item.netSales,
      item.qty,
      txCount
    ]);
  }

  summaryRows.sort(function(a, b) {
    if (a[0] > b[0]) return -1;
    if (a[0] < b[0]) return 1;
    return 0;
  });

  sheetSummary.clearContents();
  var headers = ["Period", "Location", "Total Net Sales", "Total Qty", "Total Transactions"];
  sheetSummary.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  if (summaryRows.length > 0) {
    sheetSummary.getRange(2, 1, summaryRows.length, headers.length).setValues(summaryRows);
  }
}

// Fungsi Trigger Otomatis
function autoSyncDailySales() {
  var today = new Date();
  var start = new Date();
  start.setDate(today.getDate() - 3); // Cek 3 hari ke belakang untuk sinkronisasi
  
  var startStr = Utilities.formatDate(start, "GMT+7", "yyyy-MM-dd");
  var endStr = Utilities.formatDate(today, "GMT+7", "yyyy-MM-dd");
  
  fetchDailySales(startStr, endStr);
}


