function panggilAPI(endpoint) {
  try {
    var options = {
      "method": "get",
      "muteHttpExceptions": true,
      "headers": {
        "Authorization": API_TOKEN
      }
    };
    var response = UrlFetchApp.fetch(BASE_URL + endpoint, options);
    var code = response.getResponseCode();
    if (code === 200) {
      var text = response.getContentText();
      if(!text) return { error: "Data Kosong" };
      return { data: JSON.parse(text) };
    } else {
      return { error: "Error HTTP " + code };
    }
  } catch (e) {
    return { error: "Gagal Koneksi" };
  }
}

function getTargetSheet(namaSheetIdaman) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(namaSheetIdaman);
  try {
    var ui = SpreadsheetApp.getUi();
    if (!sheet && ui) {
      ui.alert("Sheet tidak ditemukan", "Pilih menu 'Install / Perbaiki Struktur Sheet' terlebih dahulu.", ui.ButtonSet.OK);
    }
  } catch(e) {
    // Abaikan jika dijalankan dari Web App
  }
  return sheet;
}

function testerAPI() {
  var ui = SpreadsheetApp.getUi();
  var urlTes = BASE_URL + "/catalogproduct2?keyword=683W6F&category=&page=1&limit=20";
  
  try {
    var options = {
      "method": "get",
      "muteHttpExceptions": true,
      "headers": {
        "Authorization": API_TOKEN
      }
    };
    var response = UrlFetchApp.fetch(urlTes, options);
    var kodeStatus = response.getResponseCode();
    var teksRespon = response.getContentText();
    
    var cuplikan = teksRespon.length > 500 ? teksRespon.substring(0, 500) + "..." : teksRespon;
    
    ui.alert(
      "Hasil Tes API",
      "Status Kode: " + kodeStatus + "\n\nRespon:\n" + cuplikan,
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert("Gagal Total", "Google Script gagal menyambung ke IP. Pesan Error:\n" + e.toString(), ui.ButtonSet.OK);
  }
}
