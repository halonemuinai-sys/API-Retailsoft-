function fetchCatalogProduct() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Catalog);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Info", "Pencarian dibatalkan.\nPastikan Anda sudah mengisi 'Keyword' di Kolom A (mulai baris 2).", ui.ButtonSet.OK);
    return;
  }

  var keywords = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var hasil = [];

  for (var i = 0; i < keywords.length; i++) {
    var word = keywords[i][0];
    if (!word || word.toString().trim() === "") { 
      hasil.push(["", "", "", ""]); 
      continue; 
    }
    
    var res = panggilAPI("/catalogproduct2?keyword=" + encodeURIComponent(word) + "&category=&page=1&limit=20");
    if (res && res.data && res.data.length > 0) {
      var itm = res.data[0]; 
      hasil.push([itm.itemName || "", itm.itemCode || "", itm.itemSku || "", itm.itemPrice || ""]); 
    } else if (res && res.error) {
      hasil.push(["(" + res.error + ")", "", "", ""]);
    } else {
      hasil.push(["(Tidak ditemukan / Kosong)", "", "", ""]);
    }
  }
  
  if(hasil.length > 0) {
    sheet.getRange(2, 2, hasil.length, 4).setValues(hasil);
    ui.alert("Selesai", "Data Catalog berhasil ditarik dan diupdate dalam Sheet!", ui.ButtonSet.OK);
  }
}

function fetchProductDetail() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Detail);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Info", "Penarikan dibatalkan.\nPastikan Anda sudah mengisi 'Item Code' di Kolom A (mulai baris 2).", ui.ButtonSet.OK);
    return;
  }
  
  var itemCodes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var hasil = [];

  for (var i = 0; i < itemCodes.length; i++) {
    var code = itemCodes[i][0];
    if (!code || code.toString().trim() === "") { 
      hasil.push(["", "", "", ""]); 
      continue; 
    }
    
    var res = panggilAPI("/itemview?itemCode=" + encodeURIComponent(code));
    if (res && res.data && res.data.length > 0) {
      var prod = res.data[0];
      hasil.push([prod.itemName || "", prod.description || "", prod.itemPrice || "", prod.kategori || ""]);
    } else if (res && res.error) {
      hasil.push(["(" + res.error + ")", "", "", ""]);
    } else {
      hasil.push(["(Tidak ditemukan / Kosong)", "", "", ""]);
    }
  }
  
  if(hasil.length > 0) {
    sheet.getRange(2, 2, hasil.length, 4).setValues(hasil);
    ui.alert("Selesai", "Data Detail Produk berhasil ditarik dan diupdate dalam Sheet!", ui.ButtonSet.OK);
  }
}

function fetchStockLocation() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Stock);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Info", "Penarikan dibatalkan.\nPastikan Anda sudah mengisi 'Item SKU' di Kolom A, dan 'Item Name' di Kolom B.", ui.ButtonSet.OK);
    return;
  }
  
  var inputs = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var hasil = [];

  for (var i = 0; i < inputs.length; i++) {
    var sku = inputs[i][0];
    var nama = inputs[i][1];
    
    if (!sku || !nama || sku.toString().trim() === "" || nama.toString().trim() === "") { 
      hasil.push(["", ""]); 
      continue; 
    }
    
    var res = panggilAPI("/stocklocation?itemSku=" + encodeURIComponent(sku) + "&itemName=" + encodeURIComponent(nama));
    if (res && res.data && res.data.length > 0) {
      var lokasi = [], total = 0;
      for (var j = 0; j < res.data.length; j++) {
         lokasi.push((res.data[j].locationName || "") + " ("+ (res.data[j].qty || 0) +")");
         total += (res.data[j].qty || 0);
      }
      hasil.push([lokasi.join(", "), total]);
    } else if (res && res.error) {
      hasil.push(["(" + res.error + ")", 0]);
    } else {
      hasil.push(["(Stok Kosong)", 0]);
    }
  }
  
  if(hasil.length > 0) {
    sheet.getRange(2, 3, hasil.length, 2).setValues(hasil); 
    ui.alert("Selesai", "Data Lokasi Stok berhasil ditarik dan diupdate dalam Sheet!", ui.ButtonSet.OK);
  }
}

