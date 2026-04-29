function fetchCategorySummary() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Summary);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Info", "Penarikan dibatalkan.\nPastikan Anda sudah mengisi 'Category ID' di Kolom A.", ui.ButtonSet.OK);
    return;
  }
  
  var catIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var hasil = [];

  for (var i = 0; i < catIds.length; i++) {
    var id = catIds[i][0];
    if (!id || id.toString().trim() === "") { 
      hasil.push(["", "", ""]); 
      continue; 
    }
    
    var res = panggilAPI("/categorysummary?categoryId=" + encodeURIComponent(id));
    if (res && res.data && res.data.length > 0) {
      var catName = res.data[0].kategori || "Tanpa Nama";
      var tQty = res.data.reduce(function(sum, x) { return sum + (x.qty || 0); }, 0);
      var tVal = res.data.reduce(function(sum, x) { return sum + ((x.qty || 0) * (x.itemPrice || 0)); }, 0);
      
      hasil.push([catName, tQty, tVal]);
    } else if (res && res.error) {
      hasil.push(["(" + res.error + ")", 0, 0]);
    } else {
      hasil.push(["(Data Kosong)", 0, 0]);
    }
  }
  
  if(hasil.length > 0) {
    sheet.getRange(2, 2, hasil.length, 3).setValues(hasil); 
    ui.alert("Selesai", "Data Category Summary berhasil ditarik dan diupdate dalam Sheet!", ui.ButtonSet.OK);
  }
}

function fetchCategoryLocation() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getTargetSheet(SHEET_CONFIG.Location);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Info", "Penarikan dibatalkan.\nPastikan Anda sudah mengisi 'Category ID' di Kolom A.", ui.ButtonSet.OK);
    return;
  }
  
  var catIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var hasil = [];

  for (var i = 0; i < catIds.length; i++) {
    var id = catIds[i][0];
    if (!id || id.toString().trim() === "") { 
      hasil.push(["", ""]); 
      continue; 
    }
    
    var res = panggilAPI("/categorylocation?categoryId=" + encodeURIComponent(id));
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
      hasil.push(["(Data Kosong)", 0]);
    }
  }
  
  if(hasil.length > 0) {
    sheet.getRange(2, 2, hasil.length, 2).setValues(hasil); 
    ui.alert("Selesai", "Data Category Location berhasil ditarik dan diupdate dalam Sheet!", ui.ButtonSet.OK);
  }
}

