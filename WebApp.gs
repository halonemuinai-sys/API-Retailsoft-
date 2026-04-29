function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Bvlgari Dashboard V2')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    // Fallback for simple HTML files if template evaluation fails
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
}

/**
 * Fetch Daily Report Data for the UI
 */
function getDailyReportData(filterLocation, filterMonth, filterYear, filterQuarter) {
  try {
      const extSS = SpreadsheetApp.openById(CONFIG_CRM.PROFILING_SS_ID);
      const trfSheet = extSS.getSheetByName(CONFIG_CRM.T_SHEET_NAME);
      if (!trfSheet) throw new Error("Traffic Sheet not found.");
      
      const data = trfSheet.getDataRange().getValues();
      const headers = data[0] || [];

      // Determine Month Range based on Quarter
      let startMonth = -1, endMonth = -1;
      if (filterQuarter === "Q1") { startMonth = 1; endMonth = 3; }
      else if (filterQuarter === "Q2") { startMonth = 4; endMonth = 6; }
      else if (filterQuarter === "Q3") { startMonth = 7; endMonth = 9; }
      else if (filterQuarter === "Q4") { startMonth = 10; endMonth = 12; }
      
      // Target Columns
      const targetCols = [
          "Tanggal Berkunjung", "Rentang Waktu", "Nama Lengkap", "Nama Panggilan", 
          "Customer Advisor", "Served By", "Lokasi Store", "Status Kedatangan", 
          "No HP", "Email", "Etnis", "Status Pelanggan", "Prospek Level", 
          "Domisili", "Domisili Luar Negeri", "Kategori Barang", 
          "Gross Sales (Retail Price)", "Penawaran Discount", "Discount (RP)", 
          "Net Sales", "Detail Items", "Descriptions", "Notes"
      ];
      
      const colIndices = {};
      for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).trim().toLowerCase();
          if (h.includes('tanggal berkunjung') || h === 'tanggal') colIndices["Tanggal Berkunjung"] = j;
          else if (h.includes('rentang waktu')) colIndices["Rentang Waktu"] = j;
          else if (h === 'nama lengkap' || h.includes('nama pelanggan')) colIndices["Nama Lengkap"] = j;
          else if (h === 'nama panggilan') colIndices["Nama Panggilan"] = j;
          else if (h.includes('customer advisor')) colIndices["Customer Advisor"] = j;
          else if (h.includes('served by')) colIndices["Served By"] = j;
          else if (h.includes('lokasi store') || h === 'lokasi') colIndices["Lokasi Store"] = j;
          else if (h.includes('status kedatangan')) colIndices["Status Kedatangan"] = j;
          else if (h.includes('no hp') || h.includes('phone') || h.includes('handphone')) colIndices["No HP"] = j;
          else if (h === 'email') colIndices["Email"] = j;
          else if (h === 'etnis') colIndices["Etnis"] = j;
          else if (h.includes('status pelanggan')) colIndices["Status Pelanggan"] = j;
          else if (h.includes('prospek level')) colIndices["Prospek Level"] = j;
          else if (h === 'kota' || h.includes('domisili')) colIndices["Domisili"] = j;
          else if (h.includes('kewarganegaraan') || h.includes('luar neg')) colIndices["Domisili Luar Negeri"] = j;
          else if (h.includes('minat barang') || h.includes('kategori')) colIndices["Kategori Barang"] = j;
          else if (h.includes('gross sales') || h.includes('retail price')) colIndices["Gross Sales (Retail Price)"] = j;
          else if (h.includes('penawaran discount')) colIndices["Penawaran Discount"] = j;
          else if (h === 'discount (rp)' || h.includes('nilai discount')) colIndices["Discount (RP)"] = j;
          else if (h.includes('net sales')) colIndices["Net Sales"] = j;
          else if (h.includes('detail item')) colIndices["Detail Items"] = j;
          else if (h === 'descriptions' || h === 'deskripsi') colIndices["Descriptions"] = j;
          else if (h === 'notes' || h === 'catatan') colIndices["Notes"] = j;
      }
      
      colIndices["Discount (RP)"] = 34; // Column AI mapping fix
      colIndices["Descriptions"] = 40; // Column AO mapping fix
      
      let out = [];
      let trafficCounts = {
          "Walk In": 0, "Follow Up": 0, "Delivery & Showing": 0, "Online Only": 0,
          "Repair Order": 0, "Repair Cancel": 0, "Repair Finish": 0, "Lainnya": 0
      };
      let advisorCounts = {};
      let totalHandling = 0;
      
      let locIdx = colIndices["Lokasi Store"];
      let tglIdx = colIndices["Tanggal Berkunjung"];
      let stIdx = colIndices["Status Kedatangan"];
      let advIdx = colIndices["Customer Advisor"];
      
      for (let i = 1; i < data.length; i++) {
          const row = data[i];
          let match = true;
          
          if (filterLocation && filterLocation !== 'All') {
              let rowLoc = locIdx !== undefined ? String(row[locIdx] || '').trim().toLowerCase() : '';
              if (rowLoc !== filterLocation.toLowerCase()) match = false;
          }
          if (!match) continue;
          
          let rawTgl = tglIdx !== undefined ? row[tglIdx] : '';
          let m = -1, y = -1;
          if (rawTgl) {
             if (rawTgl instanceof Date) {
                 m = rawTgl.getMonth() + 1;
                 y = rawTgl.getFullYear();
             } else {
                 let d = new Date(rawTgl);
                 if (!isNaN(d.getTime())) {
                     m = d.getMonth() + 1;
                     y = d.getFullYear();
                 }
             }
          }
          
          // Year Filter
          if (filterYear && filterYear !== 'All') {
              if (y === -1 || y.toString() !== filterYear.toString()) match = false;
          }
          if (!match) continue;

          // Date Range Filter (Quarter vs Month)
          if (startMonth !== -1) {
              if (m < startMonth || m > endMonth) match = false;
          } else if (filterMonth && filterMonth !== 'All') {
              if (m === -1 || m.toString() !== filterMonth.toString()) match = false;
          }
          if (!match) continue;
          
          let rowData = {};
          let isEmptyRow = true;
          targetCols.forEach(col => {
              let val = colIndices[col] !== undefined ? row[colIndices[col]] : '';
              if (val !== '' && val !== null && val !== undefined) isEmptyRow = false;
              if (val instanceof Date) {
                  let yy = val.getFullYear();
                  let mm = String(val.getMonth() + 1).padStart(2, '0');
                  let dd = String(val.getDate()).padStart(2, '0');
                  if (col === "Tanggal Berkunjung") val = `${yy}-${mm}-${dd}`;
                  else {
                      let hh = String(val.getHours()).padStart(2, '0');
                      let min = String(val.getMinutes()).padStart(2, '0');
                      let ss = String(val.getSeconds()).padStart(2, '0');
                      val = `${yy}-${mm}-${dd} ${hh}:${min}:${ss}`;
                  }
              }
              rowData[col] = String(val !== undefined && val !== null ? val : '');
          });
          
          if (isEmptyRow) continue;
          out.push(rowData);
          
          let st = stIdx !== undefined ? String(row[stIdx] || '').trim().toLowerCase() : '';
          if (st.includes('walk')) trafficCounts["Walk In"]++;
          else if (st.includes('follow')) trafficCounts["Follow Up"]++;
          else if (st.includes('delivery')) trafficCounts["Delivery & Showing"]++;
          else if (st.includes('online')) trafficCounts["Online Only"]++;
          else if (st.includes('repair order')) trafficCounts["Repair Order"]++;
          else if (st.includes('cancel') || st.includes('batal')) trafficCounts["Repair Cancel"]++;
          else if (st.includes('finish') || st.includes('selesai')) trafficCounts["Repair Finish"]++;
          else trafficCounts["Lainnya"]++;
          
          let adv = advIdx !== undefined ? String(row[advIdx] || '-').trim() : '-';
          if (!advisorCounts[adv]) advisorCounts[adv] = 0;
          advisorCounts[adv]++;
      }
      
      out.sort((a, b) => new Date(b["Tanggal Berkunjung"]) - new Date(a["Tanggal Berkunjung"]));
      
      let totalSalesTraffic = trafficCounts["Walk In"] + trafficCounts["Follow Up"] + trafficCounts["Delivery & Showing"] + trafficCounts["Online Only"];
      let totalRepairTraffic = trafficCounts["Repair Order"] + trafficCounts["Repair Cancel"] + trafficCounts["Repair Finish"];
      
      trafficCounts["Total Traffic"] = totalSalesTraffic + totalRepairTraffic + trafficCounts["Lainnya"];
      trafficCounts["Sales Traffic"] = totalSalesTraffic;
      trafficCounts["Repair Traffic"] = totalRepairTraffic;
      trafficCounts["Total Customer"] = out.length;
      
      let advSummary = [];
      totalHandling = out.length;
      Object.keys(advisorCounts).forEach(key => {
          let pct = totalHandling > 0 ? ((advisorCounts[key] / totalHandling) * 100).toFixed(2) : 0;
          advSummary.push({
              advisor: key,
              total: advisorCounts[key],
              percentage: pct + "%"
          });
      });
      advSummary.sort((a,b) => b.total - a.total);
      
      return { 
          success: true, 
          data: out, 
          columns: targetCols,
          summary: {
              traffic: trafficCounts,
              advisors: advSummary,
              totalHandling: totalHandling
          }
      };
  } catch (e) {
      return { success: false, message: e.toString() };
  }
}

/**
 * Fungsi BARU: Sinkronisasi Data dari Dashboard
 * Menarik data dari API, menyimpannya ke Sheet, lalu mengembalikan data Dashboard
 */
function syncDataFromDashboard(startDate, endDate, year) {
  try {
    // 1. Tarik data dari API ke Sheet
    fetchDailySales(startDate, endDate);
    
    // 2. Build ulang index summary (YTD)
    buildMonthlyIndexSilent();
    
    // 3. Ambil data Dashboard (Harian & Annual)
    var monthlyData = getDashboardDailySales(startDate, endDate);
    var annualData = getAnnualPerformance(year);
    
    return {
      status: "success",
      monthly: monthlyData,
      annual: annualData
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * Dipanggil dari JavaScript UI untuk menarik data Daily Sales
 */
function getDashboardDailySales(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CONFIG.Sales);
    if (!sheet) return { status: "error", message: "Sheet Daily Sales tidak ditemukan." };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { status: "success", kpi: {}, data: { sales: [] } };

    var rows = data.slice(1);

    var salesData = [], dpsData = [];
    var totalSalesValue = 0, totalQty = 0, totalTransactions = 0, totalDpsBarangJualSales = 0;
    var salesByLocation = {}, uniqueTx = {}, dailyTrend = {};

    // Parsing tanggal input
    var start = new Date(startDate);
    var end = new Date(endDate);
    start.setHours(0,0,0,0);
    end.setHours(23,59,59,999);

    var prevStart = new Date(start);
    prevStart.setMonth(prevStart.getMonth() - 1);
    var prevEnd = new Date(start);
    prevEnd.setDate(0); // hari terakhir bulan sebelumnya
    prevEnd.setHours(23,59,59,999);

    var prevSalesValue = 0, prevQty = 0, prevTransactions = 0, prevDpsBarangJualSales = 0, prevDpsServiceCenterSales = 0, prevDpsRev = 0;
    var totalDpsServiceCenterSales = 0, totalDpsRev = 0;
    var prevUniqueTx = {}, prevLocationSales = {};

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rowDateStr = row[0]; // Kolom A: Transaction Date
      var rowDate = new Date(rowDateStr);
      
      if (isNaN(rowDate.getTime())) continue;
      
      var isCurrentMonth = (rowDate >= start && rowDate <= end);
      var isPrevMonth = (rowDate >= prevStart && rowDate <= prevEnd);
      
      if (isCurrentMonth || isPrevMonth) {
        // Cek dan format jam dengan aman
        var timeStr = "";
        if (isCurrentMonth) {
          if (row[1] instanceof Date) {
            timeStr = Utilities.formatDate(row[1], ss.getSpreadsheetTimeZone(), "HH:mm:ss");
          } else {
            timeStr = row[1] ? String(row[1]) : "";
          }
        }

        var item = {
          transactionDate: Utilities.formatDate(rowDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"),
          transactionTime: timeStr,
          salesman: row[2] ? String(row[2]) : "",
          customerName: row[3] ? String(row[3]) : "",
          phoneNo: row[4] ? String(row[4]) : "",
          transactionNo: row[5] ? String(row[5]) : "",
          location: row[6] ? String(row[6]) : "",
          sapCode: row[7] ? String(row[7]) : "",
          caseNo: row[8] ? String(row[8]) : "",
          catalogueCode: row[9] ? String(row[9]) : "",
          description: row[10] ? String(row[10]) : "",
          collectionCode: row[11] ? String(row[11]) : "",
          qty: parseInt(row[12]) || 0,
          price: parseFloat(row[13]) || 0,
          discount: row[14] ? String(row[14]) : "",
          subTotalDiscount: parseFloat(row[15]) || 0,
          tax: parseFloat(row[16]) || 0,
          subTotaltax: parseFloat(row[17]) || 0
        };

        var isRB = item.location && item.location.indexOf("RB") === 0;
        var collCode = (item.collectionCode || "").toUpperCase();
        
        var qtyRaw = parseInt(item.qty);
        var qtyForCalc = (isNaN(qtyRaw) || qtyRaw === 0) ? 1 : qtyRaw;
        var unitPrice = parseFloat(item.price) || 0;
        var gross = qtyForCalc * unitPrice;
        var discount = parseFloat(item.subTotalDiscount) || 0;
        var taxAmount = parseFloat(item.subTotaltax || item.subTotalTax) || 0;

        var netSales = gross - discount - taxAmount;

        if (collCode.indexOf("PACK") !== -1 || collCode.indexOf("SVC") !== -1) {
          continue;
        } else if (collCode.indexOf("DPS") !== -1) {
          if (isCurrentMonth) {
            if (isRB) {
              item.dpsType = "Service Center";
              totalDpsServiceCenterSales += netSales;
            } else {
              item.dpsType = "Barang Jual";
              totalDpsBarangJualSales += netSales;
            }
            totalDpsRev += netSales;
            dpsData.push(item);
          } else if (isPrevMonth) {
            if (isRB) {
              prevDpsServiceCenterSales += netSales;
            } else {
              prevDpsBarangJualSales += netSales;
            }
            prevDpsRev += netSales;
          }
        } else {
          if (isRB) continue;

          if (isCurrentMonth) {
            salesData.push(item);
            totalSalesValue += netSales;
            totalQty += (isNaN(qtyRaw) ? 0 : qtyRaw);
            
            var loc = item.location || "Unknown";
            if (!salesByLocation[loc]) salesByLocation[loc] = 0;
            salesByLocation[loc] += netSales;
            
            if (item.transactionNo) uniqueTx[item.transactionNo] = true;
            
            var tDate = item.transactionDate;
            if (!dailyTrend[tDate]) dailyTrend[tDate] = 0;
            dailyTrend[tDate] += netSales;
          } else if (isPrevMonth) {
            prevSalesValue += netSales;
            prevQty += (isNaN(qtyRaw) ? 0 : qtyRaw);
            if (item.transactionNo) prevUniqueTx[item.transactionNo] = true;
            
            var loc = item.location || "Unknown";
            if (!prevLocationSales[loc]) prevLocationSales[loc] = 0;
            prevLocationSales[loc] += netSales;
          }
        }
      }
    }

    totalTransactions = Object.keys(uniqueTx).length;
    prevTransactions = Object.keys(prevUniqueTx).length;

    return {
      status: "success",
      kpi: {
        sales: totalSalesValue, qty: totalQty, transactions: totalTransactions,
        dpsBarangJualSales: totalDpsBarangJualSales,
        dpsServiceCenterSales: totalDpsServiceCenterSales,
        dpsRev: totalDpsRev,
        locationSales: salesByLocation, dailyTrend: dailyTrend
      },
      prevKpi: {
        sales: prevSalesValue, qty: prevQty, transactions: prevTransactions,
        dpsBarangJualSales: prevDpsBarangJualSales,
        dpsServiceCenterSales: prevDpsServiceCenterSales,
        dpsRev: prevDpsRev,
        locationSales: prevLocationSales
      },
      data: { sales: salesData, dps: dpsData }
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * Dipanggil dari JavaScript UI untuk menarik ringkasan tahunan dan YTD
 */
function getAnnualPerformance(year) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CONFIG.MonthlySummary);
    if (!sheet) return { status: "error", message: "Sheet Monthly Summary tidak ditemukan. Silakan tarik data harian/tahunan terlebih dahulu." };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { status: "success", ytd: {}, annualChart: {} };

    var rows = data.slice(1);
    var ytdSales = 0, ytdQty = 0, ytdTx = 0;
    var prevYtdSales = 0, prevYtdQty = 0, prevYtdTx = 0;
    var monthlySales = {}; // Format: { "01": { "Location1": 1000, "Location2": 2000 } }
    
    // Inisialisasi 12 bulan
    for (var m = 1; m <= 12; m++) {
      var mStr = ("0" + m).slice(-2);
      monthlySales[mStr] = {};
    }

    var locationSet = {}; // Untuk menyimpan daftar semua lokasi unik

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var period = row[0]; // YYYY-MM
      if (!period) continue;
      var rowYear, rowMonth;
      
      if (period instanceof Date) {
        rowYear = period.getFullYear();
        rowMonth = ("0" + (period.getMonth() + 1)).slice(-2);
      } else {
        var parts = String(period).split("-");
        if (parts.length !== 2) continue;
        rowYear = parts[0];
        rowMonth = parts[1];
      }
      
      var loc = row[1];
      var net = parseFloat(row[2]) || 0;
      var qty = parseInt(row[3]) || 0;
      var tx = parseInt(row[4]) || 0;

      // Jika tahun cocok, ambil untuk YTD dan Chart
      if (rowYear == year) {
        ytdSales += net;
        ytdQty += qty;
        ytdTx += tx;
        
        locationSet[loc] = true;
        
        if (!monthlySales[rowMonth][loc]) monthlySales[rowMonth][loc] = 0;
        monthlySales[rowMonth][loc] += net;
      } else if (rowYear == (year - 1)) {
        prevYtdSales += net;
        prevYtdQty += qty;
        prevYtdTx += tx;
      }
    }

    return {
      status: "success",
      ytd: {
        revenue: ytdSales,
        qty: ytdQty,
        tx: ytdTx
      },
      prevYtd: {
        revenue: prevYtdSales,
        qty: prevYtdQty,
        tx: prevYtdTx
      },
      annualChart: {
        months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        locations: Object.keys(locationSet),
        data: monthlySales
      }
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
