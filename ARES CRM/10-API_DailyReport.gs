/**
 * File: 10-API_DailyReport.gs
 * Handles Daily Report data fetching and email/excel exports.
 */

function getDailyReportData(filterLocation, startDate, endDate) {
  try {
      const extSS = SpreadsheetApp.openById(CONFIG_CRM.PROFILING_SS_ID);
      const trfSheet = extSS.getSheetByName(CONFIG_CRM.T_SHEET_NAME);
      if (!trfSheet) throw new Error("Traffic Sheet not found.");
      
      const data = trfSheet.getDataRange().getValues();
      const headers = data[0] || [];
      
      // Target Columns - only include essential ones to reduce processing
      const targetCols = [
          "Tanggal Berkunjung", "Rentang Waktu", "Nama Lengkap", "Customer Advisor", 
          "Lokasi Store", "Status Kedatangan", "No HP", "Net Sales"
      ];
      
      const colIndices = {};
      
      for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).trim().toLowerCase();
          
          if (h.includes('tanggal') || h.includes('tgl')) colIndices["Tanggal Berkunjung"] = j;
          else if (h.includes('rentang waktu')) colIndices["Rentang Waktu"] = j;
          else if (h === 'nama lengkap' || h.includes('nama pelanggan')) colIndices["Nama Lengkap"] = j;
          else if (h.includes('customer advisor') || h.includes('advisor')) colIndices["Customer Advisor"] = j;
          else if (h.includes('lokasi')) colIndices["Lokasi Store"] = j;
          else if (h.includes('status') && h.includes('kedatangan')) colIndices["Status Kedatangan"] = j;
          else if (h.includes('no hp') || h.includes('phone') || h.includes('handphone')) colIndices["No HP"] = j;
          else if (h.includes('net sales')) colIndices["Net Sales"] = j;
      }
      
      if (tglIdx === undefined) throw new Error("Column 'Tanggal Berkunjung' tidak ditemukan.");
      if (locIdx === undefined) throw new Error("Column 'Lokasi Store' tidak ditemukan.");
      if (stIdx === undefined) throw new Error("Column 'Status Kedatangan' tidak ditemukan.");
      if (advIdx === undefined) throw new Error("Column 'Customer Advisor' tidak ditemukan.");
      
      const parseDateValue = raw => {
          if (!raw && raw !== 0) return null;
          if (raw instanceof Date) return raw;
          if (typeof raw === 'number') return new Date(raw);
          const text = String(raw).trim();
          const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
          if (isoMatch) return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
          const dmyMatch = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
          if (dmyMatch) return new Date(parseInt(dmyMatch[3], 10), parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10));
          const parsed = new Date(text);
          return isNaN(parsed.getTime()) ? null : parsed;
      };
      
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
      
      if (tglIdx === undefined) throw new Error("Column 'Tanggal Berkunjung' tidak ditemukan.");
      if (locIdx === undefined) throw new Error("Column 'Lokasi Store' tidak ditemukan.");
      if (stIdx === undefined) throw new Error("Column 'Status Kedatangan' tidak ditemukan.");
      if (advIdx === undefined) throw new Error("Column 'Customer Advisor' tidak ditemukan.");
      
      // Parse start and end dates
      const start = new Date(startDate);
      const end = new Date(endDate);
      start.setHours(0, 0, 0, 0);
      end.setHours(23, 59, 59, 999);
      
      for (let i = 1; i < data.length; i++) {
          const row = data[i];
          
          // 1. Fast Filter Check
          let match = true;
          
          if (filterLocation && filterLocation !== 'All') {
              let rowLoc = locIdx !== undefined ? String(row[locIdx] || '').trim().toLowerCase() : '';
              if (rowLoc !== filterLocation.toLowerCase()) {
                  match = false;
              }
          }
          if (!match) continue;
          
          let rawTgl = tglIdx !== undefined ? row[tglIdx] : '';
          const rowDate = parseDateValue(rawTgl);
          
          if (!rowDate || isNaN(rowDate.getTime())) continue;
          
          // Check if date is within range
          if (rowDate < start || rowDate > end) continue;
          
          // 2. Build rowData only for matched rows
          let rowData = {};
          let isEmptyRow = true;
          targetCols.forEach(col => {
              let val = colIndices[col] !== undefined ? row[colIndices[col]] : '';
              if (val !== '' && val !== null && val !== undefined) isEmptyRow = false;
              
              if (val instanceof Date) {
                  let yy = val.getFullYear();
                  let mm = String(val.getMonth() + 1).padStart(2, '0');
                  let dd = String(val.getDate()).padStart(2, '0');
                  if (col === "Tanggal Berkunjung") {
                      val = `${yy}-${mm}-${dd}`;
                  } else {
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
          
          // 3. Increment Summaries
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
          totalHandling++;
      }
      
      out.sort((a, b) => new Date(b["Tanggal Berkunjung"]) - new Date(a["Tanggal Berkunjung"]));
      
      let totalSalesTraffic = trafficCounts["Walk In"] + trafficCounts["Follow Up"] + trafficCounts["Delivery & Showing"] + trafficCounts["Online Only"];
      let totalRepairTraffic = trafficCounts["Repair Order"] + trafficCounts["Repair Cancel"] + trafficCounts["Repair Finish"];
      
      trafficCounts["Total Traffic"] = totalSalesTraffic + totalRepairTraffic + trafficCounts["Lainnya"];
      trafficCounts["Sales Traffic"] = totalSalesTraffic;
      trafficCounts["Repair Traffic"] = totalRepairTraffic;
      trafficCounts["Total Customer"] = out.length; // usually matches out.length
      
      // Formatting Advisor Summary
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
      // Sort advisors by total descending
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
      return { success: false, message: e.message };
  }
}

function sendDailyReportEmail(filterLocation, startDate, endDate, emailTo) {
    try {
        const result = getDailyReportData(filterLocation, startDate, endDate);
        if(!result.success) throw new Error(result.message);
        
        const data = result.data;
        if(data.length === 0) throw new Error("No data found for the selected filters.");
        
        const columns = result.columns;
        
        let htmlTable = '<table border="1" style="border-collapse: collapse; font-family: sans-serif; font-size: 11px;">';
        htmlTable += '<thead style="background-color: #f2f2f2;"><tr>';
        columns.forEach(col => {
            htmlTable += `<th style="padding: 6px;">${col}</th>`;
        });
        htmlTable += '</tr></thead><tbody>';
        
        data.forEach(row => {
            htmlTable += '<tr>';
            columns.forEach(col => {
                htmlTable += `<td style="padding: 6px;">${row[col] || '-'}</td>`;
            });
            htmlTable += '</tr>';
        });
        htmlTable += '</tbody></table>';
        
        const subject = `Daily Report - Location: ${filterLocation || 'All'} [${startDate} to ${endDate}]`;
        const body = `Please find the requested daily report below.<br><br>${htmlTable}`;
        
        MailApp.sendEmail({
            to: emailTo,
            subject: subject,
            htmlBody: body
        });
        
        return { success: true, message: "Email sent successfully!" };
    } catch(e) {
        return { success: false, message: e.message };
    }
}
