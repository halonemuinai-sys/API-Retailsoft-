require('dotenv').config();
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const path = require('path');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Konfigurasi dari 0-Config.gs
const BASE_URL = process.env.BASE_URL || "http://139.99.102.231:8189/api";
const API_TOKEN = process.env.API_TOKEN || "Bearer B0KiIGq0Q7LP/Sg+mOuQNdEH6Xogt4Kf4W8sKhQJiMA6ItgTswhTtg8Mx2/Bzq3T";

/**
 * Fungsi untuk meniru HtmlService Google Apps Script
 * Mencari file di folder ROOT (Outer Folder)
 */
function renderHtml(fileName) {
  const rootPath = path.resolve(__dirname); // Pindah ke outer folder
  const filePath = path.join(rootPath, `${fileName}.html`);
  
  if (!fs.existsSync(filePath)) {
    // Jika tidak ketemu di root, coba cari di subfolder ARES CRM (sebagai fallback sample)
    const samplePath = path.join(rootPath, 'ARES CRM', `${fileName}.html`);
    if (fs.existsSync(samplePath)) return fs.readFileSync(samplePath, 'utf8').replace(/<\?!= include\(['"](.+?)['"]\); \?>/g, (m, name) => renderHtml(name));
    
    return `<!-- File ${fileName}.html not found -->`;
  }

  let content = fs.readFileSync(filePath, 'utf8');
  
  // Mendukung <?!= include('FileName'); ?> sesuai WebApp.gs
  const includeRegex = /<\?!= include\(['"](.+?)['"]\); \?>/g;
  content = content.replace(includeRegex, (match, includeName) => {
    return renderHtml(includeName);
  });

  return content;
}

// Route Utama - Mengikuti function doGet() di WebApp.gs
app.get('/', (req, res) => {
  res.send(renderHtml('Dashboard'));
});

// Dynamic Route untuk View lainnya
app.get('/:page', (req, res) => {
  res.send(renderHtml(req.params.page));
});

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://vekgzcxorvdidjutuvrj.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZla2d6Y3hvcnZkaWRqdXR1dnJqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQyOTI2NzIsImV4cCI6MjA4OTg2ODY3Mn0.Kz9udMSBq9YbyFsCmQvAWYPjNhplFsNKcjtiDdIi04I';

/**
 * Handler utama untuk semua panggilan dari google.script.run (via VercelShim)
 */
app.post('/api/call', async (req, res) => {
  const { functionName, args } = req.body;
  console.log(`[Vercel Call] ${functionName}`, args);

  try {
    // 1. Logika Sinkronisasi: Sync Data Sales (GAS Style De-duplication)
    if (functionName === 'syncDataFromDashboard' || functionName === 'triggerGlobalSupabaseSync') {
       const startDate = args[0];
       const endDate = args[1];

       console.log(`[Sync Sales] GAS-Style Fetch for ${startDate} to ${endDate}`);
       
       // Step A: Tarik data dari API Bvlgari
       const apiUrl = `http://139.99.102.231:8089/demo/dailysalestransaction?startdate=${encodeURIComponent(startDate)}&enddate=${encodeURIComponent(endDate)}`;
       const response = await axios.get(apiUrl, { headers: { 'Authorization': API_TOKEN } });
       const rawData = response.data;
       if (!Array.isArray(rawData)) throw new Error("API Bvlgari tidak mengembalikan data array.");

       // Step B: Ambil daftar Transaction No yang SUDAH ADA di Supabase (Mimic Baris 48 di GAS)
       const existingRes = await axios.get(`${SUPABASE_URL}/rest/v1/bvlgari_sales?select=transaction_no`, {
         headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` }
       });
       const existingTx = new Set((existingRes.data || []).map(r => r.transaction_no));

       // Step C: Filter & Mapping (Hanya masukkan yang BELUM ADA)
       const newRows = [];
       rawData.forEach(itm => {
           // Skip jika sudah ada (Mimic Baris 55 di GAS)
           if (existingTx.has(itm.transactionNo)) return;

           const qtyRaw = parseInt(itm.qty) || 0;
           const unitPrice = parseFloat(itm.price) || 0;
           const discountVal = parseFloat(itm.subTotalDiscount) || 0;
           let taxAmount = parseFloat(itm.subTotaltax || itm.subTotalTax) || 0;
           const taxRate = parseFloat(itm.tax) || 0;
           const collCode = (itm.collectionCode || "").toUpperCase();
           const grossAfterDiscount = (qtyRaw || 1) * unitPrice - discountVal;

           if (taxAmount === 0 && (collCode === "PFM" || taxRate > 1)) {
               const divisor = taxRate > 1 ? taxRate : 1.11;
               taxAmount = grossAfterDiscount - (grossAfterDiscount / divisor);
           }

           newRows.push({
               transaction_date: itm.transactionDate || null,
               transaction_time: itm.transactionTime || null,
               salesman: itm.salesman || null,
               customer_name: itm.customerName || null,
               phone_no: itm.phoneNo || null,
               transaction_no: itm.transactionNo || null,
               location: itm.location || null,
               sap_code: itm.sapCode || null,
               catalogue_code: itm.catalogueCode || null,
               description: itm.description || null,
               collection: itm.collectionCode || null,
               qty: qtyRaw,
               price: unitPrice,
               sub_total_discount: discountVal,
               sub_total_tax: taxAmount,
               net_sales: grossAfterDiscount - taxAmount
           });
       });

       // Step D: Insert data baru ke Supabase
       if (newRows.length > 0) {
           await axios.post(`${SUPABASE_URL}/rest/v1/bvlgari_sales`, newRows, {
             headers: {
               'apikey': SUPABASE_KEY,
               'Authorization': `Bearer ${SUPABASE_KEY}`,
               'Content-Type': 'application/json',
               'Prefer': 'return=minimal'
             }
           });
       }

       console.log(`[Sync Success] Added ${newRows.length} rows.`);

       // Step E: Tarik seluruh data bulan tersebut dari Supabase
       const monthDataRes = await axios.get(
         `${SUPABASE_URL}/rest/v1/bvlgari_sales?transaction_date=gte.${startDate}&transaction_date=lte.${endDate}&order=transaction_date.desc`,
         { headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` } }
       );
       const monthData = monthDataRes.data || [];
       let totalMonthSales = 0, totalMonthQty = 0;
       const monthTxSet = new Set();
       const locationSales = {};
       const dailyTrend = {};
       
       // Helper untuk menghitung summary per lokasi
       const locSummary = {};

       monthData.forEach(r => {
         const net = Number(r.net_sales) || 0;
         const qty = Number(r.qty) || 0;
         const loc = r.location || "Unknown";
         const dateStr = r.transaction_date;
         const txNo = r.transaction_no;

         totalMonthSales += net;
         totalMonthQty += qty;
         monthTxSet.add(txNo);

         if (!locationSales[loc]) locationSales[loc] = 0;
         locationSales[loc] += net;

         if (dateStr) {
           if (!dailyTrend[dateStr]) dailyTrend[dateStr] = 0;
           dailyTrend[dateStr] += net;
         }

         // Agregasi untuk bvlgari_monthly_summary
         if (!locSummary[loc]) locSummary[loc] = { net_sales: 0, qty: 0, txSet: new Set() };
         locSummary[loc].net_sales += net;
         locSummary[loc].qty += qty;
         locSummary[loc].txSet.add(txNo);
       });

       // Step E.2: Push hasil agregasi ke bvlgari_monthly_summary
       const periodStr = startDate.substring(0, 7); // e.g., "2026-05"
       const summaryRows = Object.keys(locSummary).map(loc => {
           return {
               period: periodStr,
               location: loc,
               total_net_sales: locSummary[loc].net_sales,
               total_qty: locSummary[loc].qty,
               total_transactions: locSummary[loc].txSet.size
           };
       });

       if (summaryRows.length > 0) {
           await axios.post(
             `${SUPABASE_URL}/rest/v1/bvlgari_monthly_summary?on_conflict=period,location`,
             summaryRows,
             {
               headers: {
                 'apikey': SUPABASE_KEY,
                 'Authorization': `Bearer ${SUPABASE_KEY}`,
                 'Content-Type': 'application/json',
                 'Prefer': 'resolution=merge-duplicates'
               }
             }
           );
           console.log(`[Sync] Updated ${summaryRows.length} locations in bvlgari_monthly_summary`);
       }

       // Step F: Tarik data tahunan untuk grafik YTD (Bisa disederhanakan pakai bvlgari_monthly_summary di tahap selanjutnya)
       const year = startDate.split('-')[0];
       const yearDataRes = await axios.get(
         `${SUPABASE_URL}/rest/v1/bvlgari_sales?transaction_date=gte.${year}-01-01&transaction_date=lte.${year}-12-31`,
         { headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` } }
       );
       const allSales = yearDataRes.data || [];
       const monthlyData = {}; 
       const locationSet = new Set();
       let totalYtdRevenue = 0, totalYtdQty = 0;
       const ytdTxSet = new Set();

       allSales.forEach(r => {
          if (!r.transaction_date) return;
          const m = r.transaction_date.split('-')[1];
          const loc = r.location || "Unknown";
          const net = Number(r.net_sales) || 0;
          const qty = Number(r.qty) || 0;
          
          locationSet.add(loc);
          totalYtdRevenue += net;
          totalYtdQty += qty;
          ytdTxSet.add(r.transaction_no);

          if (!monthlyData[m]) monthlyData[m] = {};
          if (!monthlyData[m][loc]) monthlyData[m][loc] = 0;
          monthlyData[m][loc] += net;
       });

       return res.json({ 
         status: "success", 
         message: `Berhasil menambah ${newRows.length} data transaksi baru!`,
         monthly: { 
           status: "success", 
           kpi: { sales: totalMonthSales, qty: totalMonthQty, transactions: monthTxSet.size, locationSales, dailyTrend }, 
           data: { sales: monthData } 
         },
         annual: { 
           status: "success", 
           ytd: { revenue: totalYtdRevenue, qty: totalYtdQty, tx: ytdTxSet.size, target: 10000000000, percentage: Math.round((totalYtdRevenue / 10000000000) * 100) }, 
           annualChart: { months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], locations: Array.from(locationSet), data: monthlyData } 
         }
       });
    }

    // 2. Logika Fetch Data Dashboard: Ambil dari bvlgari_sales (Real KPI & Trend)
    if (functionName === 'getDashboardDailySales') {
       const startDate = args[0];
       const endDate = args[1];

       console.log(`[Dashboard] Calculating KPI & Trend from ${startDate} to ${endDate}`);

       // Tarik data bulan ini
       const supabaseData = await axios.get(
         `${SUPABASE_URL}/rest/v1/bvlgari_sales?transaction_date=gte.${startDate}&transaction_date=lte.${endDate}&order=transaction_date.desc`,
         { headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` } }
       );
       const data = supabaseData.data || [];
       let totalSales = 0, totalQty = 0;
       const txSet = new Set();
       const locationSales = {};
       const dailyTrend = {};

       data.forEach(r => {
         const net = Number(r.net_sales) || 0;
         const qty = Number(r.qty) || 0;
         const loc = r.location || "Unknown";
         totalSales += net;
         totalQty += qty;
         txSet.add(r.transaction_no);
         if (!locationSales[loc]) locationSales[loc] = 0;
         locationSales[loc] += net;
         if (r.transaction_date) {
           if (!dailyTrend[r.transaction_date]) dailyTrend[r.transaction_date] = 0;
           dailyTrend[r.transaction_date] += net;
         }
       });

       // Hitung tanggal bulan lalu
       let [yStr, mStr] = startDate.split('-');
       let py = parseInt(yStr);
       let pm = parseInt(mStr) - 1;
       if (pm === 0) { pm = 12; py -= 1; }
       const prevStart = `${py}-${String(pm).padStart(2, '0')}-01`;
       const prevEndObj = new Date(py, pm, 0); 
       const prevEnd = `${prevEndObj.getFullYear()}-${String(prevEndObj.getMonth()+1).padStart(2, '0')}-${String(prevEndObj.getDate()).padStart(2, '0')}`;

       // Tarik data bulan lalu (untuk trend 'vs last month')
       const prevDataRes = await axios.get(
         `${SUPABASE_URL}/rest/v1/bvlgari_sales?transaction_date=gte.${prevStart}&transaction_date=lte.${prevEnd}`,
         { headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` } }
       );
       const prevData = prevDataRes.data || [];
       let prevSales = 0, prevQty = 0;
       const prevTxSet = new Set();
       const prevLocSales = {};
       
       prevData.forEach(r => {
         const net = Number(r.net_sales) || 0;
         prevSales += net;
         prevQty += Number(r.qty) || 0;
         prevTxSet.add(r.transaction_no);
         const loc = r.location || "Unknown";
         if (!prevLocSales[loc]) prevLocSales[loc] = 0;
         prevLocSales[loc] += net;
       });

       return res.json({
         status: "success",
         kpi: { sales: totalSales, qty: totalQty, transactions: txSet.size, locationSales, dailyTrend },
         prevKpi: { sales: prevSales, qty: prevQty, transactions: prevTxSet.size, locationSales: prevLocSales },
         data: { sales: data, dps: [] }
       });
    }

    // 3. Real Annual Performance (Grafik YTD & Trend vs Last Year)
    if (functionName === 'getAnnualPerformance') {
       const year = parseInt(args[0] || 2026);
       console.log(`[Annual] Calculating YTD for ${year} and prev year ${year-1}`);

       // Data Tahun Ini
       const resSales = await axios.get(
         `${SUPABASE_URL}/rest/v1/bvlgari_sales?transaction_date=gte.${year}-01-01&transaction_date=lte.${year}-12-31`,
         { headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` } }
       );
       const allSales = resSales.data || [];
       
       const monthlyData = {}; 
       const locationSet = new Set();
       let totalYtdRevenue = 0, totalYtdQty = 0;
       const ytdTxSet = new Set();

       allSales.forEach(r => {
          if (!r.transaction_date) return;
          const m = r.transaction_date.split('-')[1];
          const loc = r.location || "Unknown";
          const net = Number(r.net_sales) || 0;
          
          locationSet.add(loc);
          totalYtdRevenue += net;
          totalYtdQty += Number(r.qty) || 0;
          ytdTxSet.add(r.transaction_no);

          if (!monthlyData[m]) monthlyData[m] = {};
          if (!monthlyData[m][loc]) monthlyData[m][loc] = 0;
          monthlyData[m][loc] += net;
       });

       // Data Tahun Lalu (Untuk trend 'vs last year')
       const prevYearRes = await axios.get(
         `${SUPABASE_URL}/rest/v1/bvlgari_sales?transaction_date=gte.${year-1}-01-01&transaction_date=lte.${year-1}-12-31`,
         { headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` } }
       );
       const prevYearSales = prevYearRes.data || [];
       let prevYtdRev = 0, prevYtdQty = 0;
       const prevYtdTxSet = new Set();
       
       prevYearSales.forEach(r => {
          prevYtdRev += Number(r.net_sales) || 0;
          prevYtdQty += Number(r.qty) || 0;
          prevYtdTxSet.add(r.transaction_no);
       });

       return res.json({
         status: "success",
         ytd: { revenue: totalYtdRevenue, qty: totalYtdQty, tx: ytdTxSet.size, target: 10000000000, percentage: Math.round((totalYtdRevenue / 10000000000) * 100) },
         prevYtd: { revenue: prevYtdRev, qty: prevYtdQty, tx: prevYtdTxSet.size },
         annualChart: { 
           months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 
           locations: Array.from(locationSet), 
           data: monthlyData 
         }
       });
    }

    // 3. Logika CRM Daily Report: Ambil dari mirror_traffic
    if (functionName === 'getDailyReportData') {
       const [loc, monthStr, yearStr, quarter] = args;
       
       console.log(`[DailyReport] Fetching from mirror_traffic for ${loc}, ${monthStr}/${yearStr}`);

       // Build Query
       let query = `${SUPABASE_URL}/rest/v1/mirror_traffic?select=*`;
       
       if (loc && loc !== 'All') query += `&location=eq.${loc}`;
       // Filter tanggal di Postgres/Supabase
       if (yearStr && yearStr !== 'All') query += `&transaction_date=gte.${yearStr}-01-01&transaction_date=lte.${yearStr}-12-31`;

       const response = await axios.get(query, {
         headers: {
           'apikey': SUPABASE_KEY,
           'Authorization': `Bearer ${SUPABASE_KEY}`
         }
       });

       const data = response.data || [];
       
       // Filter tambahan untuk bulan (karena Supabase REST API filter bulan butuh format spesifik)
       let filteredData = data;
       if (monthStr && monthStr !== 'All') {
         const m = monthStr.padStart(2, '0');
         filteredData = data.filter(r => r.transaction_date && r.transaction_date.includes(`-${m}-`));
       }

       // Hitung Summary Traffic
       const trafficSummary = {
          "Walk In": 0, "Follow Up": 0, "Delivery & Showing": 0, "Online Only": 0,
          "Repair Order": 0, "Repair Cancel": 0, "Repair Finish": 0, "Lainnya": 0,
          "Total Traffic": 0, "Sales Traffic": 0, "Repair Traffic": 0, "Total Customer": 0
       };

       const advisorMap = {};
       const mappedData = [];

       filteredData.forEach(r => {
          // Mapping data untuk Frontend & Excel
          mappedData.push({
             "Date": r.transaction_date || "-",
             "Time": r.time || "-",
             "Location": r.location || "-",
             "Customer Advisor": r.salesman || "-",
             "Customer Name": r.customer_name || "-",
             "Phone No": r.phone_no || "-",
             "Status Kedatangan": r.status_kedatangan || "-",
             "Prospect Item": r.prospect_item || "-",
             "Keterangan": r.keterangan || "-"
          });

          // Traffic count
          const st = (r.status_kedatangan || "").toLowerCase();
          if (st.includes('walk')) trafficSummary["Walk In"]++;
          else if (st.includes('follow')) trafficSummary["Follow Up"]++;
          else if (st.includes('delivery')) trafficSummary["Delivery & Showing"]++;
          else if (st.includes('online')) trafficSummary["Online Only"]++;
          else if (st.includes('repair order')) trafficSummary["Repair Order"]++;
          else if (st.includes('cancel')) trafficSummary["Repair Cancel"]++;
          else if (st.includes('finish')) trafficSummary["Repair Finish"]++;
          else trafficSummary["Lainnya"]++;

          // Advisor count
          const adv = r.salesman || "Unknown";
          if (!advisorMap[adv]) advisorMap[adv] = 0;
          advisorMap[adv]++;
       });

       trafficSummary["Sales Traffic"] = trafficSummary["Walk In"] + trafficSummary["Follow Up"] + trafficSummary["Delivery & Showing"] + trafficSummary["Online Only"];
       trafficSummary["Repair Traffic"] = trafficSummary["Repair Order"] + trafficSummary["Repair Cancel"] + trafficSummary["Repair Finish"];
       trafficSummary["Total Traffic"] = trafficSummary["Sales Traffic"] + trafficSummary["Repair Traffic"] + trafficSummary["Lainnya"];
       trafficSummary["Total Customer"] = filteredData.length;

       // Calculate Advisors summary
       const totalHandlingSum = filteredData.length;
       const advisors = Object.keys(advisorMap).map(adv => {
           const total = advisorMap[adv];
           const pct = totalHandlingSum > 0 ? ((total / totalHandlingSum) * 100).toFixed(2) + "%" : "0.00%";
           return { advisor: adv, total: total, percentage: pct };
       }).sort((a, b) => b.total - a.total);

       return res.json({
         success: true,
         data: mappedData,
         columns: [
           "Date", "Time", "Location", "Customer Advisor", "Customer Name", 
           "Phone No", "Status Kedatangan", "Prospect Item", "Keterangan"
         ],
         summary: {
           traffic: trafficSummary,
           advisors: advisors 
         }
       });
    }

  } catch (error) {
    console.error("Error in /api/call:", error.response ? error.response.data : error.message);
    res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * Proxy untuk memanggil API Bvlgari
 */
app.post('/api/fetch-bvlgari', async (req, res) => {
  const { endpoint, params } = req.body;
  
  try {
    const response = await axios({
      method: 'get',
      url: `${BASE_URL}/${endpoint}`,
      headers: {
        'Authorization': API_TOKEN,
        'Content-Type': 'application/json'
      },
      params: params
    });

    res.json({
      success: true,
      data: response.data
    });
  } catch (error) {
    console.error("API Error:", error.message);
    res.status(500).json({
      success: false,
      message: error.message,
      detail: error.response ? error.response.data : null
    });
  }
});

// Implementasi fungsi-fungsi lain dari .gs akan diletakkan di sini secara bertahap

app.listen(port, () => {
  console.log(`Server bridge berjalan di port ${port}`);
});

module.exports = app;
