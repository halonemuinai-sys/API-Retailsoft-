@echo off
cd /d "d:\Project Ares\API"
echo ===========================================
echo   PROSES PUSH KE GITHUB - API RETAILSOFT
echo ===========================================

echo.
echo [1/5] Inisialisasi Git...
git init

echo.
echo [2/5] Menghubungkan ke GitHub...
git remote add origin https://github.com/halonemuinai-sys/API-Retailsoft-.git
git remote set-url origin https://github.com/halonemuinai-sys/API-Retailsoft-.git

echo.
echo [3/5] Mengatur Branch Utama ke 'main'...
git branch -M main

echo.
echo [4/5] Menambahkan File dan Membuat Commit...
git add .
git commit -m "Update Dashboard, Tata Letak Tabel, dan Filter Kuartal"

echo.
echo [5/5] Mengunggah (Push) ke GitHub...
git push -u origin main

echo.
echo ===========================================
echo   SELESAI!
echo ===========================================
pause
