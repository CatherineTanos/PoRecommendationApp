# PO Recommendation Dashboard

Streamlit app untuk hasilkan draft PO mingguan dari 2 file export iPOS 5.0
kamu: **Daftar Penjualan per Item** (30 hari terakhir) & **Daftar Item**
(stok). Upload file aslinya apa adanya, tidak perlu dirapikan dulu.

## Rumus & Rules (per request terakhir)

1. **Slow moving**: tidak pakai angka ambang tetap. Aturannya sederhana -
   kalau `Stok == Penjualan 30 Hari`, SKU ditandai 🟡 **CHECK SALES**
   (termasuk kasus 0 = 0, yang berarti barang mati total - tidak laku &
   tidak ada stok).
2. Hanya 2 sumber data: **penjualan 30 hari** + **stok**. Tidak ada file
   master barang terpisah.
3. Kedua file digabung (join) berdasarkan **Product ID** (`Kode Item` di
   export iPOS kamu).
4. Baris duplikat (Product ID yang sama muncul lebih dari sekali di satu
   file) otomatis dideteksi & digabung (qty dijumlahkan) sebelum diproses,
   supaya tidak ada data dobel.
5. Formula PO: **PO = Penjualan 30 hari − Stok**. Kalau hasilnya negatif
   atau nol → tidak perlu PO.

Di atas formula dasar itu, ada 1 lapisan tambahan supaya kamu tidak perlu
cek semua SKU satu-satu: kalau stok sudah nol (atau sisa stok diperkirakan
habis dalam <3 hari berdasarkan kecepatan jual 30 hari terakhir), item
ditandai 🔴 **URGENT** meskipun secara matematis sama-sama "perlu PO".
Sisanya yang perlu PO tapi tidak mendesak masuk 🟢 **NORMAL**.

## Menjalankan lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

Lalu buka link `http://localhost:8501` yang muncul di terminal, upload 2
file export iPOS kamu di sidebar.

## Deploy gratis (biar sales/tim kamu bisa akses dari browser tanpa install apa-apa)

1. Push folder ini (`app.py`, `po_logic.py`, `requirements.txt`) ke repo
   GitHub (boleh private).
2. Buka https://share.streamlit.io -> "New app" -> hubungkan ke repo itu.
3. Streamlit Community Cloud otomatis build & hosting gratis. Kamu dapat
   URL publik yang bisa dibagikan (bisa dibuat butuh login lewat
   pengaturan "sharing" di Streamlit Cloud kalau mau dibatasi).

## Kalau format export iPOS kamu berubah

`po_logic.py` sudah didesain untuk baca langsung dari layout report iPOS
yang aslinya berantakan (merged cell, baris kosong, baris "Total :", dll) -
kolom dicari otomatis berdasarkan nama header (`Kode Item`, `Nama Item`,
`Jenis`, `Jumlah`/`Stok`, `Satuan`, `Total Harga`/`Harga Pokok`), bukan
posisi kolom yang di-hardcode. Jadi kalau suatu saat kolomnya sedikit
bergeser, kemungkinan besar tetap terbaca otomatis.

Kalau kamu suatu saat punya file yang sudah rapi (1 baris = 1 SKU dengan
header di baris pertama), file itu juga tetap kebaca lewat fallback
loader-nya, asal ada kolom yang mengandung kata seperti "Kode Item"/
"Product ID"/"SKU" dan kolom qty/stok yang jelas.

## File di folder ini

- `app.py` - UI Streamlit (upload, dashboard, filter, download Excel).
- `po_logic.py` - semua logic: parsing file, dedup, join, klasifikasi PO.
  Dipisah dari `app.py` supaya bisa dites sendiri tanpa buka browser.
- `requirements.txt` - dependency untuk `pip install`.
- `PO_Recommendation_Preview.xlsx` - contoh hasil dari data kamu minggu ini
  (20 Jul - 20 Agu 2026), biar kamu bisa langsung lihat hasilnya tanpa
  perlu jalankan app dulu.
