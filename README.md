# PO Recommendation, Deadstock & Overstock Dashboard

Streamlit app dengan 4 menu untuk K. Beauty / The Beauty Shop:

1. **🏠 Welcome** - overview cepat (total SKU, total qty stok, total brand)
   begitu Data Stok diupload.
2. **📊 PO Recommendation** - draft PO mingguan dari **Data Stok** +
   **Data Penjualan 30 Hari**.
3. **🧟 Deadstock** - cari barang yang tidak laku dalam **3 bulan**
   terakhir dari **Data Stok** + **Data Penjualan 3 Bulan**. Produk baru
   yang baru direstock (dicek lewat **Data Pembelian**, opsional) tidak
   ikut ke-flag - wajar kalau belum sempat laku.
4. **📈 Overstock** - barang dengan stok LEBIH BANYAK dari penjualan 3
   bulan (lebih luas dari Deadstock, termasuk yang sempat laku sedikit
   tapi masih kelebihan stok jauh).

Upload file aslinya apa adanya dari export iPOS 5.0, tidak perlu dirapikan
dulu. Data Stok dipakai bersama di menu 2, 3 & 4 - upload sekali saja di
sidebar. Data Penjualan 3 Bulan dipakai bersama di menu 3 & 4.

## Rumus & Rules

### PO Recommendation (menu 2)
1. **Slow moving**: kalau `Stok == Penjualan 30 Hari`, SKU ditandai 🟡
   **CHECK SALES** (termasuk kasus 0 = 0, barang mati total).
2. Baris duplikat (Product ID sama muncul lebih dari sekali di satu file)
   otomatis digabung (qty dijumlahkan) sebelum diproses.
3. Formula PO: **PO = Penjualan 30 hari − Stok**. Negatif/nol → tidak perlu PO.
4. Kalau stok sudah nol (atau diperkirakan habis <3 hari lagi berdasarkan
   kecepatan jual), SKU ditandai 🔴 **URGENT**. Sisanya yang perlu PO tapi
   tidak mendesak → 🟢 **NORMAL**.

### Deadstock (menu 3)
Barang dianggap **deadstock** kalau: `Qty Terjual 3 Bulan == 0` (termasuk
SKU yang sama sekali tidak muncul di data penjualan 3 bulan - dianggap 0)
**DAN** `Stok > 0` saat ini **DAN** SKU-nya tidak muncul di Data Pembelian
(kalau diupload) - karena kalau baru direstock, wajar belum sempat laku,
jadi bukan "mati" beneran.

### Overstock (menu 4)
Barang dianggap **overstock** kalau `Stok > Qty Terjual 3 Bulan`. Ini
mencakup semua kasus Deadstock (stok>0, terjual=0 otomatis stok>terjual)
plus kasus tambahan: barang yang sempat laku tapi cuma sedikit dibanding
stok yang ada. Beda dengan Deadstock, Overstock tidak exclude produk baru
- karena barang baru dengan stok masuk besar tapi belum laku tetap relevan
untuk dipantau di sini.

### Sales Staff (menu 3 & 4)
Hasil Deadstock dan Overstock dikelompokkan per **Sales Staff** berdasarkan
brand yang mereka pegang (lihat `STAFF_BRAND_MAP` di `po_logic.py`). Excel
export-nya jadi 1 sheet per staff (bukan per brand lagi), plus ada filter
Sales Staff & Brand di tampilan web-nya. Brand yang belum ada pemiliknya di
mapping otomatis masuk sheet **"Belum Terpetakan"**.

Kalau ada brand baru atau staff pindah brand, edit langsung dict
`STAFF_BRAND_MAP` di bagian atas `po_logic.py` - tidak perlu ubah logic lain.

## Menjalankan lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

Lalu buka link `http://localhost:8501` yang muncul di terminal, upload
file-file yang dibutuhkan di sidebar.

## Deploy gratis (biar sales/tim kamu bisa akses dari browser tanpa install apa-apa)

1. Push folder ini (`app.py`, `po_logic.py`, `requirements.txt`, folder
   `assets`, folder `.streamlit`) ke repo GitHub (boleh private).
2. Buka https://share.streamlit.io -> "New app" -> hubungkan ke repo itu.
3. Streamlit Community Cloud otomatis build & hosting gratis.

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

## Tampilan / Branding

Dashboard-nya sudah pakai logo & warna K. Beauty / The Beauty Shop (teal
`#186156` dan lime `#ADF901`, diambil langsung dari logo kamu). Kalau
suatu saat logo berubah, tinggal ganti file `assets/logo.jpg` dengan nama
file yang sama - tidak perlu ubah kode.

## File di folder ini

- `app.py` - UI Streamlit: 4 menu (Welcome, PO Recommendation, Deadstock, Overstock).
- `po_logic.py` - semua logic: parsing file, dedup, join, klasifikasi PO,
  deadstock/overstock finder, dan mapping brand ke sales staff. Dipisah
  dari `app.py` supaya bisa dites sendiri tanpa buka browser.
- `assets/logo.jpg` - logo K. Beauty / The Beauty Shop yang tampil di header & sidebar.
- `.streamlit/config.toml` - theme warna (teal & lime) yang otomatis dipakai Streamlit.
- `requirements.txt` - dependency untuk `pip install`.
- `.gitignore` - biar folder cache Python (`__pycache__`) tidak ikut ke-push ke GitHub.
