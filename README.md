# Studio2 v9 — Vercel-only Odoo XLSX Studio

Studio2 v9 adalah rebuild bersih untuk kembali ke Vercel tanpa Render, Cloudflare Engine, atau PC lokal aktif.

Prinsip v9:

- UI + editor XLSX berjalan di browser.
- API Vercel hanya menjadi jembatan pendek ke Odoo.
- Import dilakukan per batch kecil, bukan satu file besar.
- Export dimulai dari scan record, pilih record, pilih field, baru export.
- Foto besar, chatter panjang, dan HTML berat tidak diexport/import otomatis tanpa dipilih.

## Struktur

```txt
/
├─ src/app/page.tsx
├─ src/app/api/odoo/route.ts
├─ public/manifest.webmanifest
├─ package.json
├─ vercel.json
└─ README.md
```

## Deploy ke Vercel

1. Buat repo GitHub baru atau gunakan repo clean slate.
2. Copy semua isi folder ini ke root repo.
3. Push ke GitHub.
4. Di Vercel, pilih **New Project**.
5. Import repo.
6. Framework preset: **Next.js**.
7. Root directory: kosong/default root repo.
8. Build command: default `next build`.
9. Deploy.

## Setelah deploy

Buka production domain Vercel, bukan preview deployment kalau preview kamu masih kena Vercel Authentication.

Di Studio2, klik **⚙ Koneksi** lalu isi:

- Odoo URL: `https://nama-odoo.odoo.com`
- Database
- Username/email
- Password atau API key

Data koneksi disimpan di browser localStorage, bukan di GitHub.

## Mode Import

1. Upload XLSX.
2. Pilih sheet.
3. Cek editor sesuai model:
   - `res.partner` → Contact Editor
   - `product.template` / `product.product` → Product Editor
   - `project.*` → Project Editor
   - `knowledge.article` → Knowledge Editor
   - model lain → Dynamic Odoo Editor
4. Klik **Schema** untuk validasi field terhadap Odoo.
5. Import per batch kecil. Default 20 row.

## Mode Export

1. Isi model, contoh `res.partner`.
2. Isi fields yang ingin diexport, pisahkan koma.
3. Klik **Scan**.
4. Pilih record yang ingin diexport.
5. Klik **Export Record Terpilih**.
6. Hasil masuk ke editor XLSX.
7. Download XLSX jika sudah rapi.

## Catatan batas Vercel

Vercel bukan worker panjang. Studio2 v9 sengaja tidak membuat job berat di server. Kalau ingin export/import sangat besar:

- Kurangi field.
- Scan per halaman.
- Pilih record tertentu.
- Import batch 10–30 row.
- Hindari image base64 dan HTML/chatter panjang.

## Keamanan

Jangan commit credential Odoo ke repo.

Tidak ada `.env` wajib untuk versi ini. Semua koneksi diisi dari UI dan disimpan di browser kamu.

Kalau app ini untuk internal, gunakan akun Vercel pribadi dan jangan sebar URL Studio2 sembarangan. Untuk protection tambahan tanpa mengunci PWA/public domain, bisa dibuat password internal di versi berikutnya.


## v9.2 Command UI

- Export model sekarang memakai checklist preset, bukan kolom teks utama.
- Record picker memakai command card checklist agar tahap pemilihan tidak terasa seperti tabel mentah.
- Advanced field/domain tetap ada, tetapi disembunyikan di drawer supaya workspace tidak cluttered.
- Arah desain: professional admin command studio, bukan spreadsheet viewer.
