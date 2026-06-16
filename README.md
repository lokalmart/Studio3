# Studio2 v9.3 — Mobile Command Studio for Odoo XLSX

Studio2 v9.3 adalah rebuild UI mobile-first untuk kembali ke Vercel tanpa Render, Cloudflare Engine, atau PC lokal aktif.

Arah v9.3:

- Bukan spreadsheet viewer mentah.
- Bukan dashboard penuh card yang cluttered.
- UI dibuat seperti mobile command app: satu workflow, satu aksi utama, panel ringkas, bottom navigation, dan editor card di HP.
- API Vercel tetap hanya menjadi jembatan pendek ke Odoo.
- Import dilakukan per batch kecil agar aman untuk batas serverless.

## Struktur

```txt
/
├─ src/app/page.tsx
├─ src/app/api/odoo/route.ts
├─ src/app/globals.css
├─ public/manifest.webmanifest
├─ package.json
├─ vercel.json
└─ README.md
```

## Deploy ke Vercel

1. Copy isi ZIP ke root repo GitHub.
2. Pastikan `package.json` ada langsung di root repo.
3. Di Vercel pilih **New Project**.
4. Framework preset: **Next.js**.
5. Root directory: kosong/default.
6. Deploy.

## Setelah deploy

Buka production domain Vercel. Di Studio2 tekan **⚙ Koneksi**, lalu isi:

- Odoo URL
- Database
- Username/email
- Password atau API key

Credential disimpan di browser `localStorage`, bukan di GitHub.

## Perubahan penting v9.3

- Layout mobile-first, bukan desktop-first.
- Bottom navigation untuk Import, Export, Koneksi.
- Import dimulai dari hero upload yang jelas.
- Setelah XLSX dibuka, sheet tampil sebagai horizontal picker.
- Action utama dibuat sticky: Import / Export selalu jelas.
- Editor mobile memakai row cards, bukan tabel kecil yang tidak terbaca.
- Desktop tetap punya spreadsheet grid untuk edit massal.
- Export model memakai checklist mission cards.
- Record picker tetap card checklist.
- Log dibuat drawer bawah, bukan panel samping yang memakan ruang.

## Mode Import

1. Pilih XLSX.
2. Pilih sheet.
3. Load schema jika ingin validasi field Odoo.
4. Edit row sebagai card di mobile atau grid di desktop.
5. Import batch kecil.

## Mode Export

1. Pilih model dari mission cards.
2. Scan record.
3. Pilih record sebagai checklist card.
4. Export record terpilih.
5. Hasil export masuk ke editor XLSX.
6. Download setelah diedit.

## Catatan batas Vercel

Agar aman di Vercel free/hobby:

- Import 10–30 row per batch.
- Export record terpilih, bukan seluruh database.
- Pilih field seperlunya.
- Hindari image base64, chatter, dan HTML panjang kecuali perlu.

