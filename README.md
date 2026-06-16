# Studio2 v10 — Lokalmart Data Command Studio

Studio2 v10 adalah rebuild UI/UX untuk Vercel-only deployment. Fokusnya bukan lagi “spreadsheet viewer”, tetapi **mission-based command studio** untuk operasi data Odoo Lokalmart.

## Prinsip desain v10

- Mobile-first, bukan desktop panel yang dipaksa mengecil.
- Home sebagai **Mission Launcher**: Import, Export, Review, Koneksi.
- Export model berupa checklist card, bukan input teks mentah.
- Spreadsheet grid hanya mode tambahan; default editor berupa object cards.
- Progressive disclosure: schema, domain, field teknis, dan log disembunyikan sampai dibutuhkan.
- Vercel-safe: browser membaca XLSX, API hanya request kecil ke Odoo.

## Fitur utama

### Import Mission

Upload XLSX → Review sheet → Cek schema → Object editor → Import batch kecil ke Odoo.

### Export Mission

Pilih model via checklist card → Scan record → Pilih record → Export ke editor → Download XLSX.

Preset model:

- Contacts (`res.partner`)
- Products (`product.template`)
- Projects (`project.project`)
- Tasks (`project.task`)
- Knowledge (`knowledge.article`)
- Sales (`sale.order`)
- Categories (`product.category`)
- Web Categories (`product.public.category`)

### Review Workspace

Default editor berupa card accordion per row. Grid spreadsheet tetap tersedia untuk edit cepat di desktop.

## Deploy ke Vercel

Pastikan struktur repo root seperti ini:

```text
/package.json
/next.config.mjs
/vercel.json
/src/app/page.tsx
/src/app/api/odoo/route.ts
/public/manifest.webmanifest
```

Di Vercel:

- Framework preset: Next.js
- Root Directory: kosong/default root repo
- Build Command: default
- Output Directory: default

## Koneksi Odoo

Buka Studio2 → Koneksi → isi:

- Odoo URL
- Database
- Username/email
- Password/API key

Credential disimpan di browser `localStorage`, bukan di GitHub.

## Catatan batasan Vercel

Studio2 v10 tetap memakai Vercel serverless, jadi jangan paksa satu request untuk export/import semua database. Gunakan pola aman:

- Import 10–30 row per batch.
- Export record terpilih, bukan full database mentah.
- Pilih field seperlunya.
- Hindari chatter, HTML panjang, dan image base64 kecuali benar-benar perlu.

