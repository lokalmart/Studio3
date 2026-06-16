# Studio2 v9.1 — Web3 UI Vercel-only Clean Rebuild

Studio2 v9.1 adalah rebuild Next.js untuk Vercel dengan tampilan Web3 command center: glass panel, neon gradient, compact rail otomatis, workflow bar, editor native per model Odoo, dan import/export bertahap agar aman untuk Vercel gratis.

## Struktur

```text
/package.json
/next.config.mjs
/vercel.json
/src/app/page.tsx
/src/app/api/odoo/route.ts
/src/app/globals.css
/public/manifest.webmanifest
```

## Deploy ke Vercel

1. Upload semua isi ZIP ke root repo GitHub.
2. Import repo ke Vercel.
3. Framework: Next.js.
4. Root Directory: default/kosong.
5. Deploy.

## Cara pakai

1. Buka aplikasi.
2. Klik **Koneksi** dan isi target Odoo.
3. Import: upload XLSX → pilih sheet → load schema → edit → import batch.
4. Export: isi model → scan → pilih record → export ke editor → download XLSX.

## Catatan desain

- Panel kiri otomatis berubah menjadi compact rail setelah data aktif agar editor menjadi fokus utama.
- Editor dibedakan berdasarkan model Odoo: contact, product, project, knowledge, sales, dynamic.
- XLSX tetap diproses di browser; API Vercel hanya menerima batch kecil.
