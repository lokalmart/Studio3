# Lokalmart Studio2 v8 — Hybrid Vercel + Render

Studio2 v8 memisahkan UI dan mesin berat:

- `apps/web` — Next.js frontend untuk Vercel.
- `apps/engine` — Express backend/worker untuk Render.

Model ini dibuat supaya import/export Odoo tidak lagi bergantung pada request pendek Vercel. UI tetap cepat di Vercel, sedangkan scan record, export selected records, export project, dan import XLSX batch berjalan di Render.

## Struktur deploy

### 1. Deploy Engine ke Render

Root Directory: `apps/engine`

Build Command:

```bash
npm install
```

Start Command:

```bash
npm start
```

Environment Variables:

```bash
NODE_VERSION=20
STUDIO2_ENGINE_API_KEY=isi_secret_sendiri
CORS_ORIGIN=https://domain-vercel-studio2-kamu.vercel.app
JOB_DIR=/tmp/studio2-jobs
```

Setelah deploy, cek:

```text
https://nama-render-service.onrender.com/health
```

### 2. Deploy Web ke Vercel

Root Directory: `apps/web`

Environment Variable opsional:

```bash
NEXT_PUBLIC_DEFAULT_ENGINE_URL=https://nama-render-service.onrender.com
```

Di UI Studio2, buka tombol pengaturan, isi:

- Engine URL Render
- Engine API Key
- URL Odoo
- Database Odoo
- Username/email Odoo
- Password/API key Odoo

Data target disimpan di browser/localStorage, bukan di repo.

## Workflow Import

1. Upload XLSX.
2. Preview sheet.
3. Editor otomatis menyesuaikan model:
   - `res.partner` → Contact Editor
   - `product.template` / `product.product` → Product Editor
   - `project.project` / `project.task` → Project Editor
   - `knowledge.article` → Knowledge Editor
   - model lain → Dynamic Editor
4. Edit dan validasi.
5. Kirim job import ke Render Engine.
6. Pantau progress sampai selesai.

## Workflow Export

1. Pilih model atau project.
2. Scan record dari Odoo.
3. Pilih record yang ingin diexport.
4. Pilih field yang ingin ikut export.
5. Jalankan export job di Render.
6. Hasil export masuk kembali ke preview/editor XLSX.
7. Download XLSX.

## Kenapa v8 berbeda dari patch HTML lama?

Patch lama masih single `index.html` dan satu endpoint Vercel. Itu cepat untuk eksperimen, tetapi mentok untuk editor dinamis. v8 memakai framework agar UI lebih responsif, panel bisa fokus/collapse, schema Odoo bisa dipakai untuk editor, dan job berat tidak memicu timeout Vercel.

## Catatan keamanan

- Jangan commit kredensial Odoo ke GitHub.
- Pakai `STUDIO2_ENGINE_API_KEY` di Render.
- Batasi `CORS_ORIGIN` ke domain Vercel kamu setelah URL final sudah ada.
- Untuk file besar/foto produk, tetap pecah import menjadi batch dan jangan upload foto massal dalam satu tahap.
