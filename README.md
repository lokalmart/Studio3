# Lokalmart Studio2 v8.1 — Render-only Web + Engine

Studio2 v8.1 meninggalkan Vercel. Semua service berjalan di Render:

- `apps/web` — Next.js frontend sebagai Render Web Service.
- `apps/engine` — Express backend/engine sebagai Render Web Service.

Model ini lebih cocok untuk Lokalmart karena import/export Odoo bisa berjalan sebagai job di backend Render, sedangkan UI tetap menjadi aplikasi editor yang rapi dan responsif.

## Deploy cepat di Render

Ada dua cara.

### Opsi A — pakai Blueprint `render.yaml`

1. Push repo ke GitHub.
2. Di Render pilih **New + → Blueprint**.
3. Pilih repo Studio2.
4. Render akan membaca `render.yaml` dan membuat dua Web Service:
   - `studio2-web`
   - `studio2-engine`
5. Setelah deploy, buka URL `studio2-web`.
6. Di tombol ⚙ Studio2, isi:
   - Engine URL: URL service `studio2-engine`, contoh `https://studio2-engine.onrender.com`
   - Engine API Key: nilai `STUDIO2_ENGINE_API_KEY` dari Render
   - URL Odoo
   - Database Odoo
   - Username/email Odoo
   - Password/API key Odoo

### Opsi B — buat dua Web Service manual

#### 1. Studio2 Engine

Service type: **Web Service**

Root Directory:

```text
apps/engine
```

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
CORS_ORIGIN=*
JOB_DIR=/tmp/studio2-jobs
```

Tes engine:

```text
https://nama-engine.onrender.com/health
```

#### 2. Studio2 Web

Service type: **Web Service**

Root Directory:

```text
apps/web
```

Build Command:

```bash
npm install && npm run build
```

Start Command:

```bash
npm start
```

Environment Variables:

```bash
NODE_VERSION=20
NEXT_PUBLIC_DEFAULT_ENGINE_URL=https://nama-engine.onrender.com
```

Catatan: `NEXT_PUBLIC_DEFAULT_ENGINE_URL` hanya default awal. Kalau lupa diisi, kamu tetap bisa isi Engine URL manual dari tombol ⚙ di UI.

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
5. Jalankan export job di Render Engine.
6. Hasil export masuk kembali ke preview/editor XLSX.
7. Download XLSX.

## Catatan keamanan

- Jangan commit kredensial Odoo ke GitHub.
- Pakai `STUDIO2_ENGINE_API_KEY` di Render.
- Setelah URL final sudah stabil, sebaiknya ganti `CORS_ORIGIN=*` menjadi URL `studio2-web` kamu, misalnya `https://studio2-web.onrender.com`.
- Data target Odoo disimpan di browser/localStorage, bukan di repo.
- Untuk import foto massal, tetap pisahkan dari import data inti supaya job tidak berat.

## File lama yang perlu dihapus

Lihat `DELETE_OLD_FILES_NOTE.md`.
