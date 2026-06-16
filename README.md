# Studio2 v10.2 — Smart Bundle Export

Studio2 v10.2 adalah versi Vercel-only yang tetap fokus pada **Import / Export / Review data Odoo**, tetapi sekarang Export sudah punya **Smart Bundle**.

Smart Bundle bukan fitur AI context. Fitur ini murni untuk membawa record Odoo beserta data yang memang berkaitan secara operasional.

## Apa yang baru di v10.2

### 1. Smart Bundle Export

Export tidak lagi hanya “satu model satu sheet”. Admin bisa memilih objek utama, lalu Studio2 mengekspor sheet relasi yang relevan.

Preset bundle:

- **Project Bundle**
  - `project.project`
  - `project.task`
  - `project.milestone`
  - `project.update`
  - `project.task.type`
  - `res.partner`
  - `res.users`
  - `task_hierarchy`
  - `relationship_map`
  - `README_EXPORT`

- **Contact Bundle**
  - `res.partner`
  - child contacts / address
  - `res.partner.category`
  - optional: `sale.order`
  - optional: `project.task`

- **Product Bundle**
  - `product.template`
  - `product.product`
  - `product.supplierinfo`
  - `product.category`
  - `product.public.category`
  - `uom.uom`

- **Sales Bundle**
  - `sale.order`
  - `sale.order.line`
  - related `res.partner`
  - related `product.product` / `product.template`
  - related `res.users`

- **Knowledge Bundle**
  - `knowledge.article`
  - child articles
  - parent article reference

### 2. Single Model tetap ada

Untuk pekerjaan cepat, admin tetap bisa export model biasa:

- `res.partner`
- `product.template`
- `project.project`
- `project.task`
- `knowledge.article`
- `sale.order`
- model custom

### 3. Review Workspace tetap menjadi quality gate

Hasil export bundle langsung masuk ke Review Workspace sebagai banyak sheet. Admin bisa:

- memilih sheet,
- cek schema,
- edit object card,
- buka grid,
- download XLSX,
- import sheet aktif.

## Flow utama

### Export Smart Bundle

1. Buka **Export**.
2. Pilih **Smart Bundle**.
3. Pilih Project Bundle / Contact Bundle / Product Bundle / Sales Bundle / Knowledge Bundle.
4. Klik **Scan record utama**.
5. Pilih record utama.
6. Klik **Export Bundle**.
7. Hasil multi-sheet masuk ke Review Workspace.

### Import validasi XLSX barcode scanner

1. Buka **Import**.
2. Upload XLSX dari web app barcode scanner.
3. Studio2 mendeteksi sheet `product.template` / product.
4. Admin cek foto, barcode, harga, vendor, kategori, dan field wajib.
5. Admin edit yang kosong.
6. Import batch kecil ke Odoo.

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

Studio2 tetap memakai Vercel serverless, jadi Smart Bundle harus dipakai secara bertahap:

- pilih record utama dulu,
- jangan export full database,
- hindari chatter dan image base64,
- gunakan bundle yang relevan saja,
- relasi opsional seperti sales/project pada Contact Bundle hanya dicentang jika perlu.
