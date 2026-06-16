# File lama yang perlu dihapus dari Studio2 repo

Karena v8.1 adalah rebuild Render-only, hapus file lama yang bentrok:

- `index.html`
- `assistant.html`
- `api/odoo.js`
- `scripts/check-html.js`
- patch README lama jika sudah tidak dipakai

Setelah itu copy struktur v8.1 ini ke repo.

Deploy di Render sebagai dua Web Service:

- `apps/web` → Studio2 UI
- `apps/engine` → Studio2 Engine
