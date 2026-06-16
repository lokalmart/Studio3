# File lama yang perlu dihapus dari Studio2 repo

Karena v8 adalah rebuild hybrid, hapus file lama yang bentrok:

- `index.html`
- `assistant.html`
- `api/odoo.js`
- `scripts/check-html.js`
- patch README lama jika sudah tidak dipakai

Setelah itu copy struktur v8 ini ke repo.

Deploy:

- Vercel Root Directory: `apps/web`
- Render Root Directory: `apps/engine`
