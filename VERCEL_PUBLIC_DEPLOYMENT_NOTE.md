# Vercel Public Deployment Note

Kalau URL Vercel meminta login, periksa Deployment Protection / Vercel Authentication.

Untuk PWA/APK, gunakan production deployment yang public/unprotected.

Studio2 adalah admin tool internal, jadi keamanan utama tetap:

- Jangan commit password/API key Odoo ke GitHub.
- Simpan credential hanya di browser.
- Pakai akun Odoo dengan hak secukupnya.
