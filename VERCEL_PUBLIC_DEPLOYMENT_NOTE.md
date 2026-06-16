# Vercel Public Deployment Note

Untuk PWA/APK, gunakan production domain yang public.

Kalau muncul halaman login Vercel:

1. Pastikan kamu membuka production domain, bukan preview deployment.
2. Cek Project Settings → Deployment Protection.
3. Matikan Vercel Authentication untuk environment yang ingin dibuka publik.
4. Kalau ingin tetap terlindungi, gunakan password internal di aplikasi Studio2, bukan Vercel preview auth.

Studio2 v9 tidak membutuhkan Render, Cloudflare, ataupun PC lokal aktif.
