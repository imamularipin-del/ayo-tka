# CBT TKA - SD Negeri Seda

Tampilan dilayani Cloudflare Worker, data tetap di Google Sheets lewat Google Apps Script.

```
public/index.html      halaman aplikasi (siswa dan admin)  -> dilayani Worker
src/worker.js          Worker: sajikan halaman + teruskan /api ke Apps Script
wrangler.jsonc         pengaturan Worker (isi APPS_SCRIPT_URL)
apps-script/Code.gs    server (tempel di project Apps Script)
apps-script/Index.html salinan sama dengan public/index.html (untuk mode Apps Script)
```

Catatan: jangan taruh `Code.gs` di folder `public` karena semua isi `public` terbuka untuk umum.
Setiap mengubah `public/index.html`, salin juga ke `apps-script/Index.html`.

## Langkah singkat
1. Apps Script: tempel `apps-script/Code.gs` dan `apps-script/Index.html`, lalu Deploy > New deployment
   (Execute as: Me, Who has access: Anyone). Salin URL yang berakhiran `/exec`.
2. Isi `APPS_SCRIPT_URL` di `wrangler.jsonc` dengan URL itu.
3. Unggah folder ini ke repo GitHub (private), lalu di Cloudflare: Workers & Pages > Create > Import a repository.
   Atau dari komputer: `npx wrangler deploy`.
4. Setiap kali `Code.gs` diubah, deploy ulang Apps Script dengan "New version" (URL `/exec` tidak berubah).
