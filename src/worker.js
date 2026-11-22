/**
 * CBT TKA - SD Negeri Seda
 * Cloudflare Worker: menyajikan halaman (folder public) dan meneruskan /api ke Google Apps Script.
 * Alamat web app Apps Script disimpan di variabel APPS_SCRIPT_URL (lihat wrangler.jsonc).
 */
const MAKS_BODY = 5 * 1024 * 1024; // 5 MB: cukup untuk import soal dan upload gambar yang sudah diperkecil

function json(obj, status) {
  return new Response(JSON.stringify(obj), {
    status: status || 200,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname !== '/api') return env.ASSETS.fetch(request);

    if (request.method !== 'POST') return json({ success: false, message: 'Metode tidak didukung.' }, 405);
    if (!env.APPS_SCRIPT_URL) return json({ success: false, message: 'APPS_SCRIPT_URL belum diisi di Worker.' }, 500);

    const panjang = Number(request.headers.get('content-length') || 0);
    if (panjang > MAKS_BODY) return json({ success: false, message: 'Data terlalu besar.' }, 413);
    const body = await request.text();
    if (body.length > MAKS_BODY) return json({ success: false, message: 'Data terlalu besar.' }, 413);

    let hulu;
    try {
      // Apps Script membalas POST dengan redirect (302) ke alamat hasil; fetch mengikutinya otomatis.
      hulu = await fetch(env.APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: body,
        redirect: 'follow'
      });
    } catch (e) {
      return json({ success: false, net: true, message: 'Server database tidak dapat dihubungi. Coba lagi ya.' }, 502);
    }

    const teks = await hulu.text();
    try { JSON.parse(teks); } catch (e) {
      // Bukan JSON (misalnya halaman error/izin dari Google): jangan diteruskan mentah ke siswa.
      return json({ success: false, message: 'Server tidak membalas dengan benar. Hubungi proktor.' }, 502);
    }
    return new Response(teks, { headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } });
  }
};
