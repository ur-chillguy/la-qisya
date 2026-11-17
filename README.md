# La Qisya — Rental & Sewa Gaun

Showcase gaun untuk calon customer + dashboard pemilik (booking, stok, foto, keuangan, pelanggan, pengaturan).

- `index.html` — front-end (dipakai GitHub Pages). Tanpa backend = **mode demo** (data di browser).
- `apps-script/` — backend Google Apps Script + Sheets + Drive (`Code.gs`, `Index.html`, `appsscript.json`).

## Menyambungkan ke backend sungguhan
1. Deploy `apps-script/` sebagai Web App (Execute as: Me, Access: Anyone), jalankan `setup()` sekali.
2. Di `index.html`, isi `window.LQ_API_URL = "https://script.google.com/macros/s/…/exec"`.
3. Commit — GitHub Pages otomatis memakai data live.

PIN awal pemilik: `123456` (wajib diganti di menu Pengaturan).

Dibuat oleh [ur.chillguy](https://instagram.com/ur.chillguy)
