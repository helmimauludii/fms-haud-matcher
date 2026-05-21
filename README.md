# FMS HAUD Matching Dashboard

Dashboard statis untuk membantu user melakukan matching data FMS dan HAUD langsung dari browser.

## Fitur

- Upload file FMS dan HAUD dalam format `.xlsx`, `.xls`, atau `.csv`
- Pilih sheet dari masing-masing file
- Mapping kolom nomor/reference dan kolom tanggal
- Default mapping sesuai format terbaru:
  - FMS: `Date` dan `B-Number Processed`
  - HAUD: `Date`, `sourceAddr`, dan `destinationAddr`
- Pilih range tanggal FMS yang ingin diproses
- Range tanggal hanya memfilter FMS
- HAUD dibaca penuh, lalu dicocokkan ke setiap FMS berdasarkan H-1, H, dan H+1
- Normalisasi nomor/reference
- Preview data dan hasil matching
- Download output dalam format CSV atau JSON dengan kolom:
  `fms_row`, `fms_date`, `B-Number Processed`, `haud_row`, `haud_date`, `sourceAddr`, `destinationAddr`, `match_type`

## Cara jalan lokal

Buka `index.html` langsung di browser, atau jalankan static server:

```bash
python3 -m http.server 4173
```

Lalu buka:

```text
http://localhost:4173
```

## Deploy gratis

### Cloudflare Pages

1. Push folder ini ke GitHub.
2. Buka Cloudflare Pages.
3. Pilih repo.
4. Framework preset: `None`.
5. Build command: kosongkan.
6. Output directory: `/`.
7. Deploy.

### Netlify

1. Push folder ini ke GitHub atau drag-and-drop folder ini ke Netlify.
2. Build command: kosongkan.
3. Publish directory: `.`.
4. Deploy.

### Vercel

1. Push folder ini ke GitHub.
2. Import project di Vercel.
3. Framework preset: `Other`.
4. Build command: kosongkan.
5. Output directory: `.`.
6. Deploy.

## Catatan

File diproses di browser user. Versi ini tidak mengirim file ke backend, sehingga cocok untuk prototype publik gratis. Jika nanti butuh audit trail, login, penyimpanan file, atau proses Python server-side, dashboard ini bisa disambungkan ke API.
