# ToWord — PDF ke Word gratis, tanpa watermark

Web untuk mengubah PDF menjadi Word (.docx) yang **rapi dan bisa diedit**, dengan
**tanda tangan & stempel tetap ada**, **tanpa watermark**, dan bisa membaca PDF
**hasil scan / foto** (OCR Bahasa Indonesia + Inggris).

## Fitur
- **Mode Otomatis**: setiap halaman dideteksi sendiri.
  - Halaman PDF digital → dikonversi langsung (paragraf, font, tabel, gambar, ttd, stempel tetap terjaga).
  - Halaman hasil scan/foto, teks rusak (simbol aneh), atau lapisan OCR tersembunyi → dibaca dengan OCR Tesseract.
- **Scan jadi teks yang bisa diedit**: teks jadi paragraf Word asli (ukuran huruf, tebal,
  rata tengah/kiri/justify, tab "Nomor : …", jarak baris) — bukan gambar.
- **Tanda tangan, stempel, logo & foto** dipotong dari scan dan disisipkan sebagai gambar
  di posisinya dengan **latar transparan**. Stempel yang menimpa teks ditaruh melayang di atas
  teks, dan teks di bawah stempel tetap dibaca (OCR ulang setelah warna stempel dihilangkan).
- Scan miring diluruskan otomatis, halaman terbalik/menyamping diputar otomatis.
- **Anti gagal**: jika satu cara gagal, otomatis pindah ke cara lain (terakhir: halaman sebagai gambar).
- PDF berpassword didukung; PDF asli dihapus setelah diproses, hasil dihapus otomatis (default 30 menit).
- Banyak file sekaligus, tampilan responsif (HP & desktop), mode gelap.

## Pasang di hosting (agar bisa dibuka online)

### Render.com — gratis, paling mudah
[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/shavel28/ToWord)

1. Pastikan kode ini sudah ada di branch `main` (merge dulu jika masih di branch lain).
2. Klik tombol **Deploy to Render** di atas → masuk/daftar dengan GitHub.
3. Render membaca `render.yaml` otomatis (Docker, paket Free, pengaturan sudah terisi).
   Isi nama layanan bila diminta, lalu klik **Apply / Deploy**.
4. Tunggu ±5–10 menit sampai status **Live**, lalu buka alamatnya,
   misalnya `https://toword.onrender.com`.

Cara manual (tanpa tombol): Render → **New → Web Service** → pilih repo ini →
Runtime **Docker** → Instance **Free** → tambahkan env `TOWORD_WORKERS=1` dan
`TOWORD_MAX_MB=30` → **Deploy**.

Catatan paket gratis: server "tidur" setelah 15 menit tidak dipakai (buka pertama
±30–60 detik), dan RAM kecil (512 MB) sehingga PDF scan yang sangat tebal bisa lambat.
Untuk pemakaian rutin gunakan paket berbayar atau VPS.

### VPS (Ubuntu) dengan Docker
```bash
sudo apt update && sudo apt install -y docker.io git
git clone https://github.com/shavel28/ToWord.git && cd ToWord
sudo docker build -t toword .
sudo docker run -d --restart always -p 80:8000 --name toword toword
```
Buka `http://IP-VPS-anda`. Update versi baru:
`git pull && sudo docker build -t toword . && sudo docker rm -f toword && sudo docker run -d --restart always -p 80:8000 --name toword toword`

## Menjalankan di komputer sendiri

### Docker (paling mudah)
```bash
docker build -t toword .
docker run -p 8000:8000 toword
```
Buka http://localhost:8000

### Tanpa Docker
Butuh Python 3.10+ dan Tesseract:
```bash
# Ubuntu/Debian
sudo apt install tesseract-ocr tesseract-ocr-ind tesseract-ocr-eng
# Windows: install dari https://github.com/UB-Mannheim/tesseract/wiki (centang bahasa Indonesian)

pip install -r requirements.txt
python app.py
```

## Konfigurasi (environment variable)
| Variabel | Default | Keterangan |
|---|---|---|
| `PORT` | 8000 | Port web |
| `TOWORD_MAX_MB` | 100 | Ukuran maksimal file |
| `TOWORD_KEEP_MINUTES` | 30 | Lama hasil disimpan sebelum dihapus |
| `TOWORD_WORKERS` | CPU/2 | Jumlah konversi paralel |
| `TOWORD_OCR_LANG` | `ind+eng` | Bahasa OCR Tesseract |
| `TOWORD_OCR_FONT` | `Arial` | Font untuk teks hasil OCR |

## Struktur
- `app.py` — server web Flask (unggah, antrean, progres, unduh).
- `converter.py` — mesin konversi (pdf2docx + OCR Tesseract + deteksi ttd/stempel).
- `templates/index.html` — tampilan web.

Bisa juga dipakai dari Python:
```python
from converter import convert_pdf_to_docx
convert_pdf_to_docx("surat.pdf", "surat.docx", mode="auto")  # "auto" | "ocr" | "layout"
```
