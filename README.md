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

## Menjalankan

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
