"""
Mesin konversi PDF -> Word (DOCX) untuk ToWord.

Strategi per halaman:
  * Halaman PDF "digital" (punya lapisan teks asli)  -> pdf2docx
    (tata letak, tabel, gambar, tanda tangan & stempel tetap terjaga).
  * Halaman hasil scan / foto / teks rusak / lapisan OCR tersembunyi
    -> OCR Tesseract (Indonesia + Inggris). Teks dijadikan paragraf yang
    bisa diedit, sedangkan area non-teks (tanda tangan, stempel, logo, foto)
    dipotong dan disisipkan sebagai gambar pada posisinya.
  * Jika satu cara gagal, otomatis pindah ke cara berikutnya, dan pilihan
    terakhir adalah menyisipkan halaman sebagai gambar utuh sehingga
    konversi tidak pernah gagal total.

Tidak ada watermark yang ditambahkan ke hasil.
"""

from __future__ import annotations

import io
import logging
import os
import re
import statistics
import tempfile
import warnings
from dataclasses import dataclass, field
from typing import Callable, List, Optional

warnings.filterwarnings("ignore", category=FutureWarning)

import cv2  # noqa: E402
import numpy as np  # noqa: E402
import pymupdf  # noqa: E402
import pytesseract  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Pt  # noqa: E402

log = logging.getLogger("toword")

OCR_DPI = 300
OCR_LANG = os.environ.get("TOWORD_OCR_LANG", "ind+eng")
PX2PT = 72.0 / OCR_DPI
OCR_FONT = os.environ.get("TOWORD_OCR_FONT", "Arial")

ProgressFn = Callable[[float, str], None]


class ConversionError(Exception):
    pass


class PasswordRequired(ConversionError):
    pass


# --------------------------------------------------------------------------
# Analisis halaman
# --------------------------------------------------------------------------

_BAD_CHAR = re.compile(r"[�-\x00-\x08\x0b\x0c\x0e-\x1f]")


def _page_needs_ocr(page: "pymupdf.Page") -> bool:
    """True jika halaman sebaiknya di-OCR (scan, teks rusak, teks tak terlihat)."""
    text = page.get_text("text") or ""
    visible = text.strip()
    page_area = abs(page.rect) or 1.0

    img_area = 0.0
    try:
        for info in page.get_image_info():
            r = pymupdf.Rect(info["bbox"]) & page.rect
            img_area += abs(r)
    except Exception:
        pass
    img_cover = img_area / page_area

    if len(visible) < 25:
        # Hampir tidak ada teks. Kalau ada gambar / gambar vektor -> OCR.
        if img_cover > 0.05:
            return True
        try:
            if len(page.get_drawings()) > 30:  # teks yang digambar sebagai kurva
                return True
        except Exception:
            pass
        return False

    # Teks rusak (font tanpa peta Unicode)
    bad = len(_BAD_CHAR.findall(visible))
    if bad / max(len(visible), 1) > 0.15:
        return True

    # Lapisan OCR tersembunyi di atas gambar scan (render mode 3)
    if img_cover > 0.6:
        try:
            total = hidden = 0
            for span in page.get_texttrace():
                n = len(span.get("chars", ()))
                total += n
                if span.get("type") == 3 or span.get("opacity", 1) == 0:
                    hidden += n
            if total and hidden / total > 0.5:
                return True
        except Exception:
            pass
    return False


# --------------------------------------------------------------------------
# OCR halaman
# --------------------------------------------------------------------------

@dataclass
class Word:
    text: str
    x0: int
    y0: int
    x1: int
    y1: int
    conf: float
    key: tuple  # (block, par, line)
    fsize: float = 0.0  # perkiraan ukuran huruf (px) dari tesseract


@dataclass
class Element:
    kind: str  # "para" | "image"
    x0: float
    y0: float
    x1: float
    y1: float
    lines: list = field(default_factory=list)  # para: list[list[Word]]
    image: Optional[np.ndarray] = None
    font_pt: float = 11.0
    bold: bool = False
    floating: bool = False  # gambar menimpa teks (mis. stempel di atas nama)

    @property
    def h(self):
        return self.y1 - self.y0

    @property
    def w(self):
        return self.x1 - self.x0


def _render(page, dpi=OCR_DPI) -> np.ndarray:
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    return img.copy()


def _auto_rotate(img: np.ndarray) -> np.ndarray:
    try:
        osd = pytesseract.image_to_osd(img, config="--psm 0 -c min_characters_to_try=20",
                                       output_type=pytesseract.Output.DICT)
        rot = int(osd.get("rotate", 0))
        if float(osd.get("orientation_conf", 0)) < 2.0:
            return img
        if rot == 90:
            return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
        if rot == 180:
            return cv2.rotate(img, cv2.ROTATE_180)
        if rot == 270:
            return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    except Exception:
        pass
    return img


def _skew_angle(gray: np.ndarray) -> float:
    """Cari sudut kemiringan dengan profil proyeksi (maks ±5 derajat)."""
    small = cv2.resize(gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    bw = cv2.threshold(small, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    h, w = bw.shape
    c = (w / 2, h / 2)

    def score(a):
        m = cv2.getRotationMatrix2D(c, a, 1.0)
        r = cv2.warpAffine(bw, m, (w, h), flags=cv2.INTER_NEAREST)
        prof = r.sum(axis=1, dtype=np.float64)
        return float(np.sum(np.diff(prof) ** 2))

    best = max(np.arange(-5, 5.01, 0.5), key=score)
    best = max(np.arange(best - 0.5, best + 0.51, 0.05), key=score)
    return float(best)


def _deskew(img: np.ndarray) -> np.ndarray:
    """Luruskan scan yang sedikit miring."""
    try:
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        angle = _skew_angle(gray)
        if abs(angle) < 0.1:
            return img
        h, w = img.shape[:2]
        m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
        return cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC,
                              borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))
    except Exception:
        return img


def _hocr_title(t: str) -> dict:
    out = {}
    for part in (t or "").split(";"):
        bits = part.strip().split()
        if bits:
            out[bits[0]] = bits[1:]
    return out


def _ocr_words(img: np.ndarray, lang: str, psm: int = 3) -> List[Word]:
    """OCR via hOCR: posisi kata, keyakinan, dan perkiraan ukuran huruf per baris."""
    import xml.etree.ElementTree as ET
    raw = pytesseract.image_to_pdf_or_hocr(img, lang=lang, extension="hocr",
                                           config=f"--oem 1 --psm {psm}")
    root = ET.fromstring(raw)
    words: List[Word] = []
    pc = lc = 0
    for el in root.iter():
        cls = el.get("class", "")
        if cls == "ocr_par":
            pc += 1
        elif cls in ("ocr_line", "ocr_header", "ocr_caption", "ocr_textfloat"):
            lc += 1
            t = _hocr_title(el.get("title"))
            try:
                xsize = float(t.get("x_size", [0])[0])
            except ValueError:
                xsize = 0.0
            for w in el.iter():
                if "ocrx_word" not in w.get("class", ""):
                    continue
                txt = "".join(w.itertext()).strip()
                if not txt:
                    continue
                wt = _hocr_title(w.get("title"))
                x0, y0, x1, y1 = (int(v) for v in wt["bbox"])
                conf = float(wt.get("x_wconf", [0])[0])
                wd = Word(txt, x0, y0, x1, y1, conf, (0, pc, lc))
                wd.fsize = xsize
                words.append(wd)
    return words


def _overlap_ratio(a, b) -> float:
    """Luas irisan / luas a. a,b = (x0,y0,x1,y1)."""
    ix = max(0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0, min(a[3], b[3]) - max(a[1], b[1]))
    area = max(1, (a[2] - a[0]) * (a[3] - a[1]))
    return ix * iy / area


def _merge_boxes(boxes, gap=20):
    boxes = [list(b) for b in boxes]
    changed = True
    while changed:
        changed = False
        out = []
        while boxes:
            b = boxes.pop()
            i = 0
            while i < len(boxes):
                o = boxes[i]
                if (b[0] - gap < o[2] and o[0] - gap < b[2] and
                        b[1] - gap < o[3] and o[1] - gap < b[3]):
                    b = [min(b[0], o[0]), min(b[1], o[1]), max(b[2], o[2]), max(b[3], o[3])]
                    boxes.pop(i)
                    changed = True
                else:
                    i += 1
            out.append(b)
        boxes = out
    return boxes


def _find_graphic_regions(img: np.ndarray, words: List[Word], text_h: float):
    """Temukan tanda tangan, stempel, logo & foto (area tinta yang bukan teks)."""
    H, W = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)

    # Tinta berwarna (stempel biru/merah/ungu, tanda tangan biru)
    color = ((hsv[..., 1] > 70) & (hsv[..., 2] < 245) & (hsv[..., 2] > 40)).astype(np.uint8) * 255
    color = cv2.morphologyEx(color, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))

    # Tinta gelap
    dark = (gray < 150).astype(np.uint8) * 255
    # Buang garis panjang (garis tabel, garis bawah)
    hk = cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, W // 12), 1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, H // 12)))
    lines = cv2.morphologyEx(dark, cv2.MORPH_OPEN, hk) | cv2.morphologyEx(dark, cv2.MORPH_OPEN, vk)
    dark = cv2.subtract(dark, cv2.dilate(lines, np.ones((3, 3), np.uint8)))

    # Hapus kata yang dikenali dengan yakin dari tinta gelap
    for wd in words:
        if wd.conf >= 45 and (wd.y1 - wd.y0) < text_h * 3:
            cv2.rectangle(dark, (wd.x0 - 2, wd.y0 - 2), (wd.x1 + 2, wd.y1 + 2), 0, -1)

    boxes = []
    min_w, min_h = int(text_h * 2.2), int(text_h * 1.6)
    for mask, kernel, is_color in ((color, 15, True), (dark, 21, False)):
        dil = cv2.dilate(mask, cv2.getStructuringElement(cv2.MORPH_RECT, (kernel, kernel)))
        cnts, _ = cv2.findContours(dil, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in cnts:
            x, y, w, h = cv2.boundingRect(c)
            if w < min_w or h < min_h:
                continue
            if w * h > 0.85 * W * H:
                continue
            ink = cv2.countNonZero(mask[y:y + h, x:x + w])
            if ink < 250:
                continue
            box = (x, y, x + w, y + h)
            if is_color:
                # Teks berwarna biasa (judul, link) -> tetap teks
                covered = sum(_overlap_ratio((wd.x0, wd.y0, wd.x1, wd.y1), box) > 0.6 and wd.conf >= 60
                              for wd in words)
                txt_area = sum((wd.x1 - wd.x0) * (wd.y1 - wd.y0) for wd in words
                               if wd.conf >= 60 and _overlap_ratio((wd.x0, wd.y0, wd.x1, wd.y1), box) > 0.6)
                if covered and txt_area > 0.45 * w * h and w / max(h, 1) > 2.5:
                    continue
            boxes.append(box)

    boxes = _merge_boxes(boxes, gap=int(text_h))
    pad = max(6, int(text_h * 0.3))
    out = []
    for x0, y0, x1, y1 in boxes:
        x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
        x1, y1 = min(W, x1 + pad), min(H, y1 + pad)
        out.append((x0, y0, x1, y1))
    return out, color, dark


def _is_colored(hsv: np.ndarray, box) -> bool:
    x0, y0, x1, y1 = box
    sat = hsv[y0:y1, x0:x1, 1]
    ink = hsv[y0:y1, x0:x1, 2] < 160
    return bool(ink.any()) and float(np.mean(sat[ink] > 70)) > 0.35


def _rescue_words(img, hsv, region, idx, existing, lang, text_h) -> List[Word]:
    """OCR ulang area stempel/ttd setelah tinta berwarna dihapus,
    supaya teks hitam yang tertimpa stempel tetap terbaca."""
    x0, y0, x1, y1 = region
    crop = img[y0:y1, x0:x1].copy()
    colored = (hsv[y0:y1, x0:x1, 1] > 60) & (hsv[y0:y1, x0:x1, 2] > 40)
    crop[colored] = 255
    try:
        data = pytesseract.image_to_data(crop, lang=lang, config="--oem 1 --psm 11",
                                         output_type=pytesseract.Output.DICT)
    except Exception:
        return []
    found = []
    for i, txt in enumerate(data["text"]):
        txt = (txt or "").strip()
        conf = float(data["conf"][i])
        if not txt or conf < 70 or len(re.findall(r"\w", txt)) < 2:
            continue
        bx = (x0 + data["left"][i], y0 + data["top"][i],
              x0 + data["left"][i] + data["width"][i], y0 + data["top"][i] + data["height"][i])
        if not (text_h * 0.5 < bx[3] - bx[1] < text_h * 2.5):
            continue
        if any(_overlap_ratio(bx, (w.x0, w.y0, w.x1, w.y1)) > 0.3 for w in existing):
            continue
        found.append(Word(txt, *bx, conf, (0, 0, 0), fsize=(bx[3] - bx[1]) * 1.15))
    # kelompokkan per baris
    found.sort(key=lambda w: (w.y0 + w.y1) / 2)
    line, last_c = 0, None
    for w in found:
        c = (w.y0 + w.y1) / 2
        if last_c is not None and abs(c - last_c) > text_h * 0.6:
            line += 1
        w.key = (100000 + idx, 0, line)
        last_c = c
    return found


def _font_pt_from_height(px: float) -> float:
    pt = px * PX2PT * 1.04
    return max(7.0, min(36.0, round(pt * 2) / 2))


def _build_elements(img: np.ndarray, lang: str) -> List[Element]:
    words = _ocr_words(img, lang)
    heights = [w.y1 - w.y0 for w in words if w.conf >= 50 and len(w.text) >= 2]
    text_h = statistics.median(heights) if heights else 40.0

    regions, color_mask, dark_mask = _find_graphic_regions(img, words, text_h)
    hsv = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
    for i, r in enumerate(regions):
        words += _rescue_words(img, hsv, r, i, words, lang, text_h)

    # Kata di dalam area gambar: teks hitam yang terbaca jelas tetap jadi
    # teks (dan dihapus dari gambar), sisanya (tulisan stempel, coretan ttd)
    # dibuang agar tidak dobel.
    kept = []
    for wd in words:
        wb = (wd.x0, wd.y0, wd.x1, wd.y1)
        hits = [i for i, r in enumerate(regions) if _overlap_ratio(wb, r) > 0.3]
        if hits:
            if wd.conf >= 45 and not _is_colored(hsv, wb) and re.search(r"\w", wd.text):
                cv2.rectangle(dark_mask, (wd.x0 - 3, wd.y0 - 3), (wd.x1 + 3, wd.y1 + 3), 0, -1)
            else:
                continue
        elif wd.conf < 20 and not re.search(r"\w", wd.text):
            continue
        kept.append(wd)
    # Kelompokkan per baris lalu per paragraf tesseract
    lines = {}
    for wd in kept:
        lines.setdefault(wd.key, []).append(wd)
    paras = {}
    for key, ws in lines.items():
        ws.sort(key=lambda w: w.x0)
        paras.setdefault(key[:2], []).append(ws)

    elements: List[Element] = []
    for _, plines in paras.items():
        plines.sort(key=lambda ln: min(w.y0 for w in ln))
        # pecah paragraf jika jarak antar baris terlalu jauh
        group = [plines[0]]
        for ln in plines[1:]:
            prev = group[-1]
            gap = min(w.y0 for w in ln) - max(w.y1 for w in prev)
            if gap > text_h * 1.4:
                elements.append(_para_from_lines(group))
                group = [ln]
            else:
                group.append(ln)
        elements.append(_para_from_lines(group))

    ink_mask = color_mask | dark_mask
    n_cc, lab, stats, _ = cv2.connectedComponentsWithStats(ink_mask, connectivity=8)
    small = np.where(stats[:, cv2.CC_STAT_AREA] < max(40, int(text_h * 1.5)))[0]
    small = small[small != 0]
    if len(small):
        ink_mask[np.isin(lab, small)] = 0
    ink_mask = cv2.dilate(ink_mask, np.ones((3, 3), np.uint8))
    for i, (x0, y0, x1, y1) in enumerate(regions):
        crop = img[y0:y1, x0:x1]
        rgba = _transparent_bg(crop, ink_mask[y0:y1, x0:x1])
        floating = any(
            _overlap_ratio((w.x0, w.y0, w.x1, w.y1), (x0, y0, x1, y1)) > 0.05 for w in kept)
        elements.append(Element("image", x0, y0, x1, y1, image=rgba, floating=floating))

    # Deteksi huruf tebal dari ketebalan goresan (piksel tinta / panjang tepi)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    def stroke(w):
        b = gray[w.y0:w.y1, w.x0:w.x1] < 140
        area = int(b.sum())
        if area < 20:
            return None
        edges = int((b[:, 1:] != b[:, :-1]).sum() + (b[1:, :] != b[:-1, :]).sum())
        return 2.0 * area / max(edges, 1)
    sw = {}
    for e in elements:
        if e.kind == "para":
            for ln in e.lines:
                for w in ln:
                    v = stroke(w)
                    if v:
                        sw[id(w)] = v
    if len(sw) > 5:
        base = statistics.median(sw.values())
        for e in elements:
            if e.kind != "para":
                continue
            vals = [sw[id(w)] for ln in e.lines for w in ln if id(w) in sw]
            if vals and statistics.median(vals) > base * 1.3:
                e.bold = True

    # Ukuran huruf: samakan ukuran yang mirip dengan ukuran isi dokumen
    sizes = [e.font_pt for e in elements if e.kind == "para"]
    if sizes:
        weights = []
        for e in elements:
            if e.kind == "para":
                weights += [e.font_pt] * sum(len(ln) for ln in e.lines)
        body = statistics.median(weights)
        for e in elements:
            if e.kind == "para" and 0.88 <= e.font_pt / body <= 1.12:
                e.font_pt = body
    return elements


def _transparent_bg(crop: np.ndarray, ink: np.ndarray) -> np.ndarray:
    """Hanya tinta grafis (ttd/stempel/logo) yang terlihat; kertas & teks jadi transparan."""
    hsv = cv2.cvtColor(crop, cv2.COLOR_RGB2HSV)
    gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY).astype(np.float32)
    # transisi halus: <=170 penuh, >=235 transparan
    alpha = np.clip((235.0 - gray) / 65.0, 0, 1)
    alpha = np.maximum(alpha, (hsv[..., 1] > 60).astype(np.float32) * (hsv[..., 2] < 245))
    # foto/gambar padat (banyak tinta) dibiarkan utuh tanpa transparansi
    if float(np.mean(gray < 200)) > 0.45:
        return crop.copy()
    alpha *= (ink > 0).astype(np.float32)
    return np.dstack([crop, (alpha * 255).astype(np.uint8)])


def _para_from_lines(lines) -> Element:
    xs0 = [w.x0 for ln in lines for w in ln]
    xs1 = [w.x1 for ln in lines for w in ln]
    ys0 = [w.y0 for ln in lines for w in ln]
    ys1 = [w.y1 for ln in lines for w in ln]
    fs = [w.fsize for ln in lines for w in ln if w.fsize > 0]
    if fs:
        size = statistics.median(fs)
    else:
        hs = sorted(w.y1 - w.y0 for ln in lines for w in ln)
        size = hs[min(len(hs) - 1, int(len(hs) * 0.9))]
    return Element("para", min(xs0), min(ys0), max(xs1), max(ys1), lines=lines,
                   font_pt=_font_pt_from_height(size))


# --------------------------------------------------------------------------
# Penulisan DOCX untuk halaman OCR
# --------------------------------------------------------------------------

def _add_png(run, image: np.ndarray, width_pt: float):
    if image.shape[2] == 4:
        bgr = cv2.cvtColor(image, cv2.COLOR_RGBA2BGRA)
    else:
        bgr = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    ok, buf = cv2.imencode(".png", bgr)
    run.add_picture(io.BytesIO(buf.tobytes()), width=Pt(width_pt))


_ANCHOR_ID = [1000]


def _inline_to_anchor(run, x_pt: float, y_pt: float):
    """Ubah gambar inline jadi gambar melayang di posisi absolut halaman."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    inline = run._r.xpath(".//wp:inline")[-1]
    cx, cy = inline.extent.cx, inline.extent.cy
    _ANCHOR_ID[0] += 1
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="0" distL="0" distR="0" '
        f'simplePos="0" relativeHeight="{251658240 + _ANCHOR_ID[0]}" behindDoc="0" locked="0" '
        f'layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(x_pt * 12700)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(y_pt * 12700)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/>'
        f'<wp:docPr id="{_ANCHOR_ID[0]}" name="Gambar {_ANCHOR_ID[0]}"/>'
        f'<wp:cNvGraphicFramePr/></wp:anchor>')
    anchor.append(inline.graphic)
    inline.getparent().replace(inline, anchor)


def _set_cell_margins_zero(table):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _bands(elements: List[Element]):
    """Kelompokkan elemen yang bersebelahan secara horizontal (mis. 2 kolom ttd)."""
    elements = sorted(elements, key=lambda e: (e.y0, e.x0))
    bands = []
    for e in elements:
        if bands:
            b = bands[-1]
            by1 = max(x.y1 for x in b)
            ov = by1 - e.y0
            if ov > 0.35 * min(e.h, max(1, by1 - min(x.y0 for x in b))) and ov > 0:
                # hanya gabung bila tidak menumpuk secara horizontal dengan semua
                b.append(e)
                continue
        bands.append([e])
    return bands


def _split_columns(band: List[Element]):
    items = sorted(band, key=lambda e: e.x0)
    cols = [[items[0]]]
    right = items[0].x1
    for e in items[1:]:
        if e.x0 > right + 10:
            cols.append([e])
        else:
            cols[-1].append(e)
        right = max(right, e.x1)
    return cols


class _OcrWriter:
    def __init__(self, doc: Document):
        self.doc = doc

    def add_page(self, img: np.ndarray, elements: List[Element], page_w_pt: float,
                 page_h_pt: float, first: bool):
        H, W = img.shape[:2]
        sx = page_w_pt / (W * PX2PT)
        doc = self.doc
        section = doc.sections[-1] if first else doc.add_section(WD_SECTION.NEW_PAGE)
        section.page_width = Pt(page_w_pt)
        section.page_height = Pt(page_h_pt)

        if elements:
            left = min(e.x0 for e in elements)
            right = max(e.x1 for e in elements)
            top = min(e.y0 for e in elements)
        else:
            left, right, top = W * 0.1, W * 0.9, H * 0.08
        ml = max(18.0, min(90.0, left * PX2PT * sx - 4))
        mr = max(18.0, min(90.0, (W - right) * PX2PT * sx - 4))
        mt = max(18.0, min(72.0, top * PX2PT * sx - 4))
        section.left_margin, section.right_margin = Pt(ml), Pt(mr)
        section.top_margin, section.bottom_margin = Pt(mt), Pt(min(54.0, mt))
        section.header_distance = section.footer_distance = Pt(12)

        content_left_px = ml / (PX2PT * sx)
        content_w_pt = page_w_pt - ml - mr
        self._k = PX2PT * sx  # px -> pt
        floats = [e for e in elements if e.kind == "image" and e.floating]
        flow = [e for e in elements if not (e.kind == "image" and e.floating)]
        body = doc.element.body
        start_len = len(body)
        cursor = mt
        for band in _bands(flow):
            cols = _split_columns(band) if len(band) > 1 else [band]
            if len(cols) == 1:
                for e in sorted(cols[0], key=lambda e: e.y0):
                    cursor = self._emit(doc, e, content_left_px, content_w_pt, cursor)
            else:
                cursor = self._emit_table(cols, content_left_px, content_w_pt, cursor)

        if floats:
            # jangkar gambar melayang di paragraf pertama halaman ini
            anchor_p = None
            for el in list(body)[max(0, start_len - 1):]:
                if el.tag == qn("w:p"):
                    anchor_p = el
                    break
            if anchor_p is None:
                anchor_p = doc.add_paragraph()._element
            from docx.text.paragraph import Paragraph
            par = Paragraph(anchor_p, doc._body)
            for e in floats:
                run = par.add_run()
                _add_png(run, e.image, e.w * self._k)
                _inline_to_anchor(run, e.x0 * self._k, e.y0 * self._k)

    # -- helpers --
    def _new_par(self, container, space_before):
        p = container.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(max(0.0, min(space_before, 400)))
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        return p

    def _metrics(self, e: Element):
        """(tinggi baris, total tinggi, offset atas) dalam pt untuk elemen."""
        k = self._k
        if e.kind == "image":
            return 0.0, e.h * k + 3.0, 0.0
        f = e.font_pt
        if len(e.lines) > 1:
            tops = [min(w.y0 for w in ln) for ln in e.lines]
            diffs = [b - a for a, b in zip(tops, tops[1:]) if b > a]
            pitch = statistics.median(diffs) * k if diffs else f * 1.22
            pitch = max(f * 1.05, min(pitch, f * 3))
            return pitch, pitch * len(e.lines), (pitch - f) * 0.6
        lh = f * 1.22
        return 0.0, lh, (lh - f) * 0.6

    def _emit(self, container, e: Element, left_px, avail_pt, cursor):
        """Tulis elemen; kembalikan posisi kursor vertikal (pt) setelahnya."""
        k = self._k
        pitch, total, off = self._metrics(e)
        target = e.y0 * k - off
        p = self._new_par(container, target - cursor)
        new_cursor = max(cursor, target) + total
        indent_pt = max(0.0, (e.x0 - left_px) * k)
        width_pt = e.w * k
        if e.kind == "image":
            width_pt = min(width_pt, avail_pt)
            center_off = indent_pt + width_pt / 2 - avail_pt / 2
            if abs(center_off) < avail_pt * 0.04:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.paragraph_format.left_indent = Pt(min(indent_pt, max(0.0, avail_pt - width_pt)))
            _add_png(p.add_run(), e.image, width_pt)
            return new_cursor

        # Paragraf teks
        if pitch:
            p.paragraph_format.line_spacing = Pt(pitch)  # jarak baris sesuai aslinya
        center = indent_pt + width_pt / 2
        multi = len(e.lines) > 1
        if not multi and abs(center - avail_pt / 2) < avail_pt * 0.05 and width_pt < avail_pt * 0.8 \
                and indent_pt > avail_pt * 0.08:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            indent_pt = 0
        else:
            p.paragraph_format.left_indent = Pt(min(indent_pt, avail_pt * 0.9))
        if multi:
            full = sum(1 for ln in e.lines[:-1] if ln[-1].x1 >= e.x1 - (e.x1 - e.x0) * 0.08)
            if full >= max(1, (len(e.lines) - 1) * 0.7) and width_pt > avail_pt * 0.6:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        self._fill_text(p, e, indent_pt, self._k / PX2PT, wide=width_pt > avail_pt * 0.55)
        return new_cursor

    def _fill_text(self, p, e: Element, indent_pt, sx, wide=True):
        run = p.add_run()
        run.font.size = Pt(e.font_pt)
        run.font.name = OCR_FONT
        run._element.rPr.rFonts.set(qn("w:eastAsia"), OCR_FONT)
        if e.bold:
            run.bold = True
        char_w = max(8.0, statistics.median([(w.x1 - w.x0) / max(1, len(w.text)) for ln in e.lines for w in ln]))
        tabs = set()
        para_w = e.x1 - e.x0
        out_parts = []
        for li, ln in enumerate(e.lines):
            line_txt = ""
            for wi, w in enumerate(ln):
                if wi == 0:
                    line_txt += w.text
                    continue
                gap = w.x0 - ln[wi - 1].x1
                if gap > char_w * 3.2:
                    pos = round((w.x0 - e.x0) * PX2PT * sx + indent_pt, 0)
                    tabs.add(pos)
                    line_txt += "\t" + w.text
                else:
                    line_txt += " " + w.text
            out_parts.append(line_txt)
            if li < len(e.lines) - 1:
                nxt = e.lines[li + 1]
                line_full = wide and ln[-1].x1 >= e.x0 + para_w * 0.85
                has_tab = "\t" in line_txt or any(
                    (nxt[i].x0 - nxt[i - 1].x1) > char_w * 3.2 for i in range(1, len(nxt)))
                if line_full and not has_tab:
                    if line_txt.endswith("-") and nxt[0].text[:1].islower():
                        out_parts[-1] = line_txt[:-1]
                        out_parts.append("")  # sambung tanpa spasi
                    else:
                        out_parts.append(" ")
                else:
                    out_parts.append("\n")
        text = "".join(out_parts)
        segments = text.split("\n")
        for si, seg in enumerate(segments):
            if si > 0:
                run.add_break(WD_BREAK.LINE)
            for ti, piece in enumerate(seg.split("\t")):
                if ti > 0:
                    run.add_tab()
                if piece:
                    run.add_text(piece)
        merged = []
        for t in sorted(tabs):
            if merged and t - merged[-1] < 8:
                merged[-1] = t  # satukan tab yang berdekatan agar kolom rata
            else:
                merged.append(t)
        for t in merged:
            try:
                p.paragraph_format.tab_stops.add_tab_stop(Pt(t))
            except Exception:
                pass

    def _emit_table(self, cols, left_px, avail_pt, cursor):
        k = self._k
        doc = self.doc
        band_top = min(e.y0 for col in cols for e in col) * k
        band_top -= max(self._metrics(e)[2] for col in cols for e in col)
        gap = band_top - cursor
        if gap > 1:
            spacer = self._new_par(doc, 0)
            spacer.paragraph_format.line_spacing = Pt(gap)
            spacer.add_run().font.size = Pt(1)
        start = max(cursor, band_top)
        # Batas kolom = titik tengah celah antar kolom
        bounds = [left_px]
        for a, b in zip(cols, cols[1:]):
            bounds.append((max(e.x1 for e in a) + min(e.x0 for e in b)) / 2)
        right_px = left_px + avail_pt / k
        bounds.append(max(right_px, max(e.x1 for e in cols[-1])))
        widths = [max(20.0, (bounds[i + 1] - bounds[i]) * k) for i in range(len(cols))]
        tot = sum(widths)
        widths = [w * avail_pt / tot for w in widths]

        table = doc.add_table(rows=1, cols=len(cols))
        table.autofit = False
        _set_cell_margins_zero(table)
        end = start
        for ci, col in enumerate(cols):
            cell = table.rows[0].cells[ci]
            cell.width = Pt(widths[ci])
            first_par = cell.paragraphs[0]
            c = start
            for e in sorted(col, key=lambda e: e.y0):
                c = self._emit(cell, e, bounds[ci], widths[ci], c)
            end = max(end, c)
            # hapus paragraf kosong bawaan sel
            first_par._element.getparent().remove(first_par._element)
        for ci, w in enumerate(widths):
            table.columns[ci].width = Pt(w)
        return end


def _image_page(doc: Document, page, first: bool):
    """Pilihan terakhir: halaman dimasukkan sebagai gambar utuh."""
    section = doc.sections[-1] if first else doc.add_section(WD_SECTION.NEW_PAGE)
    section.page_width = Pt(page.rect.width)
    section.page_height = Pt(page.rect.height)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Pt(18))
    pix = page.get_pixmap(dpi=200, alpha=False)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    maxw = page.rect.width - 36
    maxh = page.rect.height - 40
    scale = min(maxw / page.rect.width, maxh / page.rect.height)
    p.add_run().add_picture(io.BytesIO(pix.tobytes("png")), width=Pt(page.rect.width * scale))


def _new_ocr_doc() -> Document:
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = OCR_FONT
    st.element.rPr.rFonts.set(qn("w:eastAsia"), OCR_FONT)
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = 1.0
    # buang paragraf kosong awal
    body = doc.element.body
    for p in list(body.iterchildren(qn("w:p"))):
        body.remove(p)
    return doc


def ocr_pages_to_docx(pdf: "pymupdf.Document", page_numbers: List[int], out_path: str,
                      lang: str, progress: Optional[Callable[[int], None]] = None):
    doc = _new_ocr_doc()
    writer = _OcrWriter(doc)
    for i, pno in enumerate(page_numbers):
        page = pdf[pno]
        try:
            img = _render(page)
            img = _auto_rotate(img)
            img = _deskew(img)
            rotated = (img.shape[1] > img.shape[0]) != (page.rect.width > page.rect.height)
            pw, ph = (page.rect.height, page.rect.width) if rotated else (page.rect.width, page.rect.height)
            elements = _build_elements(img, lang)
            if not elements:
                raise ValueError("halaman kosong")
            writer.add_page(img, elements, pw, ph, first=(i == 0))
        except Exception as ex:  # pragma: no cover - jalur cadangan
            log.warning("OCR halaman %d gagal (%s), pakai gambar halaman", pno + 1, ex)
            _image_page(doc, page, first=(i == 0))
        if progress:
            progress(i + 1)
    doc.save(out_path)


# --------------------------------------------------------------------------
# Halaman digital via pdf2docx
# --------------------------------------------------------------------------

def digital_pages_to_docx(pdf_path: str, page_numbers: List[int], out_path: str):
    from pdf2docx import Converter
    logging.getLogger("pdf2docx").setLevel(logging.ERROR)
    cv = Converter(pdf_path)
    try:
        cv.convert(out_path, pages=page_numbers,
                   multi_processing=False,
                   # pertahankan gambar (ttd / stempel) dengan resolusi tinggi
                   clip_image_res_ratio=4.0)
    finally:
        cv.close()


# --------------------------------------------------------------------------
# Penggabungan
# --------------------------------------------------------------------------

def _merge_docx(parts: List[str], out_path: str):
    """Gabungkan beberapa DOCX dengan mempertahankan ukuran halaman & margin
    masing-masing bagian (tiap bagian = section sendiri, mulai di halaman baru)."""
    if len(parts) == 1:
        os.replace(parts[0], out_path)
        return
    from copy import deepcopy
    from docx.section import Section
    from docxcompose.composer import Composer
    master = Document(parts[0])
    comp = Composer(master)
    body = master.element.body
    for p in parts[1:]:
        d = Document(p)
        # tutup section terakhir master dengan pengaturannya sendiri
        last = body.sectPr
        para = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ppr.append(deepcopy(last))
        para.append(ppr)
        last.addprevious(para)
        comp.append(d)
        # section terakhir sekarang memakai pengaturan dokumen yang ditambahkan
        new = deepcopy(d.element.body.sectPr)
        body.sectPr.getparent().replace(body.sectPr, new)
        Section(new, master.part).start_type = WD_SECTION.NEW_PAGE
    comp.save(out_path)


# --------------------------------------------------------------------------
# API utama
# --------------------------------------------------------------------------

def convert_pdf_to_docx(pdf_path: str, out_path: str, mode: str = "auto",
                        password: str = "", lang: str = OCR_LANG,
                        progress: Optional[ProgressFn] = None) -> dict:
    """
    mode: "auto"   -> deteksi per halaman (disarankan)
          "ocr"    -> paksa OCR semua halaman (untuk scan / foto)
          "layout" -> paksa mode digital (pdf2docx) semua halaman
    """
    def report(frac, msg):
        if progress:
            progress(max(0.0, min(1.0, frac)), msg)

    try:
        pdf = pymupdf.open(pdf_path)
    except Exception as ex:
        raise ConversionError("File bukan PDF yang valid atau rusak.") from ex

    work = tempfile.mkdtemp(prefix="toword_")
    try:
        if pdf.needs_pass:
            if not password or not pdf.authenticate(password):
                raise PasswordRequired("PDF ini dikunci kata sandi. Masukkan kata sandinya.")
        # Simpan salinan tanpa enkripsi/batasan agar semua isi bisa dibaca
        src = os.path.join(work, "src.pdf")
        try:
            pdf.save(src, encryption=pymupdf.PDF_ENCRYPT_NONE, garbage=1, deflate=True)
        except Exception:
            pdf.save(src)
        pdf.close()
        pdf = pymupdf.open(src)

        n = pdf.page_count
        if n == 0:
            raise ConversionError("PDF tidak memiliki halaman.")

        report(0.03, "Menganalisis halaman…")
        kinds = []
        for i in range(n):
            if mode == "ocr":
                kinds.append("ocr")
            elif mode == "layout":
                kinds.append("digital")
            else:
                try:
                    kinds.append("ocr" if _page_needs_ocr(pdf[i]) else "digital")
                except Exception:
                    kinds.append("ocr")

        # Kelompokkan halaman berurutan dengan jenis yang sama
        segments = []
        for i, k in enumerate(kinds):
            if segments and segments[-1][0] == k:
                segments[-1][1].append(i)
            else:
                segments.append((k, [i]))

        done = 0
        parts = []
        for si, (kind, pages) in enumerate(segments):
            part = os.path.join(work, f"part{si}.docx")
            if kind == "digital":
                report(0.05 + 0.9 * done / n, f"Mengonversi halaman {pages[0] + 1}–{pages[-1] + 1}…")
                try:
                    digital_pages_to_docx(src, pages, part)
                except Exception as ex:
                    log.warning("pdf2docx gagal (%s), beralih ke OCR", ex)
                    kind = "ocr"
                    kinds = [("ocr" if i in pages else k) for i, k in enumerate(kinds)]
                else:
                    done += len(pages)
            if kind == "ocr":
                base = done

                def cb(k, base=base):
                    report(0.05 + 0.9 * (base + k) / n,
                           f"Membaca teks (OCR) halaman {min(n, base + k + 1)} dari {n}…")
                report(0.05 + 0.9 * done / n, f"Membaca teks (OCR) halaman {pages[0] + 1} dari {n}…")
                ocr_pages_to_docx(pdf, pages, part, lang, cb)
                done += len(pages)
            parts.append(part)

        report(0.96, "Menyusun dokumen Word…")
        _merge_docx(parts, out_path)
        report(1.0, "Selesai")
        return {
            "pages": n,
            "ocr_pages": sum(1 for k in kinds if k == "ocr"),
            "digital_pages": sum(1 for k in kinds if k == "digital"),
        }
    finally:
        try:
            pdf.close()
        except Exception:
            pass
        for f in os.listdir(work):
            try:
                os.remove(os.path.join(work, f))
            except Exception:
                pass
        try:
            os.rmdir(work)
        except Exception:
            pass
