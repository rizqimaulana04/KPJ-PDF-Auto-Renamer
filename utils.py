"""Pure processing helpers. Uploaded documents are handled only in memory."""
from __future__ import annotations

from collections import Counter
from io import BytesIO
import re
import unicodedata
import zipfile

import fitz
import pandas as pd
from pdfminer.high_level import extract_text

KPJ_RE = re.compile(r"(?<!\d)[0-9]{11}(?!\d)")
MAX_FILES = 2000
MAX_FILE_BYTES = 30 * 1024 * 1024
MAX_TOTAL_BYTES = 250 * 1024 * 1024
REPORT_COLUMNS = ["Nama File Awal", "KPJ Terdeteksi", "Status Matching", "Nama File Baru", "Keterangan"]


def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def detect_column(columns, kind):
    aliases = ({"kpj": 100, "nokpj": 100, "nomorkpj": 100, "nomorkpjbpjs": 95}
               if kind == "kpj" else {"namabaru": 100, "namafilebaru": 100,
               "filenamebaru": 100, "newfilename": 100, "namafile": 80, "filename": 80})
    scores = [aliases.get(normalize_header(c), 60 if
              ("kpj" in normalize_header(c) if kind == "kpj" else
               "baru" in normalize_header(c) and any(t in normalize_header(c) for t in ("nama", "file"))) else 0)
              for c in columns]
    return max(range(len(scores)), key=scores.__getitem__) if scores and max(scores) else None


def normalize_kpj(value):
    if value is None or pd.isna(value):
        return None
    value = re.sub(r"\s+", "", str(value)).strip("\"'")
    value = re.sub(r"\.0+$", "", value)
    return value if re.fullmatch(r"[0-9]{11}", value) else None


def read_workbook(data, filename, sheet=None):
    engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
    with pd.ExcelFile(BytesIO(data), engine=engine) as book:
        if sheet is None:
            return book.sheet_names
        return pd.read_excel(book, sheet_name=sheet, dtype=str, keep_default_na=False)


def safe_filename(value):
    value = unicodedata.normalize("NFC", str(value)).strip()
    value = re.sub(r'[\\/:*?"<>|\x00-\x1f\x7f]', "_", value)
    value = value.replace("..", "_").strip(" .")
    stem = value[:-4] if value.lower().endswith(".pdf") else value
    stem = stem.strip(" .") or "dokumen"
    if re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", stem, re.I):
        stem = "_" + stem
    while len(stem.encode("utf-8")) > 180:
        stem = stem[:-1]
    return stem + ".pdf"


def unique_path(folder, name, used):
    base = safe_filename(name)
    candidate = folder + "/" + base
    count = 1
    while candidate.casefold() in used:
        count += 1
        candidate = folder + "/" + base[:-4] + f"_{count}.pdf"
    used.add(candidate.casefold())
    return candidate, count > 1


def build_mapping(frame, kpj_column, name_column):
    mapping, issues, duplicate_keys, conflicts = {}, [], set(), set()
    for row_number, (_, row) in enumerate(frame.iterrows(), 2):
        raw_kpj, raw_name = row[kpj_column], str(row[name_column]).strip()
        kpj = normalize_kpj(raw_kpj)
        if not kpj or not raw_name:
            issues.append({"Baris Excel": row_number, "KPJ": str(raw_kpj), "Keterangan":
                           "KPJ tidak valid (harus 11 digit)" if not kpj else "Nama Baru kosong"})
            continue
        name = safe_filename(raw_name)
        if name != raw_name:
            issues.append({"Baris Excel": row_number, "KPJ": kpj, "Keterangan": f"Nama disanitasi: {name}"})
        if kpj in mapping:
            duplicate_keys.add(kpj)
            if mapping[kpj] != name:
                conflicts.add(kpj)
            issues.append({"Baris Excel": row_number, "KPJ": kpj, "Keterangan":
                           "KPJ duplikat dengan nama berbeda; perlu koreksi Excel" if kpj in conflicts else "KPJ duplikat dengan nama sama"})
        else:
            mapping[kpj] = name
    for kpj in conflicts:
        mapping.pop(kpj, None)
    counts = Counter(n.casefold() for n in mapping.values())
    for kpj, name in mapping.items():
        if counts[name.casefold()] > 1:
            issues.append({"Baris Excel": "", "KPJ": kpj, "Keterangan": f"Nama akhir digunakan beberapa KPJ: {name}"})
    return mapping, issues, duplicate_keys, conflicts


def ocr_pdf(data):
    """Optional hook: install pytesseract/Pillow and system Tesseract first."""
    import pytesseract
    from PIL import Image
    parts = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=200, alpha=False)
            parts.append(pytesseract.image_to_string(Image.open(BytesIO(pix.tobytes("png")))))
    return "\n".join(parts)


def read_pdf_kpjs(data, enable_ocr=False):
    notes, success = [], False
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            if doc.needs_pass:
                raise ValueError("PDF dilindungi password")
            text = "\n".join(page.get_text() for page in doc)
        success = True
        candidates = sorted(set(KPJ_RE.findall(text)))
        if candidates:
            return candidates, "PyMuPDF", False
        notes.append("PyMuPDF tidak menemukan KPJ")
    except Exception as exc:
        notes.append(f"PyMuPDF: {type(exc).__name__}: {exc}")
    try:
        text = extract_text(BytesIO(data))
        success = True
        candidates = sorted(set(KPJ_RE.findall(text)))
        if candidates:
            return candidates, "; ".join(notes + ["Fallback pdfminer.six berhasil"]), False
    except Exception as exc:
        notes.append(f"pdfminer.six: {type(exc).__name__}: {exc}")
    if enable_ocr:
        try:
            candidates = sorted(set(KPJ_RE.findall(ocr_pdf(data))))
            success = True
            if candidates:
                return candidates, "; ".join(notes + ["OCR berhasil; verifikasi hasil"]), False
        except Exception as exc:
            notes.append(f"OCR: {type(exc).__name__}: {exc}")
    return [], "; ".join(notes + ["Tidak ada KPJ 11 digit; OCR nonaktif" if not enable_ocr else "KPJ tidak ditemukan"]), not success


def report_excel(rows, issues, stats):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(rows, columns=REPORT_COLUMNS).to_excel(writer, sheet_name="Hasil Mapping", index=False)
        pd.DataFrame(issues, columns=["Baris Excel", "KPJ", "Keterangan"]).to_excel(writer, sheet_name="Validasi Excel", index=False)
        pd.DataFrame(stats.items(), columns=["Statistik", "Jumlah"]).to_excel(writer, sheet_name="Ringkasan", index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cells in sheet.columns:
                sheet.column_dimensions[cells[0].column_letter].width = min(65, max(18, max(len(str(c.value or "")) for c in cells) + 2))
                for cell in cells:
                    if isinstance(cell.value, str):
                        # Force literal strings, including untrusted values beginning '='.
                        cell.data_type = "s"
    return output.getvalue()


def process_zip(data, mapping, issues, duplicate_keys, conflicts, progress=None):
    rows, used, seen = [], set(), Counter()
    stats = dict.fromkeys(["Total PDF", "KPJ berhasil dibaca", "Berhasil cocok dengan Excel", "Berhasil di-rename",
                          "KPJ tidak ditemukan di Excel", "Gagal membaca KPJ", "Duplikat"], 0)
    output = BytesIO()
    with zipfile.ZipFile(BytesIO(data)) as source:
        pdfs = [i for i in source.infolist() if not i.is_dir() and i.filename.lower().endswith(".pdf")]
        if not pdfs:
            raise ValueError("ZIP tidak berisi file PDF.")
        if len(pdfs) > MAX_FILES or sum(i.file_size for i in pdfs) > MAX_TOTAL_BYTES or any(i.file_size > MAX_FILE_BYTES for i in pdfs):
            raise ValueError("Batas pemrosesan: 2.000 PDF, 30 MB per PDF, total PDF 250 MB setelah dekompresi. Pecah ZIP menjadi beberapa batch.")
        if any(i.flag_bits & 1 for i in pdfs):
            raise ValueError("ZIP terenkripsi tidak didukung. Upload ZIP tanpa password.")
        stats["Total PDF"] = len(pdfs)
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as result:
            for folder in ("Hasil_Rename", "Tidak_Ditemukan", "Gagal_Membaca_KPJ"):
                result.writestr(folder + "/", b"")
            for index, info in enumerate(pdfs):
                # Never extract archive paths onto disk; read each ZipInfo independently.
                try:
                    payload = source.read(info)
                except Exception as exc:
                    raise ValueError(f"Entri ZIP tidak dapat dibaca: {info.filename}. Proses dibatalkan agar hasil tidak kehilangan PDF. {exc}") from exc
                candidates, note, failed = read_pdf_kpjs(payload)
                original = info.filename.replace("\\", "/").rsplit("/", 1)[-1]
                name, folder, duplicate = original, "Gagal_Membaca_KPJ", False
                if candidates:
                    stats["KPJ berhasil dibaca"] += 1
                    matched = [k for k in candidates if k in mapping or k in conflicts]
                    selected = matched[0] if len(matched) == 1 else (candidates[0] if len(candidates) == 1 else None)
                    if selected:
                        seen[selected] += 1
                        duplicate = seen[selected] > 1 or selected in duplicate_keys
                    if selected in conflicts or len(matched) > 1:
                        folder, status = "Tidak_Ditemukan", "DUPLIKAT"
                        duplicate = True
                        note += "; KPJ/nama ambigu, file asli dipertahankan; koreksi data diperlukan"
                    elif selected in mapping:
                        folder, name, status = "Hasil_Rename", mapping[selected], "BERHASIL"
                        stats["Berhasil cocok dengan Excel"] += 1
                        stats["Berhasil di-rename"] += 1
                    else:
                        folder, status = "Tidak_Ditemukan", "TIDAK ADA DI EXCEL"
                        stats["KPJ tidak ditemukan di Excel"] += 1
                        note += "; Tidak ada kandidat KPJ yang cocok dengan Excel"
                    if len(candidates) > 1:
                        note += "; Beberapa angka 11 digit terdeteksi"
                else:
                    status = "ERROR" if failed else "KPJ TIDAK TERBACA"
                    stats["Gagal membaca KPJ"] += 1
                path, collision = unique_path(folder, name, used)
                if safe_filename(name) != name:
                    note += "; Nama asli disanitasi agar aman"
                if duplicate or collision:
                    stats["Duplikat"] += 1
                    note += f"; Duplikat (status dasar: {status}); disimpan tanpa overwrite"
                    status = "DUPLIKAT"
                result.writestr(path, payload)
                rows.append(dict(zip(REPORT_COLUMNS, [info.filename, ", ".join(candidates), status, path, note])))
                if progress:
                    progress((index + 1) / len(pdfs))
            report = report_excel(rows, issues, stats)
            result.writestr("laporan_mapping.xlsx", report)
    return {"zip": output.getvalue(), "report": report, "rows": rows, "stats": stats}
