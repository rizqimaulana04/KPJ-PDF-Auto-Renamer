from hashlib import sha256
import streamlit as st
import pandas as pd
from utils import read_workbook, detect_column, build_mapping, process_zip

st.set_page_config(page_title="KPJ PDF Auto Renamer", page_icon="📄", layout="wide")
st.title("KPJ PDF Auto Renamer")
st.caption("Rename file PDF otomatis berdasarkan Nomor KPJ dan data Excel.")
st.info("Upload dokumen, periksa kolom Excel, lalu proses. PDF asli tetap disertakan jika tidak cocok atau gagal dibaca.")
if st.button("Bersihkan sesi", help="Hapus input dan hasil dari sesi aplikasi ini."):
    next_generation = st.session_state.get("upload_generation", 0) + 1
    for key in list(st.session_state):
        del st.session_state[key]
    st.session_state["upload_generation"] = next_generation
    st.rerun()

generation = st.session_state.get("upload_generation", 0)
left, right = st.columns(2)
with left:
    zip_upload = st.file_uploader("1. Upload ZIP", type=["zip"], key=f"zip_input_{generation}")
with right:
    excel_upload = st.file_uploader("2. Upload Excel", type=["xlsx", "xls"], key=f"excel_input_{generation}")
st.caption("Maks. upload 100 MB/file • 2.000 PDF/batch • 30 MB/PDF • total PDF setelah dekompresi 250 MB. PDF scan memerlukan OCR (nonaktif).")

if excel_upload is None:
    st.session_state.pop("result", None)
    st.stop()
excel_bytes = excel_upload.getvalue()
excel_id = sha256(excel_bytes).hexdigest()
try:
    sheets = read_workbook(excel_bytes, excel_upload.name)
    sheet = st.selectbox("3. Pilih Sheet Excel", sheets, key=f"sheet_{excel_id}")
    frame = read_workbook(excel_bytes, excel_upload.name, sheet)
except Exception as exc:
    st.error(f"Excel tidak dapat dibaca: {exc}")
    st.stop()
st.subheader("4. Preview Excel")
st.dataframe(frame.head(5), hide_index=True, use_container_width=True)
columns = list(frame.columns)
if len(columns) < 2 or frame.empty:
    st.warning("Excel harus berisi data dengan minimal dua kolom dan header pada baris pertama.")
    st.stop()
kpj_index, name_index = detect_column(columns, "kpj"), detect_column(columns, "name")
kpj_col = st.selectbox("5. Kolom Nomor KPJ", columns, index=kpj_index, placeholder="Pilih kolom KPJ", key=f"kpj_{excel_id}_{sheet}")
name_col = st.selectbox("6. Kolom Nama File Baru", columns, index=name_index, placeholder="Pilih kolom Nama Baru", key=f"name_{excel_id}_{sheet}")
if kpj_col is None or name_col is None or kpj_col == name_col:
    st.warning("Pilih dua kolom yang berbeda sebelum memproses.")
    st.stop()
mapping, issues, duplicates, conflicts = build_mapping(frame, kpj_col, name_col)
st.caption(f"{len(mapping):,} KPJ valid • {len(duplicates):,} KPJ duplikat • {len(conflicts):,} KPJ dengan nama bertentangan")
if issues:
    st.warning("Ada catatan validasi Excel. KPJ dengan nama bertentangan tidak akan di-rename otomatis.")
    with st.expander("Lihat catatan validasi Excel"):
        st.dataframe(pd.DataFrame(issues), hide_index=True, use_container_width=True)
if zip_upload is None:
    st.session_state.pop("result", None)
    st.stop()
zip_bytes = zip_upload.getvalue()
signature = (sha256(zip_bytes).hexdigest(), excel_id, sheet, kpj_col, name_col)
if st.session_state.get("signature") != signature:
    st.session_state.pop("result", None)
if st.button("Proses File", type="primary", disabled=not mapping):
    st.session_state.pop("result", None)
    bar = st.progress(0, text="Membaca PDF dan mencocokkan KPJ…")
    try:
        result = process_zip(zip_bytes, mapping, issues, duplicates, conflicts, bar.progress)
        st.session_state["result"] = result
        st.session_state["signature"] = signature
        bar.progress(1.0, text="Proses selesai")
    except Exception as exc:
        bar.empty()
        st.error(f"Proses gagal: {exc}. Input asli tidak diubah; hasil parsial tidak diterbitkan.")

if "result" in st.session_state:
    result = st.session_state["result"]
    st.success("Proses selesai. Download hasil sebelum membersihkan atau meninggalkan sesi.")
    for start in (0, 4):
        items = list(result["stats"].items())[start:start + 4]
        for column, (label, value) in zip(st.columns(len(items)), items):
            column.metric(label, value)
    st.caption("Duplikat menghitung PDF yang terdampak KPJ ganda, konflik, atau benturan nama; dapat tumpang tindih dengan statistik lainnya.")
    st.dataframe(pd.DataFrame(result["rows"]), hide_index=True, use_container_width=True)
    first, second = st.columns(2)
    first.download_button("Download ZIP Hasil", result["zip"], "hasil_kpj.zip", "application/zip", use_container_width=True)
    second.download_button("Download Laporan Excel", result["report"], "laporan_mapping.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
