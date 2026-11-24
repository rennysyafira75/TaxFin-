import base64
import io
import json
import os
import re
from datetime import datetime

import docx
from groq import Groq
import pandas as pd
from PIL import Image
from pypdf import PdfReader
import streamlit as st

# ==========================================
# KONFIGURASI HALAMAN & PENYIMPANAN PROJECT
# ==========================================
st.set_page_config(
    page_title="TaxFin Agent",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

PROJECTS_DIR = "data_projects"
os.makedirs(PROJECTS_DIR, exist_ok=True)

REGIONAL_TAX_SCHEMES = {
    "Nasional (Skema Pajak Pusat)": {
        "PPN Umum (11.0%)": 11.0,
        "PPN Umum (12.0% - Skema Baru)": 12.0,
        "PPh Final UMKM (0.5%)": 0.5,
        "PPh Pasal 21 (Tarif Efektif Rata-Rata 5.0%)": 5.0,
        "PPh Pasal 22 Impor API (2.5%)": 2.5,
        "PPh Pasal 23 Jasa/Sewa/Royalti (2.0%)": 2.0,
        "PPh Pasal 4(2) Sewa Tanah/Bangunan (10.0%)": 10.0,
        "PPh Pasal 4(2) Jasa Konstruksi Pelaksana (1.75%)": 1.75,
        "PPh Pasal 4(2) Jasa Konstruksi Perencana/Pengawas (4.0%)": 4.0,
    },
    "DKI Jakarta": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (15.0%)": 15.0,
        "PBJT Jasa Tenaga Listrik (2.4%)": 2.4,
        "Pajak Reklame (25.0%)": 25.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "Jawa Barat": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "Jawa Tengah": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "DI Yogyakarta": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "Jawa Timur": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "Banten": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "Bali": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
    "Riau": {
        "PBJT Makanan & Minuman (10.0%)": 10.0,
        "PBJT Jasa Perhotelan (10.0%)": 10.0,
        "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
        "PBJT Jasa Parkir (10.0%)": 10.0,
        "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
        "Pajak Air Tanah (20.0%)": 20.0,
        "Pajak MBLB (20.0%)": 20.0,
        "BPHTB (5.0%)": 5.0,
    },
}

OTHER_38_PROVINCES = [
    "Aceh", "Sumatera Utara", "Sumatera Barat", "Jambi", "Sumatera Selatan",
    "Bengkulu", "Lampung", "Kepulauan Bangka Belitung", "Kepulauan Riau",
    "Nusa Tenggara Barat", "Nusa Tenggara Timur", "Kalimantan Barat",
    "Kalimantan Tengah", "Kalimantan Selatan", "Kalimantan Timur",
    "Kalimantan Utara", "Sulawesi Utara", "Sulawesi Tengah",
    "Sulawesi Selatan", "Sulawesi Tenggara", "Gorontalo", "Sulawesi Barat",
    "Maluku", "Maluku Utara", "Papua", "Papua Barat", "Papua Barat Daya",
    "Papua Pegunungan", "Papua Selatan", "Papua Tengah",
]

DEFAULT_PROVINCE_TAX_TEMPLATE = {
    "PBJT Makanan & Minuman (10.0%)": 10.0,
    "PBJT Jasa Perhotelan (10.0%)": 10.0,
    "PBJT Jasa Kesenian & Hiburan Umum (10.0%)": 10.0,
    "PBJT Jasa Parkir (10.0%)": 10.0,
    "PBJT Jasa Tenaga Listrik (3.0%)": 3.0,
    "Pajak Air Tanah (20.0%)": 20.0,
    "Pajak MBLB (20.0%)": 20.0,
    "BPHTB (5.0%)": 5.0,
}

for _prov in OTHER_38_PROVINCES:
    if _prov not in REGIONAL_TAX_SCHEMES:
        REGIONAL_TAX_SCHEMES[_prov] = DEFAULT_PROVINCE_TAX_TEMPLATE.copy()


def slugify_project_id(name: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_-]+", "_", name.strip().lower()).strip("_")
    return cleaned or "project_utama"


def default_project_data(project_id: str, project_name: str) -> dict:
    default_region = "DKI Jakarta"
    default_scheme = list(REGIONAL_TAX_SCHEMES[default_region].keys())[0]
    default_rate = REGIONAL_TAX_SCHEMES[default_region][default_scheme]
    return {
        "id": project_id,
        "name": project_name,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "region": default_region,
        "tax_scheme": default_scheme,
        "use_custom_rate": False,
        "custom_rate": default_rate,
        "use_manual_dpp": False,
        "manual_dpp": 0.0,
        "doc_name": "",
        "doc_summary": "",
        "doc_records": [],
        "corrections": [],
        "messages": [
            {
                "role": "assistant",
                "content": (
                    f"Halo! Project **{project_name}** aktif. Silakan unggah dokumen keuangan "
                    "(Excel, CSV, Word, PDF) atau kirim foto/*screenshot* & pertanyaan di chat. "
                    "Saya akan menghitung dan menganalisis sesuai aturan pajak yang berlaku, serta "
                    "menyampaikan secara jujur jika ada data atau regulasi spesifik yang belum tersedia."
                ),
            }
        ],
    }


def list_saved_projects() -> list[dict]:
    items = []
    for fname in sorted(os.listdir(PROJECTS_DIR)):
        if fname.endswith(".json"):
            fpath = os.path.join(PROJECTS_DIR, fname)
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                items.append({
                    "id": data.get("id", fname[:-5]),
                    "name": data.get("name", fname[:-5]),
                    "updated_at": data.get("updated_at", "-"),
                })
            except Exception:
                continue
    return items


def load_project(project_id: str, fallback_name: str = "Project Utama") -> dict:
    fpath = os.path.join(PROJECTS_DIR, f"{project_id}.json")
    if os.path.exists(fpath):
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                raw = json.load(f)
            base = default_project_data(project_id, raw.get("name", fallback_name))
            base.update(raw)
            return base
        except Exception:
            pass
    proj = default_project_data(project_id, fallback_name)
    save_project(proj)
    return proj


def save_project(proj: dict) -> None:
    proj["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    fpath = os.path.join(PROJECTS_DIR, f"{proj['id']}.json")
    with open(fpath, "w", encoding="utf-8") as f:
        json.dump(proj, f, ensure_ascii=False, indent=2)


def get_groq_api_key() -> str:
    try:
        if "GROQ_API_KEY" in st.secrets and st.secrets["GROQ_API_KEY"]:
            return str(st.secrets["GROQ_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("GROQ_API_KEY", "").strip()


def clean_numeric_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype(str)
        .str.replace("Rp", "", regex=False)
        .str.replace("rp", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    both_sep = cleaned.str.contains(r"\.") & cleaned.str.contains(",")
    cleaned = cleaned.where(
        ~both_sep,
        cleaned.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    )
    no_comma_dot_thousands = (~both_sep) & cleaned.str.match(r"^-?\d{1,3}(\.\d{3})+$")
    cleaned = cleaned.where(
        ~no_comma_dot_thousands,
        cleaned.str.replace(".", "", regex=False),
    )
    cleaned = cleaned.str.replace(",", "", regex=False)
    return pd.to_numeric(cleaned, errors="coerce")


def extract_numbers_from_text(text: str) -> list[float]:
    values = []
    for match in re.findall(r"(?:Rp\s*)?-?\d[\d.,]{2,}", text):
        num_s = re.sub(r"(?i)^rp\s*", "", match).strip()
        if "." in num_s and "," in num_s:
            num_s = num_s.replace(".", "").replace(",", ".")
        elif re.match(r"^-?\d{1,3}(\.\d{3})+$", num_s):
            num_s = num_s.replace(".", "")
        else:
            num_s = num_s.replace(",", "")
        try:
            val = float(num_s)
            if abs(val) >= 100:
                values.append(val)
        except ValueError:
            continue
    return values


def analyze_document_file(uploaded_file) -> tuple[str, list[dict], list[str], float | None]:
    """Mengembalikan (ringkasan_teks, records_untuk_grafik, daftar_koreksi, estimasi_dpp)."""
    name = uploaded_file.name
    ext = os.path.splitext(name)[1].lower()
    corrections = []
    records = []
    estimated_dpp = None
    summary_parts = [f"File: {name}"]

    if ext in [".xlsx", ".xls", ".csv"]:
        if ext == ".csv":
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        summary_parts.append(
            f"Baris: {len(df)}, Kolom: {len(df.columns)} ({', '.join(map(str, df.columns[:10]))})"
        )

        dup_count = int(df.duplicated().sum())
        if dup_count > 0:
            corrections.append(f"Ditemukan {dup_count} baris duplikat persis; periksa apakah transaksi ganda.")

        null_counts = df.isna().sum()
        high_null_cols = [f"{col} ({int(cnt)} kosong)" for col, cnt in null_counts.items() if cnt > 0]
        if high_null_cols:
            corrections.append("Ada nilai kosong pada kolom: " + ", ".join(high_null_cols[:5]) + ".")

        numeric_cols = []
        for col in df.columns:
            s_num = clean_numeric_series(df[col])
            valid_ratio = s_num.notna().mean()
            if valid_ratio >= 0.4 and s_num.notna().sum() > 0:
                df[col] = s_num
                numeric_cols.append(col)

        if numeric_cols:
            main_col = numeric_cols[0]
            priority_keywords = [
                "dpp", "dasar", "nominal", "nilai", "jumlah", "total",
                "amount", "omzet", "pendapatan", "biaya",
            ]
            for col in numeric_cols:
                if any(k in str(col).lower() for k in priority_keywords):
                    main_col = col
                    break

            neg_count = int((df[main_col] < 0).sum())
            zero_count = int((df[main_col] == 0).sum())
            if neg_count > 0:
                corrections.append(
                    f"Terdapat {neg_count} nilai negatif pada kolom '{main_col}' (cek retur/nota kredit/salah input)."
                )
            if zero_count > 0:
                corrections.append(f"Terdapat {zero_count} baris bernilai 0 pada kolom '{main_col}'.")

            total_val = float(df[main_col].sum(skipna=True))
            estimated_dpp = max(0.0, total_val)
            summary_parts.append(f"Kolom nominal utama '{main_col}': total = {total_val:,.2f}")

            label_col = None
            for col in df.columns:
                if col != main_col and df[col].dtype == object:
                    label_col = col
                    break

            for idx, row in df.head(30).iterrows():
                lbl = str(row[label_col])[:30] if label_col else f"Baris {idx + 1}"
                val = row[main_col]
                records.append({
                    "label": lbl if lbl.strip() else f"Baris {idx + 1}",
                    "value": float(val) if pd.notna(val) else 0.0,
                })
        else:
            corrections.append("Tidak ditemukan kolom angka nominal yang jelas untuk dihitung otomatis.")

        summary_parts.append("Cuplikan data:\n" + df.head(8).to_string(index=False))

    elif ext == ".docx":
        doc = docx.Document(uploaded_file)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        table_lines = []
        for tbl in doc.tables:
            for row in tbl.rows:
                table_lines.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
        full_text = "\n".join(paragraphs + table_lines)
        nums = extract_numbers_from_text(full_text)
        if nums:
            estimated_dpp = sum(nums)
            summary_parts.append(f"Angka nominal terdeteksi ({len(nums)} item), jumlah total kasar: {sum(nums):,.2f}")
            for i, n_val in enumerate(nums[:25]):
                records.append({"label": f"Item {i + 1}", "value": float(n_val)})
        else:
            corrections.append("Tidak ada angka nominal jelas yang berhasil diekstrak dari dokumen Word.")
        summary_parts.append("Isi ringkas:\n" + full_text[:1800])

    elif ext == ".pdf":
        reader = PdfReader(uploaded_file)
        pages_text = []
        for p_idx, page in enumerate(reader.pages[:15]):
            txt = page.extract_text() or ""
            if txt.strip():
                pages_text.append(f"[Hal {p_idx + 1}] {txt.strip()}")
        full_text = "\n".join(pages_text)
        if not full_text.strip():
            corrections.append(
                "Teks PDF kosong/hasil scan gambar; unggah screenshot halaman PDF agar bisa dibaca visual oleh AI."
            )
        else:
            nums = extract_numbers_from_text(full_text)
            if nums:
                estimated_dpp = sum(nums)
                summary_parts.append(f"Angka nominal terdeteksi ({len(nums)} item), jumlah total kasar: {sum(nums):,.2f}")
                for i, n_val in enumerate(nums[:25]):
                    records.append({"label": f"Item {i + 1}", "value": float(n_val)})
            else:
                corrections.append("Tidak ditemukan angka nominal yang jelas pada teks PDF.")
            summary_parts.append("Isi ringkas:\n" + full_text[:1800])

    else:
        corrections.append(f"Format file '{ext}' belum didukung untuk pembacaan tabel otomatis.")

    return "\n".join(summary_parts), records, corrections, estimated_dpp


def pil_image_to_data_url(uploaded_img) -> tuple[str, Image.Image]:
    pil_img = Image.open(uploaded_img)
    rgb_img = pil_img.convert("RGB")
    max_side = 1400
    w, h = rgb_img.size
    if max(w, h) > max_side:
        scale = max_side / max(w, h)
        rgb_img = rgb_img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    rgb_img.save(buf, format="JPEG", quality=85)
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    return f"data:image/jpeg;base64,{b64}", pil_img


def build_system_prompt(proj: dict, active_rate: float, dpp: float, tax_amount: float, net_amount: float) -> str:
    scheme_label = "Tarif Kustom Pengguna" if proj.get("use_custom_rate") else proj.get("tax_scheme", "-")
    corrections_txt = "; ".join(proj.get("corrections", [])) or "Tidak ada catatan koreksi otomatis."
    doc_txt = proj.get("doc_summary") or "Belum ada dokumen terstruktur yang diunggah."

    return f"""Anda adalah **TaxFin Agent**, asisten keuangan dan perpajakan yang teliti, jelas, dan patuh aturan.

PRINSIP UTAMA & KEJUJURAN DATA:
1. Selalu berpedoman pada peraturan perpajakan Indonesia yang berlaku (seperti UU HPP, UU HKPD, dan ketentuan pajak daerah/nasional terkait) serta lakukan perhitungan matematis langkah demi langkah secara akurat.
2. **WAJIB JUJUR JIKA TIDAK TAHU ATAU DATA KURANG**: Jika informasi spesifik tidak ada dalam dokumen/gambar/konteks yang diberikan, atau jika aturan daerah/ketentuan tertentu belum pasti atau bergantung pada perda/dokumen asli yang tidak tersedia, katakan dengan terus terang bahwa data atau kepastian aturan tersebut belum tersedia/tidak diketahui—JANGAN mengarang angka, mengarang isi dokumen, atau menebak aturan seolah-olah pasti.
3. Jika pengguna meminta simulasi atau perhitungan dengan asumsi tertentu, sebutkan asumsi tersebut secara transparan dan tunjukkan bahwa angka tersebut bersifat simulasi/estimasi.

KONTEKS PROJECT AKTIF SAAT INI:
- Nama Project: {proj.get("name")}
- Daerah: {proj.get("region")}
- Skema Tarif: {scheme_label} ({active_rate:.2f}%)
- DPP Simulasi Saat Ini: Rp {dpp:,.2f}
- Estimasi Pajak Saat Ini: Rp {tax_amount:,.2f}
- Nilai Bersih / Setelah Pajak: Rp {net_amount:,.2f}
- Catatan Koreksi / Anomali Data: {corrections_txt}
- Ringkasan Dokumen Terunggah:
{doc_txt[:2500]}
"""


def ask_groq_agent(
    api_key: str,
    proj: dict,
    user_prompt: str,
    active_rate: float,
    dpp: float,
    tax_amount: float,
    net_amount: float,
    image_data_url: str | None = None,
) -> str:
    client = Groq(api_key=api_key)
    sys_msg = {"role": "system", "content": build_system_prompt(proj, active_rate, dpp, tax_amount, net_amount)}

    history_msgs = []
    for m in proj.get("messages", [])[-10:]:
        if m.get("role") in ("user", "assistant") and m.get("content"):
            history_msgs.append({"role": m["role"], "content": m["content"]})

    if image_data_url:
        user_msg = {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": user_prompt
                    if user_prompt.strip()
                    else "Tolong baca gambar/screenshot ini, periksa apakah ada angka atau potensi koreksi, dan hitung/analisis pajaknya sesuai konteks.",
                },
                {"type": "image_url", "image_url": {"url": image_data_url}},
            ],
        }
        model_name = "meta-llama/llama-4-scout-17b-16e-instruct"
    else:
        user_msg = {"role": "user", "content": user_prompt}
        model_name = "llama-3.3-70b-versatile"

    resp = client.chat.completions.create(
        model=model_name,
        messages=[sys_msg] + history_msgs + [user_msg],
        temperature=0.2,
        max_tokens=1800,
    )
    return resp.choices[0].message.content or "Tidak ada respons teks yang diterima."


# ==========================================
# INISIALISASI STATE & PROJECT AKTIF
# ==========================================
saved_list = list_saved_projects()
if "active_project_id" not in st.session_state:
    if saved_list:
        st.session_state.active_project_id = saved_list[0]["id"]
    else:
        first_proj = load_project("project_utama", "Project Utama")
        st.session_state.active_project_id = first_proj["id"]

proj = load_project(st.session_state.active_project_id)

# ==========================================
# PANEL KIRI (SEMUA FITUR TERPUSAT)
# ==========================================
with st.sidebar:
    st.title("💼 TaxFin Agent")
    st.caption("Semua fitur project, data, grafik, dan simulasi pajak terpusat di sini.")

    api_key = get_groq_api_key()
    if api_key:
        st.success("Groq API aktif")
    else:
        st.warning("GROQ_API_KEY belum terdeteksi di Secrets.")

    st.divider()

    # 1. MANAJEMEN PROJECT (AGAR DATA TIDAK HILANG)
    st.subheader("1. Manajemen Project")
    saved_list = list_saved_projects()
    id_to_label = {p["id"]: f"{p['name']} ({p['updated_at']})" for p in saved_list}
    if proj["id"] not in id_to_label:
        id_to_label[proj["id"]] = f"{proj['name']} ({proj['updated_at']})"

    selected_pid = st.selectbox(
        "Pilih Project Tersimpan",
        options=list(id_to_label.keys()),
        index=list(id_to_label.keys()).index(proj["id"]),
        format_func=lambda k: id_to_label[k],
    )
    if selected_pid != proj["id"]:
        st.session_state.active_project_id = selected_pid
        st.rerun()

    new_proj_name = st.text_input("Nama Project Baru", placeholder="Contoh: Klien A - Nov 2026")
    if st.button("Buat & Buka Project Baru", use_container_width=True):
        if new_proj_name.strip():
            base_id = slugify_project_id(new_proj_name)
            candidate_id = base_id
            counter = 2
            while os.path.exists(os.path.join(PROJECTS_DIR, f"{candidate_id}.json")):
                candidate_id = f"{base_id}_{counter}"
                counter += 1
            created = default_project_data(candidate_id, new_proj_name.strip())
            save_project(created)
            st.session_state.active_project_id = candidate_id
            st.rerun()
        else:
            st.info("Masukkan nama project terlebih dahulu.")

    with st.expander("Backup / Restore JSON Project"):
        st.download_button(
            label="Unduh Backup Project (.json)",
            data=json.dumps(proj, ensure_ascii=False, indent=2),
            file_name=f"{proj['id']}_backup.json",
            mime="application/json",
            use_container_width=True,
        )
        uploaded_json = st.file_uploader("Pulihkan dari file .json", type=["json"], key="restore_json")
        if uploaded_json is not None:
            try:
                restored = json.load(uploaded_json)
                if isinstance(restored, dict) and "messages" in restored:
                    restored["id"] = proj["id"]
                    restored["name"] = restored.get("name", proj["name"])
                    save_project(restored)
                    st.success("Project berhasil dipulihkan.")
                    st.rerun()
            except Exception as e:
                st.error(f"Gagal membaca JSON: {e}")

    st.divider()

    # 2. UNGGAH DOKUMEN & GRAFIK INSTAN
    st.subheader("2. Dokumen & Grafik Data")
    uploaded_doc = st.file_uploader(
        "Masukkan File (Excel, CSV, Word, PDF)",
        type=["xlsx", "xls", "csv", "docx", "pdf"],
        key="sidebar_doc_uploader",
    )
    if uploaded_doc is not None and uploaded_doc.name != proj.get("doc_name"):
        try:
            summary, records, corrections, est_dpp = analyze_document_file(uploaded_doc)
            proj["doc_name"] = uploaded_doc.name
            proj["doc_summary"] = summary
            proj["doc_records"] = records
            proj["corrections"] = corrections
            if est_dpp is not None and not proj.get("use_manual_dpp"):
                proj["manual_dpp"] = round(est_dpp, 2)
            save_project(proj)
            st.success(f"Dokumen '{uploaded_doc.name}' diproses.")
        except Exception as e:
            st.error(f"Gagal memproses dokumen: {e}")

    if proj.get("doc_name"):
        st.caption(f"Dokumen aktif: **{proj['doc_name']}**")

    if proj.get("doc_records"):
        df_chart = pd.DataFrame(proj["doc_records"])
        st.bar_chart(df_chart.set_index("label")["value"], height=180)
    else:
        st.caption("Unggah file Excel/CSV/Word/PDF untuk melihat grafik data di sini.")

    st.divider()

    # 3. TARIF PAJAK SEMUA DAERAH + TARIF KUSTOM
    st.subheader("3. Tarif Pajak Daerah & Kustom")
    regions = list(REGIONAL_TAX_SCHEMES.keys())
    curr_region = proj.get("region", "DKI Jakarta")
    if curr_region not in regions:
        curr_region = "DKI Jakarta"

    selected_region = st.selectbox(
        "Pilih Daerah / Provinsi",
        options=regions,
        index=regions.index(curr_region),
    )
    if selected_region != proj.get("region"):
        proj["region"] = selected_region
        schemes_in_region = list(REGIONAL_TAX_SCHEMES[selected_region].keys())
        proj["tax_scheme"] = schemes_in_region[0]
        proj["custom_rate"] = REGIONAL_TAX_SCHEMES[selected_region][schemes_in_region[0]]
        save_project(proj)

    region_schemes = REGIONAL_TAX_SCHEMES[selected_region]
    scheme_keys = list(region_schemes.keys())
    curr_scheme = proj.get("tax_scheme", scheme_keys[0])
    if curr_scheme not in scheme_keys:
        curr_scheme = scheme_keys[0]

    selected_scheme = st.selectbox(
        "Skema Tarif di Daerah Ini",
        options=scheme_keys,
        index=scheme_keys.index(curr_scheme),
    )
    default_scheme_rate = float(region_schemes[selected_scheme])

    if selected_scheme != proj.get("tax_scheme"):
        proj["tax_scheme"] = selected_scheme
        if not proj.get("use_custom_rate"):
            proj["custom_rate"] = default_scheme_rate
        save_project(proj)

    st.caption("Daftar tarif daerah:")
    st.dataframe(
        pd.DataFrame([
            {"Jenis Pajak": k, "Tarif (%)": v} for k, v in region_schemes.items()
        ]),
        use_container_width=True,
        hide_index=True,
        height=165,
    )

    use_custom = st.checkbox(
        "Sesuaikan Tarif Sendiri (%)",
        value=bool(proj.get("use_custom_rate", False)),
    )
    custom_val = st.number_input(
        "Tarif Kustom (%)",
        min_value=0.0,
        max_value=100.0,
        value=float(proj.get("custom_rate", default_scheme_rate)),
        step=0.25,
        disabled=not use_custom,
    )
    if use_custom != proj.get("use_custom_rate") or custom_val != proj.get("custom_rate"):
        proj["use_custom_rate"] = use_custom
        proj["custom_rate"] = float(custom_val)
        save_project(proj)

    active_rate = float(custom_val) if use_custom else default_scheme_rate

    st.divider()

    # 4. SIMULASI PERHITUNGAN PAJAK & DETEKSI KOREKSI
    st.subheader("4. Simulasi & Deteksi Koreksi")
    use_manual_dpp = st.checkbox(
        "Atur Nominal DPP Manual",
        value=bool(proj.get("use_manual_dpp", False)),
    )
    dpp_input = st.number_input(
        "Dasar Pengenaan Pajak / DPP (Rp)",
        min_value=0.0,
        value=float(proj.get("manual_dpp", 0.0)),
        step=100000.0,
        disabled=not use_manual_dpp and bool(proj.get("doc_records")),
    )
    if use_manual_dpp != proj.get("use_manual_dpp") or dpp_input != proj.get("manual_dpp"):
        proj["use_manual_dpp"] = use_manual_dpp
        proj["manual_dpp"] = float(dpp_input)
        save_project(proj)

    dpp = float(proj.get("manual_dpp", 0.0))
    tax_amount = dpp * (active_rate / 100.0)
    net_amount = dpp - tax_amount

    m1, m2 = st.columns(2)
    m1.metric("Tarif Aktif", f"{active_rate:.2f}%")
    m2.metric("Estimasi Pajak", f"Rp {tax_amount:,.0f}")
    st.metric("Nilai Bersih (DPP - Pajak)", f"Rp {net_amount:,.0f}")

    st.markdown("**Deteksi Koreksi / Anomali Data:**")
    if proj.get("corrections"):
        for corr in proj["corrections"]:
            st.warning(corr)
    else:
        st.info("Belum ada catatan koreksi. Unggah file atau tanyakan langsung ke AI di chat.")

    if st.button("Kirim Data & Koreksi ke AI untuk Dianalisis", use_container_width=True):
        if not api_key:
            st.error("GROQ_API_KEY belum tersedia di Secrets.")
        else:
            auto_q = (
                f"Tolong analisis simulasi pajak untuk project '{proj['name']}' "
                f"(Daerah: {selected_region}, Tarif aktif: {active_rate:.2f}%, "
                f"DPP: Rp {dpp:,.0f}, Estimasi Pajak: Rp {tax_amount:,.0f}, Nilai Bersih: Rp {net_amount:,.0f}) "
                f"serta periksa catatan koreksi/anomali yang ada. Jelaskan sesuai aturan dan jujur bila ada data yang belum lengkap."
            )
            proj["messages"].append({"role": "user", "content": auto_q})
            try:
                reply = ask_groq_agent(api_key, proj, auto_q, active_rate, dpp, tax_amount, net_amount)
                proj["messages"].append({"role": "assistant", "content": reply})
                save_project(proj)
                st.rerun()
            except Exception as e:
                st.error(f"Gagal memanggil AI: {e}")


# ==========================================
# AREA UTAMA: RUANG CHAT AI + FOTO / FILE
# ==========================================
st.header(f"💬 {proj['name']}")
st.caption(
    f"Daerah: **{selected_region}** • Tarif Aktif: **{active_rate:.2f}%** • "
    "Chat AI berpedoman pada aturan & jujur jika data/aturan spesifik tidak tersedia."
)

for msg in proj.get("messages", []):
    with st.chat_message(msg.get("role", "assistant")):
        if msg.get("image_b64"):
            st.image(base64.b64decode(msg["image_b64"]), caption="Lampiran Gambar / Screenshot", width=320)
        if msg.get("attached_file_name"):
            st.caption(f"📎 File terlampir: `{msg['attached_file_name']}`")
        st.markdown(msg.get("content", ""))

with st.expander("📎 Lampirkan Screenshot / Foto atau File ke Pesan Berikutnya", expanded=False):
    c_img, c_file = st.columns(2)
    with c_img:
        chat_img = st.file_uploader(
            "Foto / Screenshot (.png, .jpg, .jpeg, .webp)",
            type=["png", "jpg", "jpeg", "webp"],
            key="chat_image_uploader",
        )
    with c_file:
        chat_extra_file = st.file_uploader(
            "File Tambahan (.xlsx, .csv, .docx, .pdf)",
            type=["xlsx", "xls", "csv", "docx", "pdf"],
            key="chat_extra_file_uploader",
        )

prompt_text = st.chat_input("Tulis pertanyaan, minta perhitungan pajak, atau minta analisis data di sini...")

if prompt_text is not None:
    if not api_key:
        st.error("GROQ_API_KEY belum terdeteksi di Secrets Streamlit Cloud.")
    else:
        extra_note = ""
        attached_name = None
        image_data_url = None
        image_b64_save = None

        if chat_extra_file is not None:
            attached_name = chat_extra_file.name
            try:
                f_sum, f_recs, f_corrs, f_dpp = analyze_document_file(chat_extra_file)
                proj["doc_name"] = chat_extra_file.name
                proj["doc_summary"] = f_sum
                if f_recs:
                    proj["doc_records"] = f_recs
                proj["corrections"] = f_corrs
                if f_dpp is not None and not proj.get("use_manual_dpp"):
                    proj["manual_dpp"] = round(f_dpp, 2)
                    dpp = float(proj["manual_dpp"])
                    tax_amount = dpp * (active_rate / 100.0)
                    net_amount = dpp - tax_amount
                extra_note += f"\n\n[Konteks File Terlampir '{attached_name}']:\n{f_sum[:1800]}"
            except Exception as e:
                extra_note += f"\n\n[Catatan File '{attached_name}']: Gagal dibaca ({e})."

        if chat_img is not None:
            try:
                image_data_url, pil_img = pil_image_to_data_url(chat_img)
                buf_thumb = io.BytesIO()
                pil_img.convert("RGB").save(buf_thumb, format="JPEG", quality=80)
                image_b64_save = base64.b64encode(buf_thumb.getvalue()).decode("utf-8")
            except Exception as e:
                extra_note += f"\n\n[Catatan Gambar]: Gagal memproses gambar ({e})."

        user_record = {
            "role": "user",
            "content": prompt_text if prompt_text.strip() else "Tolong analisis lampiran ini.",
        }
        if attached_name:
            user_record["attached_file_name"] = attached_name
        if image_b64_save:
            user_record["image_b64"] = image_b64_save

        proj["messages"].append(user_record)

        with st.chat_message("user"):
            if image_b64_save:
                st.image(base64.b64decode(image_b64_save), caption="Lampiran Gambar / Screenshot", width=320)
            if attached_name:
                st.caption(f"📎 File terlampir: `{attached_name}`")
            st.markdown(user_record["content"])

        with st.chat_message("assistant"):
            with st.spinner("Sedang menganalisis..."):
                try:
                    ai_reply = ask_groq_agent(
                        api_key=api_key,
                        proj=proj,
                        user_prompt=(prompt_text + extra_note).strip(),
                        active_rate=active_rate,
                        dpp=dpp,
                        tax_amount=tax_amount,
                        net_amount=net_amount,
                        image_data_url=image_data_url,
                    )
                except Exception as e:
                    ai_reply = f"Maaf, terjadi kendala saat memproses permintaan ke AI: `{e}`"
            st.markdown(ai_reply)

        proj["messages"].append({"role": "assistant", "content": ai_reply})
        save_project(proj)
        st.rerun()
