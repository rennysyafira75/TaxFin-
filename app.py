import io
import os
import re
import pandas as pd
import streamlit as st
from docx import Document
from groq import Groq
from pypdf import PdfReader

st.set_page_config(
    page_title="TaxFin Agent",
    page_icon="📊",
    layout="wide",
)

TAX_REGIONS = {
    "DKI Jakarta (Skema Standar 11%)": {
        "region_code": "ID-JK",
        "rate": 0.11,
        "notes": "Tarif simulasi standar 11% dari nilai dasar kena pajak (DPP).",
    },
    "Riau / Pekanbaru (Skema Standar 11%)": {
        "region_code": "ID-RI",
        "rate": 0.11,
        "notes": "Tarif simulasi standar 11% untuk evaluasi entitas daerah Riau.",
    },
    "Jawa Barat (Skema Insentif / Khusus 9.5%)": {
        "region_code": "ID-JB",
        "rate": 0.095,
        "notes": "Tarif simulasi insentif/khusus 9.5% untuk perbandingan skenario.",
    },
    "Bali (Skema Pariwisata & Jasa 10%)": {
        "region_code": "ID-BA",
        "rate": 0.10,
        "notes": "Tarif simulasi 10% untuk skenario sektor jasa/usaha tertentu.",
    },
}

if "entity_records" not in st.session_state:
    st.session_state.entity_records = {}


def get_entity_store(entity_key: str) -> dict:
    if entity_key not in st.session_state.entity_records:
        st.session_state.entity_records[entity_key] = {
            "history": [],
            "last_ai_output": "",
            "last_qa_output": "",
        }
    return st.session_state.entity_records[entity_key]


def get_groq_api_key() -> str | None:
    try:
        if "GROQ_API_KEY" in st.secrets and st.secrets["GROQ_API_KEY"]:
            return str(st.secrets["GROQ_API_KEY"]).strip()
    except Exception:
        pass
    env_key = os.getenv("GROQ_API_KEY")
    return env_key.strip() if env_key else None


def parse_excel_or_csv(uploaded_file) -> dict:
    filename = uploaded_file.name.lower()
    raw_bytes = uploaded_file.getvalue()

    if filename.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw_bytes))
    else:
        df = pd.read_excel(io.BytesIO(raw_bytes))

    df = df.dropna(how="all")

    numeric_cols = []
    for col in df.columns:
        if df[col].dtype == object:
            cleaned = (
                df[col]
                .astype(str)
                .str.replace("Rp", "", regex=False)
                .str.replace("rp", "", regex=False)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.strip()
            )
            converted = pd.to_numeric(cleaned, errors="coerce")
            if converted.notna().sum() >= max(1, int(len(df) * 0.4)):
                df[col] = converted

        if pd.api.types.is_numeric_dtype(df[col]):
            numeric_cols.append(col)

    return {
        "file_type": "excel/csv",
        "dataframe": df,
        "numeric_cols": numeric_cols,
        "text": df.head(30).to_string(index=False),
    }


def parse_word(uploaded_file) -> dict:
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    parts = []

    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text.strip())

    for idx, table in enumerate(doc.tables, start=1):
        parts.append(f"\n[Tabel Word #{idx}]")
        for row in table.rows:
            row_cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            parts.append(" | ".join(row_cells))

    full_text = "\n".join(parts).strip()
    return {
        "file_type": "word",
        "dataframe": None,
        "numeric_cols": [],
        "text": full_text,
    }


def parse_pdf(uploaded_file) -> dict:
    reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
    pages_text = []

    for page_num, page in enumerate(reader.pages, start=1):
        extracted = (page.extract_text() or "").strip()
        if extracted:
            pages_text.append(f"--- Halaman {page_num} ---\n{extracted}")

    full_text = "\n\n".join(pages_text).strip()
    return {
        "file_type": "pdf",
        "dataframe": None,
        "numeric_cols": [],
        "text": full_text,
    }


def extract_numbers_from_text(text: str) -> list[float]:
    if not text:
        return []
    matches = re.findall(r"-?\d[\d.,]*", text)
    values = []
    for m in matches:
        token = m.strip()
        if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", token):
            token = token.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(,\d{3})+(\.\d+)?", token):
            token = token.replace(",", "")
        else:
            token = token.replace(",", ".")
        try:
            val = float(token)
            if abs(val) >= 100:
                values.append(val)
        except ValueError:
            continue
    return values


def run_tax_simulation(
    parsed_data: dict | None,
    selected_region: str,
    selected_column: str | None = None,
    manual_base_amount: float | None = None,
) -> dict:
    region_info = TAX_REGIONS[selected_region]
    rate = region_info["rate"]
    corrections = []
    df = parsed_data["dataframe"] if parsed_data else None

    if manual_base_amount is not None and manual_base_amount > 0:
        base_amount = float(manual_base_amount)
        source_desc = "Input Manual Pengguna"
    elif df is not None and not df.empty and selected_column and selected_column in df.columns:
        series = pd.to_numeric(df[selected_column], errors="coerce")
        valid_positive = series[series > 0]
        base_amount = float(valid_positive.sum()) if not valid_positive.empty else 0.0
        source_desc = f"Kolom Excel/CSV: '{selected_column}'"
    elif parsed_data and parsed_data.get("text"):
        nums = extract_numbers_from_text(parsed_data["text"])
        positive_nums = [n for n in nums if n > 0]
        base_amount = float(sum(positive_nums[:20])) if positive_nums else 0.0
        source_desc = f"Ekstraksi Teks ({parsed_data.get('file_type', 'dokumen')})"
    else:
        base_amount = 0.0
        source_desc = "Belum ada data angka"

    if df is not None and not df.empty:
        dup_count = int(df.duplicated().sum())
        if dup_count > 0:
            corrections.append({
                "level": "Peringatan",
                "issue": f"Ditemukan {dup_count} baris duplikat identik pada tabel.",
                "action": "Periksa apakah transaksi tersebut memang berulang atau tidak sengaja terinput dua kali.",
            })

        for col in parsed_data.get("numeric_cols", []):
            s = pd.to_numeric(df[col], errors="coerce")
            neg_count = int((s < 0).sum())
            nan_count = int(s.isna().sum())
            zero_count = int((s == 0).sum())

            if neg_count > 0:
                corrections.append({
                    "level": "Koreksi",
                    "issue": f"Kolom '{col}' memiliki {neg_count} nilai negatif.",
                    "action": "Pastikan nilai minus merupakan retur/potongan sah atau salah tanda minus.",
                })
            if nan_count > 0:
                corrections.append({
                    "level": "Koreksi",
                    "issue": f"Kolom '{col}' memiliki {nan_count} sel kosong/tidak terbaca sebagai angka.",
                    "action": "Lengkapi atau rapikan format angka pada sel yang kosong/teks.",
                })
            if zero_count > 0:
                corrections.append({
                    "level": "Info",
                    "issue": f"Kolom '{col}' memiliki {zero_count} nilai nol (0).",
                    "action": "Cek apakah nilai nol memang wajar atau belum terisi.",
                })

    if parsed_data and parsed_data.get("text"):
        neg_text_nums = [n for n in extract_numbers_from_text(parsed_data["text"]) if n < 0]
        if neg_text_nums:
            corrections.append({
                "level": "Koreksi",
                "issue": f"Terdeteksi {len(neg_text_nums)} angka bernilai negatif di dalam teks dokumen.",
                "action": "Verifikasi apakah angka negatif tersebut pengurang/koreksi sah atau salah ketik.",
            })

    if base_amount <= 0:
        corrections.append({
            "level": "Koreksi",
            "issue": "Dasar pengenaan pajak (DPP) masih 0 atau belum terdeteksi.",
            "action": "Pilih kolom angka yang sesuai atau isi Nominal Dasar Simulasi Manual.",
        })

    estimated_tax = base_amount * rate
    net_amount = base_amount - estimated_tax

    comparison_rows = []
    for region_name, info in TAX_REGIONS.items():
        r = info["rate"]
        tax_val = base_amount * r
        comparison_rows.append({
            "Daerah / Skema": region_name,
            "Kode": info["region_code"],
            "Tarif (%)": round(r * 100, 2),
            "Dasar Pajak (DPP)": round(base_amount, 2),
            "Estimasi Pajak": round(tax_val, 2),
            "Nilai Bersih": round(base_amount - tax_val, 2),
        })

    return {
        "region": selected_region,
        "rate": rate,
        "source_desc": source_desc,
        "base_amount": base_amount,
        "estimated_tax": estimated_tax,
        "net_amount": net_amount,
        "corrections": corrections,
        "comparison_df": pd.DataFrame(comparison_rows),
    }


def analyze_with_groq(
    entity_name: str,
    period_label: str,
    sim_result: dict,
    document_excerpt: str,
    user_question: str | None = None,
) -> str:
    api_key = get_groq_api_key()
    if not api_key:
        return (
            "GROQ_API_KEY belum ditemukan. Pastikan `GROQ_API_KEY` sudah diisi di "
            "pengaturan **Secrets** aplikasi sebelum meminta analisis AI."
        )

    client = Groq(api_key=api_key)

    corrections_text = (
        "\n".join(
            f"- [{c['level']}] {c['issue']} -> Saran: {c['action']}"
            for c in sim_result.get("corrections", [])
        )
        if sim_result.get("corrections")
        else "- Tidak ada anomali/koreksi kritis terdeteksi secara otomatis."
    )

    system_prompt = (
        "Kamu adalah TaxFin Agent, asisten analisis keuangan dan simulasi perhitungan pajak. "
        "Jawablah dalam Bahasa Indonesia yang jelas, terstruktur, dan ringkas. "
        "Selalu hormati batas isolasi entitas dan periode yang sedang aktif. "
        "Ingatkan secara wajar bahwa hasil simulasi ini bersifat bantu analisis/simulasi dan "
        "perlu diverifikasi terhadap aturan pajak resmi yang berlaku."
    )

    context_block = f"""
Konteks Entitas: {entity_name or "Entitas Umum"}
Periode Laporan: {period_label or "Tidak ditentukan"}
Daerah / Skema Tarif Aktif: {sim_result["region"]} (Tarif: {sim_result["rate"] * 100:.2f}%)
Sumber Nilai Dasar: {sim_result["source_desc"]}
Dasar Pengenaan Pajak (DPP): {sim_result["base_amount"]:,.2f}
Estimasi Pajak Terutang: {sim_result["estimated_tax"]:,.2f}
Nilai Bersih Setelah Pajak: {sim_result["net_amount"]:,.2f}

Temuan Deteksi Koreksi Otomatis:
{corrections_text}

Cuplikan Isi Dokumen (terbatas):
{document_excerpt[:2500] if document_excerpt else "(Tidak ada teks dokumen)"}
""".strip()

    if user_question and user_question.strip():
        final_prompt = (
            f"{context_block}\n\n"
            f"Pertanyaan Pengguna:\n{user_question.strip()}\n\n"
            "Jawab pertanyaan berdasarkan konteks data di atas secara akurat dan praktis."
        )
    else:
        final_prompt = (
            f"{context_block}\n\n"
            "Berikan:\n"
            "1. Ringkasan singkat kondisi keuangan & simulasi pajak.\n"
            "2. Evaluasi atas temuan koreksi/anomali data.\n"
            "3. Rekomendasi langkah pemeriksaan berikutnya."
        )

    completion = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": final_prompt},
        ],
        temperature=0.2,
        max_tokens=1200,
    )

    return completion.choices[0].message.content


st.title("TaxFin Agent")
st.caption(
    "Asisten analisis keuangan & simulasi perhitungan pajak adaptif multi-daerah "
    "(PDF, Word, Excel) dengan deteksi koreksi data dan dukungan Groq AI."
)

with st.sidebar:
    st.header("Konfigurasi & Isolasi Entitas")
    entity_name = st.text_input("Nama Entitas / Unit", value="Entitas Utama")
    period_label = st.text_input("Periode Laporan", value="November 2026")
    entity_key = f"{entity_name.strip() or 'Entitas Utama'}::{period_label.strip() or 'Umum'}"

    selected_region = st.selectbox(
        "Pilih Daerah / Skema Tarif Pajak",
        options=list(TAX_REGIONS.keys()),
    )
    st.info(TAX_REGIONS[selected_region]["notes"])

    st.divider()
    st.subheader("Status Groq API")
    if get_groq_api_key():
        st.success("GROQ_API_KEY terdeteksi.")
    else:
        st.warning("GROQ_API_KEY belum terpasang di Secrets.")

    store = get_entity_store(entity_key)
    if st.button("Bersihkan Riwayat Entitas Ini", use_container_width=True):
        store["history"] = []
        store["last_ai_output"] = ""
        store["last_qa_output"] = ""
        st.rerun()

st.header("1. Unggah & Pratinjau Dokumen")
uploaded_file = st.file_uploader(
    "Unggah file laporan keuangan / transaksi (.xlsx, .csv, .docx, .pdf)",
    type=["xlsx", "csv", "docx", "pdf"],
)

parsed_data = None
selected_column = None

if uploaded_file is not None:
    try:
        lower_name = uploaded_file.name.lower()
        if lower_name.endswith((".xlsx", ".csv")):
            parsed_data = parse_excel_or_csv(uploaded_file)
        elif lower_name.endswith(".docx"):
            parsed_data = parse_word(uploaded_file)
        elif lower_name.endswith(".pdf"):
            parsed_data = parse_pdf(uploaded_file)

        st.success(f"Berhasil membaca file: **{uploaded_file.name}**")

        if parsed_data["dataframe"] is not None:
            st.dataframe(parsed_data["dataframe"].head(25), use_container_width=True)
            if parsed_data["numeric_cols"]:
                selected_column = st.selectbox(
                    "Pilih Kolom Angka untuk Dasar Pengenaan Pajak (DPP)",
                    options=parsed_data["numeric_cols"],
                )
            else:
                st.warning("Tidak ada kolom numerik otomatis terdeteksi pada file tabel ini.")
        else:
            st.text_area(
                "Hasil Ekstraksi Teks Dokumen",
                value=parsed_data["text"][:4000] if parsed_data["text"] else "Teks kosong.",
                height=200,
                disabled=True,
            )
    except Exception as exc:
        st.error(f"Gagal membaca dokumen: {exc}")

st.header("2. Simulasi Perhitungan Pajak & Deteksi Koreksi")
use_manual = st.checkbox("Gunakan Nominal Dasar Pengenaan Pajak (DPP) Manual")
manual_base = None
if use_manual:
    manual_base = st.number_input(
        "Masukkan Nominal Dasar Pajak Manual",
        min_value=0.0,
        step=100000.0,
        format="%.2f",
    )

sim_result = run_tax_simulation(
    parsed_data=parsed_data,
    selected_region=selected_region,
    selected_column=selected_column,
    manual_base_amount=manual_base if use_manual else None,
)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Dasar Pajak (DPP)", f"{sim_result['base_amount']:,.2f}")
col2.metric(f"Tarif ({sim_result['rate']*100:.2f}%)", sim_result["region"].split("(")[0].strip())
col3.metric("Estimasi Pajak", f"{sim_result['estimated_tax']:,.2f}")
col4.metric("Nilai Bersih", f"{sim_result['net_amount']:,.2f}")

st.caption(f"Sumber perhitungan: {sim_result['source_desc']}")

st.subheader("Deteksi Koreksi & Anomali Data")
if sim_result["corrections"]:
    for item in sim_result["corrections"]:
        if item["level"] == "Koreksi":
            st.error(f"**[{item['level']}]** {item['issue']} — *Saran:* {item['action']}")
        elif item["level"] == "Peringatan":
            st.warning(f"**[{item['level']}]** {item['issue']} — *Saran:* {item['action']}")
        else:
            st.info(f"**[{item['level']}]** {item['issue']} — *Saran:* {item['action']}")
else:
    st.success("Tidak ditemukan anomali atau koreksi kritis pada data yang diperiksa.")

st.subheader("Tabel Perbandingan Tarif Multi-Daerah")
st.dataframe(sim_result["comparison_df"], use_container_width=True)

st.header("3. Visualisasi Grafik")
g_col1, g_col2 = st.columns(2)
with g_col1:
    st.subheader("Komposisi Dasar Pajak, Pajak & Nilai Bersih")
    comp_df = pd.DataFrame(
        {
            "Kategori": ["Dasar Pajak (DPP)", "Estimasi Pajak", "Nilai Bersih"],
            "Nominal": [
                sim_result["base_amount"],
                sim_result["estimated_tax"],
                sim_result["net_amount"],
            ],
        }
    ).set_index("Kategori")
    st.bar_chart(comp_df, height=280)

with g_col2:
    st.subheader("Perbandingan Estimasi Pajak Antar-Daerah")
    reg_df = sim_result["comparison_df"].set_index("Daerah / Skema")[["Estimasi Pajak"]]
    st.bar_chart(reg_df, height=300)

st.header("4. Analisis & Tanya-Jawab Asisten AI (Groq)")
doc_excerpt = parsed_data["text"] if parsed_data else ""

if st.button("Jalankan Analisis AI untuk Data Ini", type="primary"):
    with st.spinner("Menghubungi Groq AI..."):
        try:
            ai_res = analyze_with_groq(
                entity_name=entity_name,
                period_label=period_label,
                sim_result=sim_result,
                document_excerpt=doc_excerpt,
            )
            store["last_ai_output"] = ai_res
            store["history"].append({
                "entity": entity_name,
                "period": period_label,
                "region": selected_region,
                "base_amount": sim_result["base_amount"],
                "estimated_tax": sim_result["estimated_tax"],
            })
        except Exception as exc:
            st.error(f"Terjadi kesalahan saat memanggil Groq API: {exc}")

if store["last_ai_output"]:
    st.markdown("### Hasil Analisis AI")
    st.markdown(store["last_ai_output"])

st.markdown("### Tanya-Jawab Kontekstual Dokumen & Simulasi")
question = st.text_input(
    "Ajukan pertanyaan terkait angka, tarif daerah, atau koreksi data pada dokumen ini:"
)
if st.button("Kirim Pertanyaan"):
    if not question.strip():
        st.warning("Silakan tulis pertanyaannya terlebih dahulu.")
    else:
        with st.spinner("Menyusun jawaban..."):
            try:
                qa_res = analyze_with_groq(
                    entity_name=entity_name,
                    period_label=period_label,
                    sim_result=sim_result,
                    document_excerpt=doc_excerpt,
                    user_question=question,
                )
                store["last_qa_output"] = qa_res
            except Exception as exc:
                st.error(f"Gagal menjawab pertanyaan: {exc}")

if store["last_qa_output"]:
    st.markdown(store["last_qa_output"])

st.header("5. Riwayat Simulasi Entitas & Unduh Ringkasan")
if store["history"]:
    st.dataframe(store["history"], use_container_width=True)
else:
    st.caption("Belum ada riwayat analisis tersimpan untuk entitas & periode aktif ini.")

summary_text = (
    f"LAPORAN SIMULASI TAXFIN AGENT\n"
    f"Entitas: {entity_name}\n"
    f"Periode: {period_label}\n"
    f"Daerah / Skema: {sim_result['region']}\n"
    f"Dasar Pajak (DPP): {sim_result['base_amount']:,.2f}\n"
    f"Estimasi Pajak: {sim_result['estimated_tax']:,.2f}\n"
    f"Nilai Bersih: {sim_result['net_amount']:,.2f}\n\n"
    f"RINGKASAN ANALISIS AI:\n{store['last_ai_output'] or '(Belum dijalankan)'}\n"
)

st.download_button(
    label="Unduh Ringkasan Hasil (.txt)",
    data=summary_text,
    file_name=f"taxfin_ringkasan_{entity_name.lower().replace(' ', '_')}.txt",
    mime="text/plain",
)
