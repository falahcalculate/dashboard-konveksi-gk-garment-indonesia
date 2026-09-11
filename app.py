import streamlit as st
import pandas as pd
import io
import json
import datetime
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# =========================================================
# KONFIGURASI TAMPILAN WEB
# =========================================================
st.set_page_config(page_title="Dashboard Konveksi GK Garment Indonesia", layout="wide", page_icon="🧵")

st.title("🧵 Dashboard Produksi Konveksi (AI-Powered)")
st.write(
    "Upload file Excel pesanan customer, atur fitur otomatis yang kamu butuhkan, lalu biarkan AI "
    "merapikan datanya dan hasilkan Excel rekap lengkap dengan rumus yang bisa kamu ubah sendiri."
)

# =========================================================
# SIDEBAR PENGATURAN
# =========================================================
with st.sidebar:
    st.header("⚙️ Pengaturan AI")
    api_key = st.text_input("Masukkan Gemini API Key:", type="password", help="Dapatkan gratis di aistudio.google.com")
    model_name = st.selectbox("Model Gemini", ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite"], index=0)

    st.markdown("---")
    st.header("🧩 Fitur Otomatis AI (opsional)")
    st.caption("Matikan sesuai kebutuhan kalau admin mau merekap sendiri secara manual — cocok untuk mode rekap polos.")
    sertakan_standarisasi = st.checkbox("Standarisasi & gabung baris duplikat oleh AI", value=True)
    sertakan_prioritas = st.checkbox("Prioritas otomatis (Tinggi/Sedang/Rendah)", value=True)
    sertakan_estimasi_kain = st.checkbox("Estimasi Kain & Biaya (rumus Excel)", value=True)

    rasio_kain, harga_kain = 3.5, 45000
    prioritas_tinggi, prioritas_sedang = 300, 100
    if sertakan_estimasi_kain:
        st.markdown("**🧮 Asumsi Perhitungan Kain**")
        rasio_kain = st.number_input("Rasio Kain (pcs per kg)", min_value=0.1, value=3.5, step=0.1,
                                      help="Berapa pcs pakaian yang bisa dibuat dari 1 kg kain.")
        harga_kain = st.number_input("Harga Kain per Kg (Rp)", min_value=0, value=45000, step=1000)
    if sertakan_prioritas:
        st.markdown("**🚦 Ambang Batas Prioritas**")
        prioritas_tinggi = st.number_input("Batas pcs untuk Prioritas Tinggi", min_value=1, value=300, step=10)
        prioritas_sedang = st.number_input("Batas pcs untuk Prioritas Sedang", min_value=1, value=100, step=10)

    st.markdown("---")
    st.markdown("**Cara Pakai:**")
    st.markdown("1. Masukkan API Key & pilih fitur otomatis di atas")
    st.markdown("2. Upload file Excel pesanan")
    st.markdown("3. Sesuaikan instruksi/prompt bila perlu")
    st.markdown("4. Klik **Olah Data Sekarang**")

if api_key:
    genai.configure(api_key=api_key)


# =========================================================
# PROMPT DEFAULT — MENYESUAIKAN FITUR YANG DIAKTIFKAN
# =========================================================
def build_default_prompt(sertakan_standarisasi, sertakan_prioritas, sertakan_estimasi_kain,
                          prioritas_tinggi, prioritas_sedang):
    langkah = []
    n = 1
    if sertakan_standarisasi:
        langkah.append(
            f"{n}. STANDARISASI: samakan penulisan Nama Item dan Warna yang sebenarnya sama tapi ditulis beda "
            f'(contoh: "kaos polo", "Kaos Polo ", "KAOS POLO" -> "Kaos Polo"). Gunakan Title Case.'
        )
        n += 1
        langkah.append(
            f"{n}. GABUNGKAN DUPLIKAT: jika ada baris dengan Kode Pesanan + Nama Item + Warna yang identik "
            f"setelah distandarisasi, jumlahkan Total Pcs-nya menjadi satu baris saja."
        )
        n += 1
    langkah.append(
        f"{n}. VALIDASI: jika Total Pcs kosong, 0, negatif, atau tidak berupa angka, tetap sertakan barisnya "
        f'dengan Total Pcs = 0 dan jelaskan masalahnya di kolom Catatan (mis. "Qty tidak terbaca dari sumber").'
    )
    n += 1
    if sertakan_prioritas:
        langkah.append(
            f"{n}. TENTUKAN PRIORITAS setiap baris berdasarkan Total Pcs:\n"
            f'   - Total Pcs >= {int(prioritas_tinggi)} -> "Tinggi"\n'
            f'   - {int(prioritas_sedang)} <= Total Pcs < {int(prioritas_tinggi)} -> "Sedang"\n'
            f'   - Total Pcs < {int(prioritas_sedang)} -> "Rendah"\n'
            f"   Jika data asli punya kolom tanggal/deadline, pertimbangkan juga: deadline < 7 hari dari sekarang "
            f'otomatis "Tinggi" berapa pun jumlah pcs-nya, dan sebutkan alasannya singkat di Catatan.'
        )
        n += 1
    langkah.append(
        f"{n}. CATATAN: isi HANYA jika ada hal yang perlu diperhatikan admin (data ambigu, digabung dari baris "
        f'lain, quantity mencurigakan, deadline mepet). Kosongkan ("") jika tidak ada catatan.'
    )
    n += 1
    if sertakan_estimasi_kain:
        langkah.append(f"{n}. JANGAN menghitung estimasi kain atau biaya — itu akan dihitung otomatis oleh rumus Excel.")
        n += 1

    keys = ['"Kode Pesanan"', '"Nama Item"', '"Warna"', '"Total Pcs" (angka/integer)', '"Jenis Sablon/Bordir"']
    if sertakan_prioritas:
        keys.append('"Prioritas"')
    keys.append('"Catatan"')

    intro = "Kamu adalah admin produksi konveksi senior yang teliti. Olah data pesanan berikut menjadi\n" \
            "satu tabel rekap produksi yang BERSIH dan SIAP PAKAI"
    if not (sertakan_standarisasi or sertakan_prioritas):
        intro += " (mode rekap polos: cukup rapikan tabelnya, tanpa menilai atau memutuskan apa pun)"
    intro += ", dengan langkah wajib berikut:\n\n"

    prompt = intro + "\n".join(langkah) + "\n\n"
    prompt += (
        "Kembalikan HANYA array JSON murni (tanpa markdown, tanpa penjelasan) berisi objek dengan KUNCI PERSIS ini:\n"
        + ", ".join(keys) + "."
    )
    return prompt


REQUIRED_AI_KEYS = ["Kode Pesanan", "Nama Item", "Warna", "Total Pcs", "Jenis Sablon/Bordir", "Catatan"]
if sertakan_prioritas:
    REQUIRED_AI_KEYS.append("Prioritas")

default_prompt = build_default_prompt(sertakan_standarisasi, sertakan_prioritas, sertakan_estimasi_kain,
                                       prioritas_tinggi, prioritas_sedang)


# =========================================================
# FUNGSI: BANGUN WORKBOOK EXCEL (rumus + styling, semua bagian opsional)
# =========================================================
def build_excel_workbook(df: pd.DataFrame, rasio_kain: float, harga_kain: float,
                          include_prioritas: bool = True, include_estimasi_kain: bool = True) -> io.BytesIO:
    HEADER_FILL = PatternFill("solid", fgColor="1F3864")
    HEADER_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=11)
    ASSUMPTION_FILL = PatternFill("solid", fgColor="FFF2CC")
    ASSUMPTION_FONT = Font(color="0000FF", bold=True, name="Arial", size=11)
    TITLE_FONT = Font(bold=True, size=16, name="Arial", color="1F3864")
    TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
    TOTAL_FONT = Font(bold=True, name="Arial")
    BAND_FILL = PatternFill("solid", fgColor="F2F2F2")
    THIN = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    PRIORITY_COLORS = {
        "Tinggi": PatternFill("solid", fgColor="F8CBAD"),
        "Sedang": PatternFill("solid", fgColor="FFE699"),
        "Rendah": PatternFill("solid", fgColor="C6E0B4"),
    }
    COL_WIDTHS = {
        "No": 5, "Kode Pesanan": 14, "Nama Item": 20, "Warna": 14, "Total Pcs": 11,
        "Jenis Sablon/Bordir": 18, "Prioritas": 11, "Catatan": 30, "Sumber Tab": 14,
        "Estimasi Kain (Kg)": 16, "Estimasi Biaya Kain (Rp)": 20,
    }

    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Produksi"

    base_cols = ["No", "Kode Pesanan", "Nama Item", "Warna", "Total Pcs", "Jenis Sablon/Bordir"]
    if include_prioritas:
        base_cols.append("Prioritas")
    base_cols += ["Catatan", "Sumber Tab"]
    extra_cols = ["Estimasi Kain (Kg)", "Estimasi Biaya Kain (Rp)"] if include_estimasi_kain else []
    all_cols = base_cols + extra_cols
    n_cols = len(all_cols)

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    ws.cell(row=row, column=1, value="REKAP PRODUKSI KONVEKSI").font = TITLE_FONT
    ws.row_dimensions[row].height = 26
    row += 1
    ws.cell(row=row, column=1, value=f"Dibuat otomatis: {datetime.datetime.now().strftime('%d %B %Y %H:%M')}")
    ws.cell(row=row, column=1).font = Font(italic=True, size=9, color="808080")
    row += 2  # baris kosong pemisah

    RATIO_CELL = HARGA_CELL = None
    if include_estimasi_kain:
        ws.cell(row=row, column=1, value="Asumsi Rasio Kain (pcs/kg)").font = Font(bold=True, name="Arial")
        rasio_cell_ref = ws.cell(row=row, column=2, value=rasio_kain)
        rasio_cell_ref.fill = ASSUMPTION_FILL
        rasio_cell_ref.font = ASSUMPTION_FONT
        rasio_cell_ref.border = BORDER
        rasio_cell_ref.number_format = "0.00"
        RATIO_CELL = f"${get_column_letter(2)}${row}"
        row += 1
        ws.cell(row=row, column=1, value="Asumsi Harga Kain (Rp/kg)").font = Font(bold=True, name="Arial")
        harga_cell_ref = ws.cell(row=row, column=2, value=harga_kain)
        harga_cell_ref.fill = ASSUMPTION_FILL
        harga_cell_ref.font = ASSUMPTION_FONT
        harga_cell_ref.border = BORDER
        harga_cell_ref.number_format = '"Rp"#,##0'
        HARGA_CELL = f"${get_column_letter(2)}${row}"
        row += 2  # baris kosong pemisah

    header_row = row
    for idx, col_name in enumerate(all_cols, start=1):
        c = ws.cell(row=header_row, column=idx, value=col_name)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 32

    pcs_col_idx = base_cols.index("Total Pcs") + 1
    pcs_col_letter = get_column_letter(pcs_col_idx)
    jenis_col_idx = base_cols.index("Jenis Sablon/Bordir") + 1
    jenis_col_letter = get_column_letter(jenis_col_idx)
    kode_col_letter = get_column_letter(base_cols.index("Kode Pesanan") + 1)
    prioritas_col_idx = base_cols.index("Prioritas") + 1 if include_prioritas else None
    prioritas_col_letter = get_column_letter(prioritas_col_idx) if include_prioritas else None
    kain_col_idx = len(base_cols) + 1 if include_estimasi_kain else None
    biaya_col_idx = len(base_cols) + 2 if include_estimasi_kain else None
    kain_col_letter = get_column_letter(kain_col_idx) if include_estimasi_kain else None
    biaya_col_letter = get_column_letter(biaya_col_idx) if include_estimasi_kain else None

    first_data_row = header_row + 1
    n_rows = len(df)
    for i, (_, drow) in enumerate(df.iterrows()):
        r = first_data_row + i
        for j, col_name in enumerate(base_cols, start=1):
            val = (i + 1) if col_name == "No" else drow.get(col_name, "")
            cell = ws.cell(row=r, column=j, value=val)
            cell.border = BORDER
            if i % 2 == 1:
                cell.fill = BAND_FILL
        if include_estimasi_kain:
            kain_cell = ws.cell(row=r, column=kain_col_idx, value=f"={pcs_col_letter}{r}/{RATIO_CELL}")
            kain_cell.number_format = "0.00"
            kain_cell.border = BORDER
            biaya_cell = ws.cell(row=r, column=biaya_col_idx, value=f"={kain_col_letter}{r}*{HARGA_CELL}")
            biaya_cell.number_format = '"Rp"#,##0'
            biaya_cell.border = BORDER
            if i % 2 == 1:
                kain_cell.fill = BAND_FILL
                biaya_cell.fill = BAND_FILL

    last_data_row = first_data_row + n_rows - 1 if n_rows > 0 else first_data_row

    total_row = last_data_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = TOTAL_FONT
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=max(pcs_col_idx - 1, 1))
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="right")

    total_pcs_formula = f"=SUM({pcs_col_letter}{first_data_row}:{pcs_col_letter}{last_data_row})" if n_rows else 0
    tc = ws.cell(row=total_row, column=pcs_col_idx, value=total_pcs_formula)
    tc.number_format = "#,##0"
    if include_estimasi_kain:
        total_kain_formula = f"=SUM({kain_col_letter}{first_data_row}:{kain_col_letter}{last_data_row})" if n_rows else 0
        ws.cell(row=total_row, column=kain_col_idx, value=total_kain_formula).number_format = "0.00"
        total_biaya_formula = f"=SUM({biaya_col_letter}{first_data_row}:{biaya_col_letter}{last_data_row})" if n_rows else 0
        ws.cell(row=total_row, column=biaya_col_idx, value=total_biaya_formula).number_format = '"Rp"#,##0'

    for col in range(1, n_cols + 1):
        c = ws.cell(row=total_row, column=col)
        c.fill = TOTAL_FILL
        c.font = TOTAL_FONT
        c.border = Border(top=Side(style="double"), bottom=THIN, left=THIN, right=THIN)

    if n_rows > 0 and include_prioritas:
        prio_range = f"{prioritas_col_letter}{first_data_row}:{prioritas_col_letter}{last_data_row}"
        for label in ["Tinggi", "Sedang", "Rendah"]:
            ws.conditional_formatting.add(
                prio_range, CellIsRule(operator="equal", formula=[f'"{label}"'], fill=PRIORITY_COLORS[label])
            )
        dv = DataValidation(type="list", formula1='"Tinggi,Sedang,Rendah"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(prio_range)

    if n_rows > 0:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(n_cols)}{last_data_row}"
    ws.freeze_panes = f"A{first_data_row}"

    for i, col_name in enumerate(all_cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = COL_WIDTHS.get(col_name, 16)

    # ================= SHEET RINGKASAN =================
    ws2 = wb.create_sheet("Ringkasan")
    ws2.merge_cells("A1:D1")
    ws2["A1"] = "RINGKASAN PRODUKSI"
    ws2["A1"].font = TITLE_FONT
    ws2.row_dimensions[1].height = 26

    if n_rows > 0:
        total_pesanan_formula = (
            f"=SUMPRODUCT(1/COUNTIF('Rekap Produksi'!{kode_col_letter}{first_data_row}:"
            f"{kode_col_letter}{last_data_row},'Rekap Produksi'!{kode_col_letter}{first_data_row}:"
            f"{kode_col_letter}{last_data_row}))"
        )
    else:
        total_pesanan_formula = 0

    kpi_labels = [
        ("Total Pesanan (Kode Unik)", total_pesanan_formula, "#,##0"),
        ("Total Pcs Keseluruhan", f"='Rekap Produksi'!{pcs_col_letter}{total_row}", "#,##0"),
    ]
    if include_estimasi_kain:
        kpi_labels.append(("Total Estimasi Kain (Kg)", f"='Rekap Produksi'!{kain_col_letter}{total_row}", "0.00"))
        kpi_labels.append(("Total Estimasi Biaya Kain (Rp)", f"='Rekap Produksi'!{biaya_col_letter}{total_row}", '"Rp"#,##0'))

    r = 3
    for label, formula, fmt in kpi_labels:
        ws2.cell(row=r, column=1, value=label).font = Font(bold=True, name="Arial")
        ws2.cell(row=r, column=1).border = BORDER
        val_cell = ws2.cell(row=r, column=2, value=formula)
        val_cell.number_format = fmt
        val_cell.fill = PatternFill("solid", fgColor="D9E1F2")
        val_cell.font = Font(bold=True, size=12, name="Arial")
        val_cell.border = BORDER
        r += 1

    r += 1
    ws2.cell(row=r, column=1, value="Breakdown per Jenis Sablon/Bordir").font = Font(bold=True, size=12, name="Arial")
    r += 1
    breakdown_header_row = r
    for j, col_name in enumerate(["Jenis", "Total Pcs"], start=1):
        c = ws2.cell(row=r, column=j, value=col_name)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    r += 1
    breakdown_first_row = r
    jenis_unik = sorted(df["Jenis Sablon/Bordir"].dropna().astype(str).unique().tolist()) if n_rows > 0 else []
    for jenis in jenis_unik:
        ws2.cell(row=r, column=1, value=jenis).border = BORDER
        f = (f"=SUMIF('Rekap Produksi'!{jenis_col_letter}{first_data_row}:{jenis_col_letter}{last_data_row},"
             f"A{r},'Rekap Produksi'!{pcs_col_letter}{first_data_row}:{pcs_col_letter}{last_data_row})")
        fc = ws2.cell(row=r, column=2, value=f)
        fc.number_format = "#,##0"
        fc.border = BORDER
        r += 1
    breakdown_last_row = r - 1

    prio_header_row = prio_first_row = prio_last_row = None
    if include_prioritas:
        r += 1
        ws2.cell(row=r, column=1, value="Jumlah Pesanan per Prioritas").font = Font(bold=True, size=12, name="Arial")
        r += 1
        prio_header_row = r
        for j, col_name in enumerate(["Prioritas", "Jumlah Baris"], start=1):
            c = ws2.cell(row=r, column=j, value=col_name)
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.border = BORDER
        r += 1
        prio_first_row = r
        for label in ["Tinggi", "Sedang", "Rendah"]:
            ws2.cell(row=r, column=1, value=label).border = BORDER
            f = (f"=COUNTIF('Rekap Produksi'!{prioritas_col_letter}{first_data_row}:"
                 f"{prioritas_col_letter}{last_data_row},A{r})") if n_rows > 0 else 0
            fc = ws2.cell(row=r, column=2, value=f)
            fc.number_format = "#,##0"
            fc.border = BORDER
            r += 1
        prio_last_row = r - 1

    ws2.column_dimensions["A"].width = 32
    ws2.column_dimensions["B"].width = 22

    chart_anchor_row = 3
    if jenis_unik:
        chart = BarChart()
        chart.title = "Total Pcs per Jenis Sablon/Bordir"
        chart.y_axis.title = "Pcs"
        chart.x_axis.title = "Jenis"
        data_ref = Reference(ws2, min_col=2, min_row=breakdown_header_row, max_row=breakdown_last_row)
        cats_ref = Reference(ws2, min_col=1, min_row=breakdown_first_row, max_row=breakdown_last_row)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.width, chart.height = 16, 8
        ws2.add_chart(chart, f"D{chart_anchor_row}")
        chart_anchor_row += 16

    if include_prioritas and prio_header_row:
        pie = PieChart()
        pie.title = "Proporsi Prioritas Pesanan"
        pdata = Reference(ws2, min_col=2, min_row=prio_header_row, max_row=prio_last_row)
        pcats = Reference(ws2, min_col=1, min_row=prio_first_row, max_row=prio_last_row)
        pie.add_data(pdata, titles_from_data=True)
        pie.set_categories(pcats)
        pie.width, pie.height = 16, 8
        ws2.add_chart(pie, f"D{chart_anchor_row}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================================================
# AREA UPLOAD FILE
# =========================================================
uploaded_file = st.file_uploader("📂 Upload File Excel Pesanan Customer (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = []
        for sheet_name in xls.sheet_names:
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
            df_sheet["Nama Tab Excel"] = sheet_name
            all_sheets.append(df_sheet)
        df_raw = pd.concat(all_sheets, ignore_index=True)

        st.write("### 📄 Preview Data Asli dari Customer")
        st.dataframe(df_raw, use_container_width=True)

        st.write("---")
        st.write("### 🤖 Instruksi Pengolahan Data (Prompt Admin)")
        user_prompt = st.text_area("Tuliskan instruksi/prompt untuk AI di sini:", value=default_prompt, height=320)

        if st.button("🚀 Olah Data Sekarang", type="primary"):
            if not api_key:
                st.error("⚠️ Silakan masukkan Gemini API Key terlebih dahulu di menu sidebar sebelah kiri!")
            else:
                with st.spinner("AI sedang membaca dan mengolah data pesanan..."):
                    json_input = df_raw.to_json(orient="records", force_ascii=False)
                    system_prompt = f"""
                    Berikut data pesanan customer dari Excel (JSON):
                    {json_input}

                    Instruksi dari Admin:
                    {user_prompt}
                    """

                    try:
                        model = genai.GenerativeModel(model_name)
                        response = model.generate_content(system_prompt)
                    except Exception as api_err:
                        st.error(f"Gagal menghubungi Gemini API: {api_err}")
                        st.stop()

                    raw_text = response.text.strip()
                    for fence in ("```json", "```"):
                        if raw_text.startswith(fence):
                            raw_text = raw_text[len(fence):]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()

                    try:
                        parsed_json = json.loads(raw_text)
                        if isinstance(parsed_json, dict):
                            parsed_json = parsed_json.get("data", [parsed_json])
                        df_result = pd.DataFrame(parsed_json)

                        # === 1. KAMUS VARIASI NAMA KOLOM (fleksibel utk beda format tiap customer) ===
                        alias_map = {
                            "Kode Pesanan": ["kode pesanan", "kode", "po", "no po", "order id", "id pesanan"],
                            "Nama Item": ["nama item", "item", "produk", "artikel", "nama barang", "jenis pakaian"],
                            "Warna": ["warna", "color", "colour"],
                            "Total Pcs": ["total pcs", "qty", "jumlah", "pcs", "quantity", "total", "jumlah pcs"],
                            "Jenis Sablon/Bordir": ["jenis sablon/bordir", "jenis sablon", "sablon", "bordir", "sablon/bordir", "aplikasi"],
                            "Prioritas": ["prioritas", "priority", "urgensi"],
                            "Catatan": ["catatan", "notes", "keterangan", "remark", "catatan customer"],
                        }

                        def _normalize_col(name):
                            # samakan huruf kecil, ganti "_"/tanda baca umum jadi spasi, rapikan spasi ganda
                            s = str(name).lower().replace("_", " ")
                            for ch in [".", ",", ":", "-"]:
                                s = s.replace(ch, " ")
                            return " ".join(s.split())

                        alias_lookup = {
                            std_key: {_normalize_col(p) for p in patterns}
                            for std_key, patterns in alias_map.items()
                        }

                        # === 2. SAMAKAN NAMA KOLOM SECARA OTOMATIS ===
                        # Satu std_key hanya boleh dipetakan SEKALI, supaya tidak ada dua kolom sumber
                        # yang tabrakan jadi satu nama (mis. "Qty" & "Jumlah" sama-sama ada) —
                        # itu bisa membuat df_result[...] balik jadi DataFrame, bukan Series, dan error.
                        renamed_cols = {}
                        used_targets = set()
                        for col in df_result.columns:
                            col_clean = _normalize_col(col)
                            for std_key, patterns in alias_lookup.items():
                                if std_key in used_targets:
                                    continue
                                if col_clean in patterns:
                                    renamed_cols[col] = std_key
                                    used_targets.add(std_key)
                                    break
                        df_result.rename(columns=renamed_cols, inplace=True)

                        # === 3. FALLBACK QTY DARI DATA ASLI ===
                        # Hanya dipakai kalau AI benar-benar tidak mengisi angka SAMA SEKALI, dan
                        # hanya kalau jumlah baris masih sama persis dengan data asli (df_raw) —
                        # kalau AI sempat menggabung/membuang baris (mode standarisasi & gabung
                        # duplikat), posisi baris sudah tidak sama lagi sehingga fallback ini
                        # dilewati agar tidak salah pasang angka ke baris yang keliru.
                        has_valid_qty = (
                            "Total Pcs" in df_result.columns
                            and pd.to_numeric(df_result["Total Pcs"], errors="coerce").fillna(0).sum() > 0
                        )
                        if not has_valid_qty and len(df_result) == len(df_raw):
                            qty_cols = [c for c in df_raw.columns if _normalize_col(c) in alias_lookup["Total Pcs"]]
                            if qty_cols:
                                df_result["Total Pcs"] = df_raw[qty_cols[0]].values

                        # === 4. LENGKAPI KOLOM WAJIB & PASTIKAN TIPE DATA ===
                        for key in REQUIRED_AI_KEYS:
                            if key not in df_result.columns:
                                df_result[key] = 0 if key == "Total Pcs" else ""
                        df_result["Total Pcs"] = pd.to_numeric(df_result["Total Pcs"], errors="coerce").fillna(0).astype(int)
                        if "Sumber Tab" not in df_result.columns:
                            df_result["Sumber Tab"] = ""
                        df_result["No"] = range(1, len(df_result) + 1)

                        st.success(f"✅ Data Berhasil Diolah oleh AI! ({len(df_result)} baris rekap)")

                        cols = st.columns(4 if sertakan_estimasi_kain else 2)
                        cols[0].metric("Total Pesanan Unik", df_result["Kode Pesanan"].nunique())
                        cols[1].metric("Total Pcs", int(df_result["Total Pcs"].sum()))
                        if sertakan_estimasi_kain:
                            total_kg = df_result["Total Pcs"].sum() / rasio_kain
                            cols[2].metric("Estimasi Kain (Kg)", f"{total_kg:,.2f}")
                            cols[3].metric("Estimasi Biaya Kain", f"Rp {total_kg * harga_kain:,.0f}")

                        st.write("### 📊 Hasil Rekapitulasi Produksi")
                        st.dataframe(df_result, use_container_width=True)

                        excel_bytes = build_excel_workbook(
                            df_result, rasio_kain, harga_kain,
                            include_prioritas=sertakan_prioritas,
                            include_estimasi_kain=sertakan_estimasi_kain,
                        )

                        st.download_button(
                            label="📥 Download Hasil Olah Data (.xlsx)",
                            data=excel_bytes,
                            file_name=f"Rekap_Produksi_Konveksi_{datetime.date.today().isoformat()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as json_err:
                        st.error(f"Gagal membaca respon AI sebagai tabel ({json_err}). Respon mentah dari AI:")
                        st.code(response.text)

    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
else:
    st.info("👆 Silakan upload file Excel pesanan customer untuk memulai.")
