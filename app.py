import streamlit as st
import pandas as pd
import io
import google.generativeai as genai
import json

# Konfigurasi Tampilan Web
st.set_page_config(page_title="Dashboard Konveksi GK Garment Indonesia", layout="wide", page_icon="🧵")

st.title("🧵 Dashboard Produksi Konveksi (AI-Powered)")
st.write("Upload file Excel pesanan customer, masukkan instruksi/prompt, dan unduh hasil olahan data dalam format Excel.")

# Sidebar Pengaturan
with st.sidebar:
    st.header("⚙️ Pengaturan AI")
    api_key = st.text_input("Masukkan Gemini API Key:", type="password", help="Dapatkan gratis di aistudio.google.com")
    st.markdown("---")
    st.markdown("**Cara Pakai:**")
    st.markdown("1. Masukkan API Key di atas")
    st.markdown("2. Upload file Excel pesanan")
    st.markdown("3. Tulis instruksi/prompt")
    st.markdown("4. Klik **Olah Data Sekarang**")

if api_key:
    genai.configure(api_key=api_key)

# Area Upload File
uploaded_file = st.file_uploader("📂 Upload File Excel Pesanan Customer (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        st.write("### 📄 Preview Data Asli dari Customer")
        st.dataframe(df_raw, use_container_width=True)
        
        st.write("---")
        st.write("### 🤖 Instruksi Pengolahan Data (Prompt Admin)")
        
        default_prompt = (
            "Analisis data pesanan ini. Buatkan tabel rekap produksi rapi yang berisi kolom: "
            "No, Kode Pesanan, Nama Item, Warna, Total Pcs, Estimasi Kain (Kg), Jenis Sablon/Bordir, Catatan. "
            "Asumsi rasio kebutuhan kain: 1 kg = 3.5 pcs. "
            "Kembalikan HANYA array JSON dari objek tanpa penjelasan teks atau markdown."
        )
        
        user_prompt = st.text_area("Tuliskan instruksi/prompt untuk AI di sini:", value=default_prompt, height=120)
        
        if st.button("🚀 Olah Data Sekarang", type="primary"):
            if not api_key:
                st.error("⚠️ Silakan masukkan Gemini API Key terlebih dahulu di menu sidebar sebelah kiri!")
            else:
                with st.spinner("AI sedang membaca dan mengolah data pesanan..."):
                    # Konversi Data Excel ke JSON
                    json_input = df_raw.to_json(orient="records")
                    
                    system_prompt = f"""
                    Kamu adalah AI Asisten Operasional Produksi Konveksi berpengalaman.
                    Berikut data pesanan customer dari Excel:
                    {json_input}

                    Instruksi dari Admin:
                    {user_prompt}

                    SANGAT PENTING:
                    1. Output HARUS berupa JSON array of objects murni.
                    2. JANGAN tambahkan teks pengantar, salam, atau penjelasan apapun.
                    3. JANGAN bungkus dengan formatting codeblock markdown seperti ```json.
                    """
                    
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    response = model.generate_content(system_prompt)
                    
                    # Pembersihan format respon AI
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.startswith("```"):
                        raw_text = raw_text[3:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    try:
                        parsed_json = json.loads(raw_text)
                        df_result = pd.DataFrame(parsed_json)
                        
                        st.success("✅ Data Berhasil Diolah oleh AI!")
                        
                        st.write("### 📊 Hasil Rekapitulasi Produksi")
                        st.dataframe(df_result, use_container_width=True)
                        
                        # Buat File Excel untuk Di-download
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_result.to_excel(writer, index=False, sheet_name='Rekap_Produksi')
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Download Hasil Olah Data (.xlsx)",
                            data=output,
                            file_name="Hasil_Rekap_Produksi_Konveksi.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as json_err:
                        st.error("Gagal membaca respon AI sebagai tabel. Respon mentah dari AI:")
                        st.code(response.text)
                        
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
else:
    st.info("👆 Silakan upload file Excel pesanan customer untuk memulai.")
