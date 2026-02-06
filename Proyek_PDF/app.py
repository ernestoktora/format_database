import streamlit as st
import pandas as pd
import pdfplumber
import io

st.set_page_config(page_title="Konverter PDF GKP", layout="wide")

st.title("📄 Konverter PDF ke Excel - Wilayah 4")
st.write("Gunakan aplikasi ini untuk mengubah data jemaat dari PDF ke Excel secara otomatis.")

# Upload file
uploaded_file = st.file_uploader("Unggah file Wilayah 4.pdf", type="pdf")

if uploaded_file is not None:
    with st.spinner('Sedang memproses data jemaat...'):
        all_data = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table:
                        # Membersihkan baris dari karakter kosong/None
                        clean_row = [str(item).replace('\n', ' ') if item else "" for item in row]
                        if any(clean_row):
                            all_data.append(clean_row)
        
        if all_data:
            # Gunakan baris pertama sebagai header atau definisikan sendiri
            df = pd.DataFrame(all_data)
            
            st.subheader("Pratinjau Data")
            st.dataframe(df.head(10))

            # Fungsi konversi ke Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False)
            
            st.download_button(
                label="📥 Unduh Hasil Excel",
                data=output.getvalue(),
                file_name="Data_Jemaat_Wilayah_4.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Data siap diunduh!")
