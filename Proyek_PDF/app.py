import os
import re
import time
import PyPDF2
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "kunci_rahasia_bebas"
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Pastikan folder upload otomatis terbuat
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def extract_church_data(pdf_path):
    text = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text() + "\n"

    # Split berdasarkan angka di awal baris (mendukung '1 ALFRED' atau '1ALFRED')
    entries = re.split(r'\n(?=\d+\s*[A-Z])', text)
    
    results = []
    for entry in entries:
        clean_entry = re.sub(r'\s+', ' ', entry).strip()
        
        # Pattern untuk Nama, Usia, Hubungan, dan Status Baptis Awal
        pattern = r'^(\d+)\s*(.*?)\s*\(\s*(\d+)\s*Th\)-\s*([A-Za-z\s]+?)(?:\s*\d+)?\s*(Sudah|Belum)'
        match = re.search(pattern, clean_entry)
        
        if match:
            no, name, age, role, b_status = match.groups()
            
            # Memisahkan sisa data untuk mencari status Sidi
            parts = re.split(r'(Sudah|Belum)', clean_entry)
            
            try:
                # Bagian Baptis (setelah 'Sudah/Belum' pertama)
                b_info = parts[2] if len(parts) > 2 else ""
                # Bagian Sidi (setelah 'Sudah/Belum' kedua)
                s_status = parts[3] if len(parts) > 3 else "-"
                s_info = parts[4] if len(parts) > 4 else ""
                
                # Pola Tanggal: 01 Jan 2024
                date_pat = r'(\d{2}\s[A-Z][a-z]{2}\s\d{4})'
                
                b_date_search = re.search(date_pat, b_info)
                b_date = b_date_search.group(0) if b_date_search else "-"
                
                s_date_search = re.search(date_pat, s_info)
                s_date = s_date_search.group(0) if s_date_search else "-"

                results.append({
                    "No": no,
                    "Nama Lengkap": name.strip(),
                    "Usia": age,
                    "Hubungan": role.strip(),
                    "Status Baptis": b_status,
                    "Tanggal Baptis": b_date,
                    "Status Sidi": s_status,
                    "Tanggal Sidi": s_date
                })
            except:
                continue
            
    return pd.DataFrame(results)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('File tidak ditemukan')
        return redirect(request.url)
    
    file = request.files['file']
    if file and file.filename.endswith('.pdf'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # 1. Ekstrak data
            df = extract_church_data(filepath)
            
            # 2. Jeda estetika agar progress bar di web terlihat (1.5 detik)
            time.sleep(1.5)
            
            # 3. Simpan ke Excel
            excel_filename = filename.replace('.pdf', '.xlsx')
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            df.to_excel(excel_path, index=False, engine='openpyxl')
            
            return send_file(excel_path, as_attachment=True)
        except Exception as e:
            flash(f'Error: {str(e)}')
            return redirect(request.url)
            
    flash('Gunakan format .pdf')
    return redirect(request.url)

if __name__ == '__main__':
    app.run(port=5000, debug=True)