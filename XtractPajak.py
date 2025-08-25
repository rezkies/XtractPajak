import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

# Normalization function
def normalize_entries(data):
    normalized_entries = []
    
    for entry in data:
        pemotongan = [float(value.replace('.', '').replace(',', '.')) for value in entry['pemotongan']]
        penyetoran = [float(value.replace('.', '').replace(',', '.')) for value in entry['penyetoran']]
        saldo = [float(value.replace('.', '').replace(',', '.')) for value in entry['saldo']]
        
        for i in range(len(entry['tax'])):
            normalized_entry = {
                'date': entry['date'],
                'kwt': entry['kwt'],
                'ntpn': entry['ntpn'],
                'uraian': entry['uraian'],
                'tax': entry['tax'][i],
                'pemotongan': pemotongan[i],
                'penyetoran': penyetoran[i],
                'saldo': saldo[i]
            }
            normalized_entries.append(normalized_entry)
        
    return normalized_entries

# Header Aplikasi
st.title("🧾 XtractPajak: Konversi BKPP ke Excel")

# Deskripsi Penggunaan
st.markdown("""
Selamat datang di **XtractPajak** — alat yang dirancang untuk membantu Anda mengonversi **Buku Kas Pembantu Pajak (BKPP)** dari **Siskeudes** menjadi spreadsheet Excel yang rapi dan terstruktur hanya dalam beberapa klik.

### 📤 Cara Menggunakan:
1. Unggah file **Buku Kas Pembantu Pajak** dari aplikasi **Siskeudes** menggunakan uploader di bawah ini.
2. Tunggu proses konversi selesai dan unduh file Excel hasilnya.
""")

uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

if uploaded_file is not None:
    extracted_data = []
    pdf_filename = uploaded_file.name.rsplit('.', 1)[0]
    excel_filename = f"{pdf_filename}.xlsx"
    
    with pdfplumber.open(uploaded_file) as pdf:
        current_entry = {}

        for page in pdf.pages:
            page_text = page.extract_text()

            for line in page_text.split('\n'):
                # Regex patterns
                date_pattern = r'(\d{2}/\d{2}/\d{4})'
                kwt_pattern = r'(\d{4,5}\/[A-Z]{3}\/\d{2}\.\d{4}\/\d{4})'
                ntpn_pattern = r'NTPN\s*:\s*([A-Z0-9]+)'
                tax_pattern = r'(Uang Muka dan Jaminan|Pajak Restoran, Rumah Makan|Potongan Pajak (PPN Pusat|PPh Pasal 21|PPh Pasal 22|PPh Pasal 23|Lainnnya))'
                value_pattern = r'\d{1,3}(?:\.\d{3})*(?:,\d{2})'

                date_match = re.search(date_pattern, line)
                kwt_match = re.search(kwt_pattern, line)
                ntpn_match = re.search(ntpn_pattern, line)
                tax_match = re.search(tax_pattern, line)
                value_match = re.findall(value_pattern, line)

                if date_match and kwt_match:
                    if current_entry:
                        extracted_data.append(current_entry)
                    current_entry = {
                        'date': date_match.group(1),
                        'kwt': None,
                        'ntpn': None,
                        'uraian': '',
                        'tax': [],
                        'pemotongan': [],
                        'penyetoran': [],
                        'saldo': []
                    }

                if kwt_match:
                    current_entry['kwt'] = kwt_match.group(1)
                if ntpn_match:
                    current_entry['ntpn'] = ntpn_match.group(1)
                if tax_match and value_match:
                    tax_type = tax_match.group(1).strip()
                
                    # Ensure there are at least 3 numeric values (pemotongan, penyetoran, saldo)
                    if len(value_match) >= 3:
                        current_entry['tax'].append(tax_type)
                        current_entry['pemotongan'].append(value_match[0])
                        current_entry['penyetoran'].append(value_match[1])
                        current_entry['saldo'].append(value_match[2])
                if not (date_match or kwt_match or ntpn_match or tax_match or value_match):
                    try:
                        # Header fragment to remove if it's inside the line
                        if ("Pemotongan" in line and "Penyetoran" in line) or "Uraian" in line or "Rp" in line:
                            continue
                        # Only add if something remains
                        current_entry['uraian'] += line + ' '
                    except:
                        continue

        if current_entry:
            extracted_data.append(current_entry)

    # Normalize the extracted data
    normalized_data = normalize_entries(extracted_data)
    df = pd.DataFrame(normalized_data)

    # Buat ringkasan total berdasarkan jenis pajak
    summary = df.groupby("tax")[["pemotongan", "penyetoran"]].sum().reset_index()

    st.subheader("📊 Ringkasan Pemotongan dan Penyetoran per Jenis Pajak")
    st.dataframe(summary.style.format({"pemotongan": "Rp {:,.2f}", "penyetoran": "Rp {:,.2f}"}))
    
    # Convert DataFrame to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    st.success("Extraction and normalization complete!")
    st.download_button(
        label="Download Excel file",
        data=output,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("""
---
#### 👨‍💻 Kredit  
Dikembangkan oleh [@rezkies](https://github.com/rezkies)  
Lihat kode sumber di [GitHub](https://github.com/rezkies/XtractPajak)
""")
