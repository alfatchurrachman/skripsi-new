# =======RAW DATA INPUT=========

import pandas as pd

# 1. Baca file .header tapi lewati 4 baris pertama
header_file = 'DNADamage.header'
phsp_file = 'DNADamage.phsp'

# ========PARSING DATA=============

try:
    with open(header_file, 'r') as f:
        lines = f.readlines()
        
        # Ambil mulai dari baris ke-5 sampai akhir (index 4 ke atas)
        # Pastikan tidak ada baris kosong yang ikut terbaca
        column_names = [line.strip() for line in lines[4:] if line.strip()]
        
    print(f"Nama kolom ditemukan ({len(column_names)}): {column_names}")

    # 2. Baca data dari file .phsp
    # Gunakan nama kolom yang sudah kita ambil di atas
    df = pd.read_csv(phsp_file, sep=r'\s+', names=column_names)

    # 3. Simpan ke Excel
    output_filename = 'summary'+'_'+'.xlsx'
    df.to_excel(output_filename, header=True)
    
    print(f"Sukses! Data disimpan ke '{output_filename}'")
    print("Format: Kolom A = Nama Besaran, Kolom B = Nilai Data")

except Exception as e:
    print(f"Terjadi Error: {e}")

