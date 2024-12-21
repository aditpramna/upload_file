import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from datetime import datetime

# Direktori output untuk menyimpan file hasil penggabungan
OUTPUT_DIRECTORY = "output_files"
os.makedirs(OUTPUT_DIRECTORY, exist_ok=True)

# Fungsi untuk membuat nama file unik berdasarkan timestamp
def generate_output_filename():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Combined_Data_{timestamp}.xlsx"

# Fungsi untuk memproses file yang diunggah
def process_uploaded_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    elif uploaded_file.name.endswith(".xlsx"):
        new_data = pd.read_excel(uploaded_file, sheet_name=None, skiprows=8)

        # Mengambil hanya baris hingga 273 dan menggabungkan semua sheet
        all_sheets_limited = {sheet_name: df.iloc[:273] for sheet_name, df in new_data.items()}

        # Menambahkan nama sheet sebagai kolom sebelum menggabungkan
        all_sheets_with_name = [
            df.assign(SheetName=sheet_name) 
            for sheet_name, df in all_sheets_limited.items()
        ]

        # Menggabungkan semua sheet menjadi satu DataFrame
        combined_df = pd.concat(all_sheets_with_name, ignore_index=True)

        # Mapping nama kolom
        column_mapping = {
            "Unnamed: 0": "no",
            "Unnamed: 4": "week",
            "Unnamed: 5": "shift",
            "Unnamed: 6": "line",
            "Unnamed: 7": "model",
            "Unnamed: 8": "part_name",
            "Unnamed: 9": "part_no",
            "Unnamed: 10": "customer",
            "Unnamed: 11": "description_of_problem",
            "Unnamed: 12": "problem_category",
            "Unnamed: 13": "suplier_or_responsible",
            "Unnamed: 14": "4m_factor",
            "Unnamed: 15": "dop_repair",
            "Unnamed: 16": "dop_scrap",
            "Unnamed: 17": "dop_total"
        }

        combined_df = combined_df.rename(columns=column_mapping)

        # Kolom yang diperlukan
        columns_to_keep = [
            "DD", "MM", "YY", "week", "shift", "line", "model", "part_name", "part_no", "customer",
            "description_of_problem", "problem_category", "suplier_or_responsible", 
            "4m_factor", "dop_repair", "dop_scrap", "dop_total"
        ]

        combined_df = combined_df[columns_to_keep]

        # Mengganti nilai 0 dengan NaN pada kolom dop_repair, dop_scrap, dan dop_total
        columns_to_replace = ["dop_repair", "dop_scrap", "dop_total"]
        combined_df[columns_to_replace] = combined_df[columns_to_replace].replace(0, np.nan)

        # Menghapus baris yang semua nilainya null
        combined_df = combined_df.dropna(how="all")

        return combined_df

# Streamlit App
st.title("File Upload dan Penggabungan Data")

# Upload banyak file oleh admin
uploaded_files = st.file_uploader("Unggah file CSV atau Excel", type=["csv", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    new_combined_data = pd.DataFrame()

    for uploaded_file in uploaded_files:
        new_data = process_uploaded_file(uploaded_file)

        # Gabungkan data baru ke dalam variabel sementara
        new_combined_data = pd.concat([new_combined_data, new_data], ignore_index=True)

    # Hapus duplikat jika seluruh baris memiliki nilai yang sama
    new_combined_data = new_combined_data.drop_duplicates(keep='first')

    # Simpan data gabungan ke file Excel baru
    output_filename = generate_output_filename()
    output_path = os.path.join(OUTPUT_DIRECTORY, output_filename)
    new_combined_data.to_excel(output_path, index=False, engine="openpyxl")

    st.success(f"File baru berhasil dibuat dan data disimpan ke: {output_path}")

    # Tampilkan data yang baru digabungkan
    st.write("Data gabungan:")
    st.dataframe(new_combined_data)

    # Tombol untuk mengunduh data
    def generate_download_buffer():
        buffer = io.BytesIO()
        new_combined_data.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        return buffer

    st.download_button(
        label="Unduh Data Gabungan",
        data=generate_download_buffer(),
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
