
import streamlit as st
import pandas as pd
import io
import openpyxl

# Fungsi untuk memproses data (tidak ada perubahan di fungsi ini)
def process_data(source_df, template_df, target_column, mode):
    """
    Fungsi ini mengambil data sumber, menghitungnya, dan memperbarui DataFrame template.
    """
    # 1. Ekstrak dan hitung data dari file sumber
    # Filter baris di mana 'Student Code' dan 'Course Code' tidak kosong
    filtered_df = source_df.dropna(subset=['Student Code', 'Course Code'])
    # Hitung kemunculan setiap 'Site Name'
    site_counts = filtered_df['Site Name'].value_counts()
    # Buat salinan template_df agar tidak mengubah data asli secara langsung
    updated_df = template_df.copy()
    # Asumsikan kolom A (kolom pertama) di template adalah untuk 'Site Name'
    site_name_col_in_template = updated_df.columns[0]
    # 2. Iterasi melalui hasil hitungan dan perbarui template
    for site_name, count in site_counts.items():
        # Cari baris yang cocok dengan site_name di template
        match_row = updated_df[updated_df[site_name_col_in_template] == site_name]
        if not match_row.empty:
            # Jika Site Name ditemukan
            idx = match_row.index[0]
            if mode == "Add (Tambah)":
                # Ambil nilai saat ini, ubah ke numerik, anggap 0 jika kosong/error
                current_value = pd.to_numeric(updated_df.loc[idx, target_column], errors='coerce').fillna(0)
                updated_df.loc[idx, target_column] = current_value + count
            else:  # Mode "Replace (Ganti)"
                updated_df.loc[idx, target_column] = count
        else:
            # Jika Site Name tidak ditemukan
            # Buat baris baru sebagai dictionary
            new_row = {site_name_col_in_template: site_name, target_column: count}
            # Tambahkan baris baru ke DataFrame
            updated_df = pd.concat([updated_df, pd.DataFrame([new_row])], ignore_index=True)
    return updated_df

# --- UI Streamlit ---
st.set_page_config(layout="wide")
st.title("📊 Aplikasi Pemroses Data Excel")
st.write("Aplikasi ini menghitung data dari satu file Excel dan memasukkannya ke dalam file template.")

# Kolom untuk upload file
col1, col2 = st.columns(2)

with col1:
    st.header("1. Upload File Sumber")
    source_file = st.file_uploader("Pilih file Excel yang berisi data mentah", type=["xlsx"])

with col2:
    st.header("2. Upload File Template")
    template_file = st.file_uploader("Pilih file Excel template tujuan", type=["xlsx", "xlsm"])

# Lanjutkan hanya jika kedua file sudah di-upload
if source_file and template_file:
    try:
        # Baca file Excel
        # header=1 berarti kita menggunakan baris ke-2 di Excel sebagai header
        source_df = pd.read_excel(source_file, header=1) 
        # Membersihkan spasi ekstra dari nama kolom untuk keamanan
        source_df.columns = source_df.columns.str.strip() 
        # Baca template_df lebih awal agar bisa dipakai di selectbox
        template_df = pd.read_excel(template_file, sheet_name="Master Sheet")

        # --- Opsi Pemrosesan ---
        st.header("3. Atur Opsi Pemrosesan")

        # Pilihan kolom target dari file template
        target_column = st.selectbox(
            "Pilih kolom target di 'Master Sheet' untuk menempatkan hasil hitungan:",
            options=template_df.columns
        )

        # Pilihan mode (Add/Replace)
        mode = st.radio(
            "Pilih mode pembaruan data:",
            options=["Add (Tambah)", "Replace (Ganti)"],
            help="Add: Menambahkan hasil hitungan ke nilai yang sudah ada. Replace: Mengganti nilai yang ada dengan hasil hitungan baru."
        )

        # Tombol Submit
        if st.button("🚀 Proses Sekarang!"):
            with st.spinner("Sedang memproses data..."):
                # Panggil fungsi pemrosesan
                result_df = process_data(source_df, template_df, target_column, mode)
                st.header("4. Hasil")
                st.write("Data berhasil diproses. Berikut adalah pratinjau hasilnya:")
                st.dataframe(result_df.fillna('')) # Tampilkan hasil, ganti NaN dengan string kosong

                # --- Tombol Download ---
                # Menulis hasil ke file template agar format tetap terjaga
                template_file.seek(0)
                # openpyxl bisa membaca xlsm, macro tetap ada selama tidak diubah
                wb = openpyxl.load_workbook(template_file, keep_vba=True)
                ws = wb["Master Sheet"]

                # Ambil header dan mapping kolom
                header = [cell.value for cell in ws[1]]
                site_name_col = header[0]
                site_name_to_row = {}
                for row in range(2, ws.max_row+1):
                    val = ws.cell(row=row, column=1).value
                    if val:
                        site_name_to_row[val] = row

                # Update dan tambahkan baris baru jika perlu
                for idx, row_data in result_df.iterrows():
                    site_name = row_data[site_name_col]
                    if site_name in site_name_to_row:
                        row_idx = site_name_to_row[site_name]
                    else:
                        row_idx = ws.max_row + 1
                        ws.insert_rows(row_idx)
                        ws.cell(row=row_idx, column=1, value=site_name)
                        site_name_to_row[site_name] = row_idx
                    for col_idx, col_name in enumerate(header, start=1):
                        ws.cell(row=row_idx, column=col_idx, value=row_data.get(col_name, None))

                output = io.BytesIO()
                wb.save(output)
                processed_data = output.getvalue()

                # Tentukan ekstensi hasil sesuai template
                ext = ".xlsm" if template_file.name.lower().endswith(".xlsm") else ".xlsx"
                mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if ext == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                st.download_button(
                    label="📥 Download File Hasil",
                    data=processed_data,
                    file_name=f"hasil_proses{ext}",
                    mime=mime_type
                )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.warning("Pastikan file Excel yang di-upload benar dan terdapat sheet bernama 'Master Sheet' di file template.")
else:
    st.info("Silakan upload kedua file Excel untuk memulai.")