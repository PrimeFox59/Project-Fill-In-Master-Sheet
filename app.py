
import streamlit as st
import pandas as pd
import io
import openpyxl

# Optional viz libs
try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except Exception:  # pragma: no cover
    px = None
    PLOTLY_AVAILABLE = False

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

# --- Constants ---
# Categories are derived dynamically from the 'Master Sheet' header (excluding the first column).

# --- App Title & Sidebar ---
st.title("📊 Master Sheet Assistant")
st.caption("Pisah menu: Dashboard & Input Data")

menu = st.sidebar.radio("Menu", ["Dashboard", "Input Data", "Panduan Pengguna"], index=1)

# Session placeholders
if "template_df" not in st.session_state:
    st.session_state.template_df = None
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "last_processed_ext" not in st.session_state:
    st.session_state.last_processed_ext = ".xlsx"

def render_dashboard():
    st.header("📈 Dashboard Partner Engagement")
    st.write("Eksplorasi data Master Sheet secara interaktif: ringkasan, Top/Bottom, perbandingan kategori, dan profil perusahaan.")

    # Simple styling touch
    st.markdown(
        """
        <style>
                .section-title {margin-top: 0.5rem;}

                /* Modern KPI cards */
                .kpi-card{
                    position: relative;
                    display: flex; flex-direction: column; gap: 6px;
                    padding: 14px 16px; border-radius: 14px;
                    background: linear-gradient(135deg, #ffffff 0%, #f5f8ff 100%);
                    border: 1px solid #e6eaf2; box-shadow: 0 2px 10px rgba(28, 39, 60, 0.06);
                    min-height: 110px;
                }
                .kpi-head{display:flex; align-items:center; gap:8px; color:#4b5563; font-weight:600; font-size: 0.95rem;}
                .kpi-icon{width: 36px; height: 36px; border-radius: 10px; display:flex; align-items:center; justify-content:center; color:#fff; font-size: 18px;}
                .kpi-icon.primary{background: linear-gradient(135deg, #2563eb, #1d4ed8);} /* blue */
                .kpi-icon.success{background: linear-gradient(135deg, #10b981, #059669);}  /* green */
                .kpi-icon.warning{background: linear-gradient(135deg, #f59e0b, #d97706);}  /* orange */
                .kpi-icon.slate{background: linear-gradient(135deg, #64748b, #475569);}    /* slate */
                .kpi-value{font-size: 2rem; font-weight: 700; letter-spacing: -0.5px; color:#111827; margin-top:4px;}
                .kpi-sub{color:#6b7280; font-size: 0.85rem;}
                .kpi-badge{position:absolute; right:10px; top:10px; font-size: 0.75rem; padding: 2px 8px; border-radius: 999px; background:#eef2ff; color:#3730a3; border:1px solid #e0e7ff;}
                .muted{color:#6b7280}
        
                /* Subtle animated gradient dot */
                .kpi-dot{position:absolute; right:14px; bottom:14px; width:8px; height:8px; border-radius:999px; background: radial-gradient(circle at 30% 30%, #93c5fd, #2563eb); opacity:0.5}
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Upload template jika belum ada di session
    uploaded_template = st.file_uploader(
        "Upload file Excel template (berisi sheet 'Master Sheet')",
        type=["xlsx", "xlsm"],
        key="dashboard_template_uploader",
    )

    template_df = st.session_state.template_df
    if uploaded_template is not None:
        try:
            template_df = pd.read_excel(uploaded_template, sheet_name="Master Sheet")
            st.session_state.template_df = template_df
        except Exception as e:
            st.error(f"Gagal membaca template: {e}")
            return

    if template_df is None:
        st.info("Silakan upload template terlebih dahulu di sini atau di menu 'Input Data'.")
        return

    # Validasi kategori yang ada di template
    site_name_col = template_df.columns[0]
    available_categories = [c for c in template_df.columns if c != site_name_col]
    if not available_categories:
        st.warning("Tidak ditemukan kolom kategori di 'Master Sheet' selain kolom pertama (nama perusahaan).")
        st.write("Header yang tersedia di template:")
        st.code("\n".join(list(map(str, template_df.columns))))
        return
    # Helpers
    def as_numeric(series: pd.Series) -> pd.Series:
        return pd.to_numeric(series, errors="coerce").fillna(0)

    # Tabs for different perspectives
    tab_overview, tab_top, tab_matrix, tab_profile = st.tabs([
        "Overview", "Top/Bottom", "Matrix", "Company Profile"
    ])

    with tab_overview:
        c1, c2 = st.columns([2, 1])
        with c2:
            category = st.selectbox("Pilih Kategori", options=available_categories, key="ov_cat")
        with c1:
            st.subheader("Ringkasan Kategori", anchor=False)

        df_cat = template_df[[site_name_col, category]].copy()
        df_cat[category] = as_numeric(df_cat[category])
        df_cat = df_cat.dropna(subset=[site_name_col])
        total_val = float(df_cat[category].sum())
        avg_val = float(df_cat[category].mean())
        max_row = df_cat.loc[df_cat[category].idxmax()] if not df_cat.empty else None
        min_row = df_cat.loc[df_cat[category].idxmin()] if not df_cat.empty else None

        def fmt_value(v: float) -> str:
            if v is None:
                return "-"
            # show two decimals for small numbers, else thousands separator
            if abs(v) < 1:
                return f"{v:,.2f}"
            # if integer-like, no decimals
            if float(v).is_integer():
                return f"{int(v):,}"
            return f"{v:,.2f}"

        def metric_card(title: str, value: float, subtitle: str = "", icon: str = "📊", tone: str = "primary"):
            html = f"""
            <div class='kpi-card'>
              <div class='kpi-head'>
                 <div class='kpi-icon {tone}'>{icon}</div>
                 <div>{title}</div>
                 <div class='kpi-badge'>Kategori</div>
              </div>
              <div class='kpi-value'>{fmt_value(value)}</div>
              <div class='kpi-sub'>{subtitle}</div>
              <div class='kpi-dot'></div>
            </div>
            """
            st.markdown(html, unsafe_allow_html=True)

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Total", total_val, subtitle=f"Jumlah nilai untuk {category}", icon="Σ", tone="primary")
        with k2:
            metric_card("Rata-rata", avg_val, subtitle=f"Rata-rata per perusahaan", icon="⌀", tone="slate")
        with k3:
            metric_card(
                "Maksimum",
                float(max_row[category]) if max_row is not None else 0,
                subtitle=f"Perusahaan: {str(max_row[site_name_col]) if max_row is not None else '-'}",
                icon="⬆",
                tone="success",
            )
        with k4:
            metric_card(
                "Minimum",
                float(min_row[category]) if min_row is not None else 0,
                subtitle=f"Perusahaan: {str(min_row[site_name_col]) if min_row is not None else '-'}",
                icon="⬇",
                tone="warning",
            )

        st.markdown("### Distribusi Nilai", help="Sebaran nilai pada kategori yang dipilih.")
        if PLOTLY_AVAILABLE:
            fig = px.histogram(df_cat, x=category, nbins=20, title="Histogram", template="simple_white")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.bar_chart(df_cat[category])

        st.markdown("### Sampel Data", help="5 nilai tertinggi dan terendah.")
        colA, colB = st.columns(2)
        with colA:
            st.caption("Top 5")
            st.table(df_cat.sort_values(category, ascending=False).head(5).reset_index(drop=True))
        with colB:
            st.caption("Bottom 5")
            st.table(df_cat.sort_values(category, ascending=True).head(5).reset_index(drop=True))

    with tab_top:
        left, right = st.columns([2, 1])
        with right:
            category_tb = st.selectbox("Kategori", options=available_categories, key="tb_cat")
            mode_rank = st.radio("Mode", ["Top", "Bottom"], horizontal=True)
            top_n = st.slider("Jumlah", min_value=5, max_value=30, value=10, step=1)
            include_zero = st.checkbox("Sertakan nilai 0", value=False)

        df_tb = template_df[[site_name_col, category_tb]].copy()
        df_tb[category_tb] = as_numeric(df_tb[category_tb])
        df_tb = df_tb.dropna(subset=[site_name_col])
        if not include_zero:
            df_tb = df_tb[df_tb[category_tb] != 0]
        df_tb = df_tb.sort_values(category_tb, ascending=(mode_rank == "Bottom")).head(top_n)

        with left:
            st.subheader(f"{mode_rank} {top_n} Perusahaan — {category_tb}")
            st.dataframe(df_tb.reset_index(drop=True))

        st.markdown("### Visualisasi")
        if PLOTLY_AVAILABLE:
            fig = px.bar(
                df_tb.sort_values(category_tb, ascending=True),
                x=category_tb,
                y=site_name_col,
                orientation="h",
                title=f"{mode_rank} {top_n} — {category_tb}",
                template="simple_white",
                height=500,
            )
            fig.update_layout(margin=dict(t=60, r=20, b=40, l=80))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.bar_chart(df_tb.set_index(site_name_col), height=500, use_container_width=True)

        # Download
        csv = df_tb.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Data (CSV)", data=csv, file_name="ranking.csv", mime="text/csv")

    with tab_matrix:
        st.subheader("Perbandingan Beberapa Kategori (Matrix)")
        selected_cats = st.multiselect(
            "Pilih sampai 6 kategori",
            options=available_categories,
            default=available_categories[:3] if len(available_categories) >= 3 else available_categories,
            max_selections=6,
        )
        top_for_matrix = st.slider("Top N perusahaan (berdasarkan total terpilih)", 5, 30, 10)

        if selected_cats:
            df_m = template_df[[site_name_col] + selected_cats].copy()
            for c in selected_cats:
                df_m[c] = as_numeric(df_m[c])
            df_m["__total__"] = df_m[selected_cats].sum(axis=1)
            df_m = df_m.sort_values("__total__", ascending=False).head(top_for_matrix)
            df_show = df_m.drop(columns=["__total__"]).set_index(site_name_col)

            if PLOTLY_AVAILABLE:
                fig = px.imshow(
                    df_show,
                    color_continuous_scale="Blues",
                    aspect="auto",
                    title="Heatmap Nilai",
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.dataframe(df_show)
        else:
            st.info("Pilih minimal satu kategori.")

    with tab_profile:
        st.subheader("Profil Perusahaan")
        companies = template_df[site_name_col].dropna().astype(str).unique().tolist()
        sel_company = st.selectbox("Pilih Perusahaan", options=companies)
        # Ambil semua kategori yang tersedia untuk profil
        prof_cats = [c for c in available_categories if c in template_df.columns]
        df_p = template_df[[site_name_col] + prof_cats].copy()
        for c in prof_cats:
            df_p[c] = as_numeric(df_p[c])
        row = df_p[df_p[site_name_col].astype(str) == str(sel_company)]
        if row.empty:
            st.warning("Data perusahaan tidak ditemukan.")
        else:
            row_vals = row[prof_cats].iloc[0]
            cA, cB = st.columns([1, 1])
            with cA:
                st.markdown(f"**Partner Name:** {sel_company}")
                st.table(pd.DataFrame({"Kategori": prof_cats, "Nilai": row_vals.values}))
            with cB:
                if PLOTLY_AVAILABLE and len(prof_cats) >= 3:
                    # Radar chart
                    plot_df = pd.DataFrame({"Kategori": prof_cats, "Nilai": row_vals.values})
                    fig = px.line_polar(plot_df, r="Nilai", theta="Kategori", line_close=True, template="simple_white")
                    fig.update_traces(fill='toself')
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.bar_chart(row_vals, use_container_width=True)


def render_input():
    st.header("🧾 Input Data & Proses")
    st.write("Aplikasi ini menghitung data dari satu file Excel dan memasukkannya ke dalam file template.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Upload File Sumber")
        source_file = st.file_uploader("Pilih file Excel yang berisi data mentah", type=["xlsx"], key="source_uploader")

    with col2:
        st.subheader("2. Upload File Template")
        template_file = st.file_uploader("Pilih file Excel template tujuan", type=["xlsx", "xlsm"], key="template_uploader")

    if source_file and template_file:
        try:
            source_df = pd.read_excel(source_file, header=1)
            source_df.columns = source_df.columns.str.strip()
            template_df = pd.read_excel(template_file, sheet_name="Master Sheet")

            st.session_state.template_df = template_df

            st.subheader("3. Atur Opsi Pemrosesan")
            target_column = st.selectbox(
                "Pilih kolom target di 'Master Sheet' untuk menempatkan hasil hitungan:",
                options=template_df.columns,
                key="target_column_select",
            )

            mode = st.radio(
                "Pilih mode pembaruan data:",
                options=["Add (Tambah)", "Replace (Ganti)"],
                help="Add: Menambahkan hasil hitungan ke nilai yang sudah ada. Replace: Mengganti nilai yang ada dengan hasil hitungan baru.",
                key="mode_radio",
            )

            if st.button("🚀 Proses Sekarang!", key="process_button"):
                with st.spinner("Sedang memproses data..."):
                    result_df = process_data(source_df, template_df, target_column, mode)
                    st.session_state.result_df = result_df
                    st.subheader("4. Hasil")
                    st.write("Data berhasil diproses. Berikut adalah pratinjau hasilnya:")
                    st.dataframe(result_df.fillna(''))

                    # Tulis hasil ke workbook template untuk menjaga format
                    template_file.seek(0)
                    wb = openpyxl.load_workbook(template_file, keep_vba=True)
                    ws = wb["Master Sheet"]

                    header = [cell.value for cell in ws[1]]
                    site_name_col = header[0]
                    site_name_to_row = {}
                    for row in range(2, ws.max_row + 1):
                        val = ws.cell(row=row, column=1).value
                        if val:
                            site_name_to_row[val] = row

                    for _, row_data in result_df.iterrows():
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

                    ext = ".xlsm" if template_file.name.lower().endswith(".xlsm") else ".xlsx"
                    st.session_state.last_processed_ext = ext
                    mime_type = (
                        "application/vnd.ms-excel.sheet.macroEnabled.12"
                        if ext == ".xlsm"
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="📥 Download File Hasil",
                        data=processed_data,
                        file_name=f"hasil_proses{ext}",
                        mime=mime_type,
                        key="download_button",
                    )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
            st.warning("Pastikan file Excel yang di-upload benar dan terdapat sheet bernama 'Master Sheet' di file template.")
    else:
        st.info("Silakan upload kedua file Excel untuk memulai.")


# --- Router ---
def render_guide():
    st.header("📘 Panduan Pengguna")
    st.write("Panduan lengkap untuk menggunakan aplikasi—fokus pada versi yang sudah dideploy di Streamlit Cloud.")

    st.success("Gunakan versi cloud: https://fillmastersheet.streamlit.app/ (klik link)")
    st.markdown("""
    ### Quick Start (Paling Cepat)
    1. Buka: [fillmastersheet.streamlit.app](https://fillmastersheet.streamlit.app/)
    2. Masuk ke menu "Input Data".
    3. Upload File Sumber (.xlsx) dan File Template (.xlsx/.xlsm).
    4. Pilih kolom target dan mode (Add/Replace), lalu klik "Proses Sekarang!".
    5. Download hasil (ekstensi mengikuti template, .xlsm tetap menyimpan macro).
    6. Pindah ke menu "Dashboard" untuk eksplorasi data: Overview, Top/Bottom, Matrix, dan Profil perusahaan.
    """)

    st.markdown("""
    ## 1. Ringkasan Aplikasi
    Aplikasi ini membantu Anda:
    - Menghitung jumlah keterlibatan/kemunculan per perusahaan (berdasarkan kolom "Site Name") dari file sumber Excel.
    - Memperbarui file template Excel pada sheet "Master Sheet" sesuai kolom target yang dipilih, dengan dua mode: Add (Tambah) dan Replace (Ganti).
    - Menjelajah hasil pada menu Dashboard: Overview, Top/Bottom, Matrix perbandingan kategori, dan Profil perusahaan.
    """)

    with st.expander("2. Persiapan Data (Wajib dibaca)", expanded=False):
        st.markdown("""
        ### 2.1. File Sumber (Raw Data)
        - Format: .xlsx
        - Header berada di baris ke-2 (kode menggunakan `header=1`).
        - Kolom minimal yang dibutuhkan:
          - `Student Code` dan `Course Code` (baris dengan nilai kosong di salah satu akan diabaikan)
          - `Site Name` (nama perusahaan/partner/instansi)
        - Kolom lain akan diabaikan oleh proses hitung, tetapi tetap boleh ada.

        ### 2.2. File Template
        - Format: .xlsx atau .xlsm (macro akan dipertahankan karena `keep_vba=True`).
        - Wajib memiliki sheet bernama `Master Sheet`.
        - Kolom pertama (kolom A) adalah nama perusahaan (misal: `Partner Name / Site Name / Company Name`).
        - Kolom berikutnya (B, C, dst) adalah kategori-kategori yang akan tampil di Dashboard. Nilai non-numerik (mis. `Y`) dianggap 0 pada visualisasi.
        - Judul/header dianggap berada pada baris pertama sheet.
        """)

    with st.expander("3. Mode Pemrosesan: Add vs Replace"):
        st.markdown("""
        - Add (Tambah): nilai pada kolom target akan ditambahkan dengan hasil hitungan baru.
          - Contoh: nilai lama 10, hasil hitung baru 3 → disimpan 13.
        - Replace (Ganti): nilai pada kolom target diganti total dengan hasil hitung baru.
          - Contoh: nilai lama 10, hasil hitung baru 3 → disimpan 3.
        - Jika perusahaan belum ada di template, baris baru akan ditambahkan otomatis.
        """)

    with st.expander("4. Langkah di Menu Input Data", expanded=True):
        st.markdown("""
        1) Upload File Sumber (.xlsx).
        2) Upload File Template (.xlsx atau .xlsm) yang memiliki sheet `Master Sheet`.
        3) Pilih kolom target (dari header `Master Sheet`) untuk menampung hasil hitungan.
        4) Pilih Mode (Add/Replace).
        5) Klik "Proses Sekarang!" dan tunggu hingga pratinjau hasil muncul.
        6) Unduh hasil melalui tombol Download. Ekstensi mengikuti file template (jika template .xlsm, hasil juga .xlsm, macro tetap aman).

        Catatan teknis saat menyimpan ke template:
        - Aplikasi membaca seluruh header pada baris pertama `Master Sheet`.
        - Mencari baris perusahaan berdasarkan isi kolom pertama.
        - Jika tidak ditemukan, menyisipkan baris baru di akhir dan mengisi seluruh kolom yang ada di header.
        """)

    with st.expander("5. Langkah di Menu Dashboard"):
        st.markdown("""
        - Upload (atau gunakan yang sudah di-upload dari menu Input Data) file template untuk dipakai sebagai sumber Dashboard.
        - Kategori yang tampil di Dashboard diambil otomatis dari seluruh header `Master Sheet` selain kolom pertama.

        5.1 Overview
        - KPI Cards: Total, Rata-rata, Maksimum, Minimum (ditampilkan dengan kartu modern).
        - Distribusi Nilai: histogram nilai kategori terpilih.
        - Sampel Data: Top 5 dan Bottom 5.

        5.2 Top/Bottom
        - Pilih mode Top atau Bottom, jumlah N, serta opsi sertakan nilai 0.
        - Tabel dan grafik bar horizontal tersedia.
        - Dapat diunduh sebagai CSV.

        5.3 Matrix
        - Pilih hingga 6 kategori untuk dibandingkan.
        - Pilih Top N perusahaan berdasarkan total penjumlahan nilai kategori terpilih.
        - Ditampilkan dalam bentuk heatmap (atau tabel jika Plotly tidak tersedia).

        5.4 Company Profile
        - Pilih satu perusahaan untuk melihat semua nilai kategori yang tersedia.
        - Visualisasi radar atau bar (fallback) akan ditampilkan.
        """)

    with st.expander("6. Tips, Batasan, dan Best Practice"):
        st.markdown("""
        - Nilai non-numerik pada kategori akan dianggap 0 di Dashboard. Jika ingin dihitung, konversikan terlebih dahulu (misal `Y` → 1).
        - Pastikan nama kolom persis (case-sensitive) terutama `Student Code`, `Course Code`, `Site Name` di file sumber.
        - Untuk performa pada file besar, simpan file ke disk lokal (bukan network drive) saat pemrosesan.
        - Hindari nama perusahaan duplikat dalam template; jika terjadi, sistem akan memperbarui baris pertama yang cocok.
        - Simpan backup template sebelum overwrite, terutama untuk file `.xlsm` yang memiliki macro penting.
        """)

    with st.expander("7. Troubleshooting (Masalah Umum)"):
        st.markdown("""
        - "Sheet 'Master Sheet' tidak ditemukan": pastikan nama sheet sesuai dan huruf kapital cocok.
        - "Kolom tidak ditemukan": cek ejaan header. Untuk file sumber, pastikan ada `Student Code`, `Course Code`, `Site Name`.
        - "Grafik kosong / semua 0": kemungkinan kolom kategori berisi teks non-numerik, atau filter menghapus semua baris.
        - "Tidak bisa download": pastikan ukuran file tidak terlalu besar dan browser mengizinkan unduhan.
        """)

    with st.expander("8. Catatan untuk Pengguna di Streamlit Cloud", expanded=True):
        st.markdown("""
        - Data Anda hanya digunakan pada sesi Anda; hindari mengunggah data sensitif jika tidak diperlukan.
        - Ukuran file upload mengikuti batas default Streamlit Cloud. Jika file besar, pertimbangkan untuk:
          - Menghapus sheet yang tidak perlu
          - Mengompresi/meringkas data sumber
        - Jika upload lama atau gagal, cek koneksi internet dan coba ulang.
        - Disarankan menggunakan browser modern (Chrome/Edge) versi terbaru.
        """)

    with st.expander("9. Jalankan di Komputer Sendiri (Opsional)"):
        st.markdown("""
        Perintah opsional jika ingin menjalankan secara manual di Windows PowerShell:
        ```powershell
        # (Opsional) Install dependensi sesuai proyek Anda
        # pip install -r requirements.txt

        # Jalankan aplikasi
        streamlit run "c:\\Users\\PRIMA\\OneDrive\\Documents\\PROJECT\\0 TRIAL\\Project Fill in Master sheet\\app.py"
        ```
        """)

    with st.expander("10. FAQ"):
        st.markdown("""
        - Apakah kategori harus ditentukan manual? Tidak. Kategori dibaca otomatis dari header `Master Sheet` (kecuali kolom pertama).
        - Apakah macro hilang saat menyimpan? Tidak, macro `.xlsm` dipertahankan (`keep_vba=True`).
        - Apakah bisa mengubah baris header sumber? Saat ini diasumsikan header di baris 2 (`header=1`). Jika berbeda, kode perlu disesuaikan.
        """)


if menu == "Dashboard":
    render_dashboard()
elif menu == "Input Data":
    render_input()
else:
    render_guide()
