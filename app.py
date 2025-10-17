
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

menu = st.sidebar.radio("Menu", ["Dashboard", "Input Data"], index=1)

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
if menu == "Dashboard":
    render_dashboard()
else:
    render_input()
