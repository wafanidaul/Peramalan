import io
import streamlit as st
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Dashboard Peramalan PO Konveksi",
    page_icon="🧵",
    layout="wide"
)

NAMA_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

DAFTAR_PRODUK = ["WORKJACKET", "VARSITY", "COACH", "BOMBER", "LAINNYA"]


def hitung_tahun_bulan(periode: int):
    total_bulan = (2024 * 12 + 2) + (periode - 1)
    tahun = total_bulan // 12
    bulan = NAMA_BULAN[total_bulan % 12]
    return tahun, bulan


def inisialisasi_data():
    po_bulanan = [
        974, 924, 1622, 1301, 736, 710, 1611, 975, 1075, 972,
        1420, 1071, 1137, 1213, 1630, 1411, 877, 1657, 2175, 1539, 1135, 798,
        1207, 1231, 807, 1380, 886, 807, 942, 1052, 2043, 1728
    ]

    if "df_bulanan" not in st.session_state:
        rows = []
        for idx, val in enumerate(po_bulanan, start=1):
            thn, bln = hitung_tahun_bulan(idx)
            rows.append({
                "Tahun": thn,
                "Bulan": bln,
                "Periode": idx,
                "Jumlah PO": int(val),
                "Status Data": "Aktual"
            })
        st.session_state.df_bulanan = pd.DataFrame(rows)

    if "df_produk" not in st.session_state:
        workjacket = [
            281, 207, 575, 301, 193, 36, 345, 210, 0, 0,
            579, 313, 237, 862, 217, 455, 228, 259, 391, 511, 547, 339,
            687, 931, 265, 519, 329, 582, 720, 393, 1503, 1396
        ]
        varsity = [
            251, 544, 422, 818, 310, 560, 1203, 441, 594, 300,
            523, 658, 790, 351, 562, 597, 509, 312, 432, 254, 491, 303,
            333, 160, 100, 104, 204, 75, 70, 378, 100, 87
        ]
        coach = [
            330, 90, 573, 134, 233, 101, 26, 151, 0, 122,
            37, 64, 110, 0, 0, 0, 0, 0, 89, 197, 97, 71,
            73, 24, 0, 15, 0, 0, 22, 150, 82, 1
        ]
        bomber = [
            112, 83, 52, 48, 0, 13, 37, 173, 481, 550,
            281, 36, 0, 0, 851, 359, 140, 1086, 1263, 777, 0, 85,
            114, 116, 442, 742, 353, 150, 130, 131, 358, 244
        ]
        lainnya = [
            max(0, tot - (wj + vr + ch + bm))
            for tot, wj, vr, ch, bm in zip(po_bulanan, workjacket, varsity, coach, bomber)
        ]

        rows_prod = []
        for idx in range(1, len(po_bulanan) + 1):
            thn, bln = hitung_tahun_bulan(idx)
            i = idx - 1
            rows_prod.append({
                "Tahun": thn,
                "Bulan": bln,
                "Periode": idx,
                "WORKJACKET": workjacket[i],
                "VARSITY": varsity[i],
                "COACH": coach[i],
                "BOMBER": bomber[i],
                "LAINNYA": lainnya[i],
                "Status Data": "Aktual"
            })
        st.session_state.df_produk = pd.DataFrame(rows_prod)


def jalankan_regresi(df_aktual: pd.DataFrame, kolom_target: str, jumlah_prediksi: int):
    X = df_aktual[["Periode"]].values
    y = df_aktual[kolom_target].values

    model = LinearRegression()
    model.fit(X, y)

    y_pred_aktual = model.predict(X)
    r2 = r2_score(y, y_pred_aktual)
    mae = mean_absolute_error(y, y_pred_aktual)
    rmse = np.sqrt(mean_squared_error(y, y_pred_aktual))

    periode_terakhir = int(df_aktual["Periode"].max())
    rows_pred = []
    for step in range(1, jumlah_prediksi + 1):
        p_baru = periode_terakhir + step
        thn, bln = hitung_tahun_bulan(p_baru)
        pred_val = max(0, round(float(model.predict([[p_baru]])[0])))
        rows_pred.append({
            "Tahun": thn,
            "Bulan": bln,
            "Periode": p_baru,
            kolom_target: pred_val,
            "Status Data": "Prediksi"
        })

    df_pred = pd.DataFrame(rows_pred)
    df_gabungan = pd.concat(
        [df_aktual[["Tahun", "Bulan", "Periode", kolom_target, "Status Data"]], df_pred],
        ignore_index=True
    )
    df_gabungan["Garis Tren"] = np.maximum(0, np.round(model.predict(df_gabungan[["Periode"]].values), 1))

    metrik = {
        "intercept": model.intercept_,
        "slope": model.coef_[0],
        "r2": r2,
        "mae": mae,
        "rmse": rmse
    }
    return df_gabungan, metrik


def buat_kesimpulan_analisis(df_gabungan: pd.DataFrame, metrik: dict, kolom_target: str, label_objek: str):
    df_aktual = df_gabungan[df_gabungan["Status Data"] == "Aktual"]
    df_pred = df_gabungan[df_gabungan["Status Data"] == "Prediksi"]

    slope = metrik["slope"]
    intercept = metrik["intercept"]
    rata_aktual = df_aktual[kolom_target].mean()
    max_row = df_aktual.loc[df_aktual[kolom_target].idxmax()]
    min_row = df_aktual.loc[df_aktual[kolom_target].idxmin()]

    if slope > 5:
        status_tren = "mengalami tren **kenaikan (positif)**"
        rekomendasi = (
            "Kapasitas produksi diproyeksikan meningkat. Disarankan mulai mengamankan stok bahan baku utama, "
            "menjadwalkan kapasitas mesin jahit, serta mengantisipasi kebutuhan tenaga kerja tambahan."
        )
    elif slope < -5:
        status_tren = "mengalami tren **penurunan (negatif)**"
        rekomendasi = (
            "Permintaan cenderung melandai. Disarankan mengendalikan pembelian stok kain agar tidak menumpuk (overstock) "
            "serta mendorong promosi atau pengembangan desain artikel baru."
        )
    else:
        status_tren = "cenderung **stabil (mendatar)**"
        rekomendasi = (
            "Volume pesanan relatif stabil. Fokuskan strategi pada efisiensi biaya produksi, pengurangan limbah kain, "
            "dan menjaga ketepatan waktu pengiriman (lead time)."
        )

    poin_kesimpulan = [
        f"**Persamaan & Arah Tren:** Model regresi menghasilkan persamaan $Y = {intercept:,.2f} + ({slope:,.2f})X$. Secara historis, jumlah {label_objek} {status_tren} sebesar **{abs(slope):,.1f} pcs per periode**.",
        f"**Ringkasan Data Historis ({len(df_aktual)} Periode):** Rata-rata {label_objek} sebesar **{rata_aktual:,.0f} pcs/bulan**, dengan puncak tertinggi pada **{max_row['Bulan']} {max_row['Tahun']} ({max_row[kolom_target]:,.0f} pcs)** dan terendah pada **{min_row['Bulan']} {min_row['Tahun']} ({min_row[kolom_target]:,.0f} pcs)**."
    ]

    if not df_pred.empty:
        rata_pred = df_pred[kolom_target].mean()
        total_pred = df_pred[kolom_target].sum()
        p_awal = f"{df_pred['Bulan'].iloc[0]} {df_pred['Tahun'].iloc[0]}"
        p_akhir = f"{df_pred['Bulan'].iloc[-1]} {df_pred['Tahun'].iloc[-1]}"
        persen_perubahan = ((rata_pred - rata_aktual) / rata_aktual * 100) if rata_aktual > 0 else 0

        poin_kesimpulan.append(
            f"**Proyeksi Peramalan ({len(df_pred)} Periode ke Depan):** Untuk periode **{p_awal} – {p_akhir}**, total kebutuhan diprediksi mencapai **{total_pred:,.0f} pcs** dengan rata-rata **{rata_pred:,.0f} pcs/bulan** (**{persen_perubahan:+.1f}%** dibandingkan rata-rata historis)."
        )

    poin_kesimpulan.append(f"**Rekomendasi Operasional:** {rekomendasi}")
    return poin_kesimpulan


def buat_file_excel(df: pd.DataFrame, nama_sheet: str, kolom_angka: list, daftar_kesimpulan: list) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nama_sheet)
        ws = writer.sheets[nama_sheet]

        header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        pred_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
        pred_font = Font(color="1E293B", bold=True)
        thin_side = Side(border_style="thin", color="CBD5E1")
        border_style = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_style
        ws.row_dimensions[1].height = 24

        idx_status = df.columns.get_loc("Status Data") + 1
        for row_idx in range(2, len(df) + 2):
            is_prediksi = ws.cell(row=row_idx, column=idx_status).value == "Prediksi"
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = border_style

                if is_prediksi:
                    cell.fill = pred_fill
                    cell.font = pred_font

                if col_name in kolom_angka:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name in ["Tahun", "Periode", "Status Data"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                df[col_name].astype(str).map(len).max(),
                len(str(col_name))
            ) + 4
            ws.column_dimensions[col_letter].width = max_len

        ws.freeze_panes = "A2"

        ws_kes = writer.book.create_sheet(title="Kesimpulan Analisis")
        ws_kes["A1"] = "RINGKASAN & KESIMPULAN HASIL PERAMALAN PO"
        ws_kes["A1"].font = Font(size=13, bold=True, color="1E3A5F")
        ws_kes.row_dimensions[1].height = 26

        for i, teks in enumerate(daftar_kesimpulan, start=3):
            teks_bersih = teks.replace("**", "").replace("$", "")
            cell_no = ws_kes.cell(row=i, column=1, value=f"Poin {i-2}")
            cell_no.font = Font(bold=True)
            cell_no.alignment = Alignment(vertical="top")

            cell_isi = ws_kes.cell(row=i, column=2, value=teks_bersih)
            cell_isi.alignment = Alignment(wrap_text=True, vertical="top")
            ws_kes.row_dimensions[i].height = 42

        ws_kes.column_dimensions["A"].width = 12
        ws_kes.column_dimensions["B"].width = 95

    return output.getvalue()


def buat_grafik_peramalan(df_gabungan: pd.DataFrame, kolom_target: str, judul: str):
    df_aktual = df_gabungan[df_gabungan["Status Data"] == "Aktual"].copy()
    df_pred = df_gabungan[df_gabungan["Status Data"] == "Prediksi"].copy()

    df_gabungan["Label Waktu"] = df_gabungan["Bulan"].str[:3] + " " + df_gabungan["Tahun"].astype(str)
    df_aktual["Label Waktu"] = df_aktual["Bulan"].str[:3] + " " + df_aktual["Tahun"].astype(str)
    df_pred["Label Waktu"] = df_pred["Bulan"].str[:3] + " " + df_pred["Tahun"].astype(str)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df_aktual["Label Waktu"],
        y=df_aktual[kolom_target],
        mode="lines+markers",
        name="Data Aktual",
        line=dict(color="#1E3A5F", width=2.5),
        marker=dict(size=7)
    ))

    if not df_pred.empty:
        df_sambung = pd.concat([df_aktual.tail(1), df_pred], ignore_index=True)
        fig.add_trace(go.Scatter(
            x=df_sambung["Label Waktu"],
            y=df_sambung[kolom_target],
            mode="lines+markers",
            name="Hasil Prediksi",
            line=dict(color="#EAB308", width=3, dash="dot"),
            marker=dict(size=8, symbol="diamond")
        ))

    fig.add_trace(go.Scatter(
        x=df_gabungan["Label Waktu"],
        y=df_gabungan["Garis Tren"],
        mode="lines",
        name="Garis Tren Regresi",
        line=dict(color="#EF4444", width=2, dash="dash")
    ))

    fig.update_layout(
        title=judul,
        xaxis_title="Periode (Bulan & Tahun)",
        yaxis_title="Jumlah PO (pcs)",
        hovermode="x unified",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=430
    )
    return fig


inisialisasi_data()

st.title("🧵 Dashboard Peramalan Purchase Order (PO) Konveksi")
st.caption("Metode Regresi Linear Sederhana — Monitoring Data Historis, Proyeksi Kebutuhan Produksi, & Kesimpulan Otomatis")

with st.sidebar:
    st.header("⚙️ Pengaturan Peramalan")
    jumlah_prediksi = st.slider(
        "Jumlah Bulan Prediksi ke Depan",
        min_value=1,
        max_value=12,
        value=6
    )

    st.divider()
    if st.button("🔄 Reset ke Data Awal", use_container_width=True):
        del st.session_state["df_bulanan"]
        del st.session_state["df_produk"]
        st.rerun()

tab_total, tab_produk, tab_edit = st.tabs([
    "📈 Peramalan Total PO",
    "🧥 Peramalan Per Produk",
    "📝 Kelola & Edit Data"
])

with tab_total:
    df_hasil_total, metrik_total = jalankan_regresi(
        st.session_state.df_bulanan,
        kolom_target="Jumlah PO",
        jumlah_prediksi=jumlah_prediksi
    )
    kesimpulan_total = buat_kesimpulan_analisis(
        df_hasil_total, metrik_total, kolom_target="Jumlah PO", label_objek="Total PO"
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Slope (Tren / Bulan)", f"{metrik_total['slope']:+.2f} pcs")
    c2.metric("Intercept (a)", f"{metrik_total['intercept']:,.2f}")
    c3.metric("R-Squared (R²)", f"{metrik_total['r2']:.4f}")
    c4.metric("MAE (Rata-rata Error)", f"{metrik_total['mae']:,.1f} pcs")

    fig_total = buat_grafik_peramalan(
        df_hasil_total,
        kolom_target="Jumlah PO",
        judul="Grafik Aktual vs Prediksi Total PO Bulanan"
    )
    st.plotly_chart(fig_total, use_container_width=True)

    with st.container(border=True):
        st.subheader("📌 Kesimpulan & Rekomendasi Otomatis")
        for poin in kesimpulan_total:
            st.markdown(f"- {poin}")

    col_tbl, col_dl = st.columns([3, 1])
    with col_tbl:
        st.subheader("📋 Tabel Data Aktual & Hasil Prediksi")
    with col_dl:
        excel_bytes = buat_file_excel(
            df_hasil_total,
            nama_sheet="Peramalan_Total_PO",
            kolom_angka=["Jumlah PO", "Garis Tren"],
            daftar_kesimpulan=kesimpulan_total
        )
        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=excel_bytes,
            file_name="Hasil_Peramalan_Total_PO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.dataframe(
        df_hasil_total.style.apply(
            lambda r: ["background-color: #FEF9C3; font-weight: bold" if r["Status Data"] == "Prediksi" else "" for _ in r],
            axis=1
        ),
        use_container_width=True,
        hide_index=True
    )

with tab_produk:
    produk_pilihan = st.selectbox("Pilih Kategori Produk:", DAFTAR_PRODUK)

    df_hasil_prod, metrik_prod = jalankan_regresi(
        st.session_state.df_produk,
        kolom_target=produk_pilihan,
        jumlah_prediksi=jumlah_prediksi
    )
    kesimpulan_prod = buat_kesimpulan_analisis(
        df_hasil_prod, metrik_prod, kolom_target=produk_pilihan, label_objek=f"PO {produk_pilihan}"
    )

    p1, p2, p3, p4 = st.columns(4)
    p1.metric(f"Tren {produk_pilihan}", f"{metrik_prod['slope']:+.2f} pcs/bln")
    p2.metric("Intercept (a)", f"{metrik_prod['intercept']:,.2f}")
    p3.metric("R-Squared (R²)", f"{metrik_prod['r2']:.4f}")
    p4.metric("MAE", f"{metrik_prod['mae']:,.1f} pcs")

    fig_prod = buat_grafik_peramalan(
        df_hasil_prod,
        kolom_target=produk_pilihan,
        judul=f"Grafik Aktual vs Prediksi PO — {produk_pilihan}"
    )
    st.plotly_chart(fig_prod, use_container_width=True)

    with st.container(border=True):
        st.subheader(f"📌 Kesimpulan Khusus Produk: {produk_pilihan}")
        for poin in kesimpulan_prod:
            st.markdown(f"- {poin}")

    df_akt_prod = st.session_state.df_produk
    total_tiap_produk = {p: int(df_akt_prod[p].sum()) for p in DAFTAR_PRODUK}
    terlaris = max(total_tiap_produk, key=total_tiap_produk.get)

    slopes_prod = {}
    for p in DAFTAR_PRODUK:
        _, m_tmp = jalankan_regresi(df_akt_prod, p, 1)
        slopes_prod[p] = m_tmp["slope"]
    tumbuh_tercepat = max(slopes_prod, key=slopes_prod.get)

    st.info(
        f"**Perbandingan Lini Produk Keseluruhan:** Produk dengan volume kumulatif terbesar adalah **{terlaris}** "
        f"(**{total_tiap_produk[terlaris]:,} pcs**), sedangkan produk dengan laju pertumbuhan tren tercepat adalah "
        f"**{tumbuh_tercepat}** (**{slopes_prod[tumbuh_tercepat]:+.2f} pcs/periode**)."
    )

    excel_prod_bytes = buat_file_excel(
        df_hasil_prod,
        nama_sheet=f"Peramalan_{produk_pilihan}",
        kolom_angka=[produk_pilihan, "Garis Tren"],
        daftar_kesimpulan=kesimpulan_prod
    )
    st.download_button(
        label=f"📥 Download Excel {produk_pilihan} (.xlsx)",
        data=excel_prod_bytes,
        file_name=f"Hasil_Peramalan_{produk_pilihan}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with tab_edit:
    st.subheader("✏️ Edit Data Aktual atau Tambah Periode Baru")
    st.write("Ubah angka pada tabel di bawah ini, lalu klik **Simpan Perubahan** untuk memperbarui grafik dan kesimpulan secara otomatis.")

    df_editor = st.data_editor(
        st.session_state.df_produk,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        disabled=["Tahun", "Bulan", "Periode", "Status Data"]
    )

    if st.button("💾 Simpan Perubahan Data", type="primary"):
        df_baru = df_editor.copy()
        for i in range(len(df_baru)):
            periode_ke = i + 1
            thn, bln = hitung_tahun_bulan(periode_ke)
            df_baru.loc[i, "Periode"] = periode_ke
            df_baru.loc[i, "Tahun"] = thn
            df_baru.loc[i, "Bulan"] = bln
            df_baru.loc[i, "Status Data"] = "Aktual"
            for p in DAFTAR_PRODUK:
                if pd.isna(df_baru.loc[i, p]):
                    df_baru.loc[i, p] = 0
                df_baru.loc[i, p] = int(df_baru.loc[i, p])

        st.session_state.df_produk = df_baru

        df_bulanan_baru = df_baru[["Tahun", "Bulan", "Periode", "Status Data"]].copy()
        df_bulanan_baru["Jumlah PO"] = df_baru[DAFTAR_PRODUK].sum(axis=1).astype(int)
        df_bulanan_baru = df_bulanan_baru[["Tahun", "Bulan", "Periode", "Jumlah PO", "Status Data"]]
        st.session_state.df_bulanan = df_bulanan_baru

        st.success("✅ Data berhasil diperbarui! Silakan cek Tab Peramalan untuk melihat hasil & kesimpulan terbaru.")
        st.rerun()
