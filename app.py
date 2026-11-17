import io
import streamlit as st
import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard Peramalan PO Konveksi", layout="wide")


def buat_file_excel(df: pd.DataFrame, nama_sheet: str, kolom_angka: list, daftar_kesimpulan: list = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nama_sheet)
        ws = writer.sheets[nama_sheet]

        # Definisi Style
        header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        pred_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
        pred_font = Font(color="1E293B", bold=True)
        thin_side = Side(border_style="thin", color="CBD5E1")
        border_style = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # Styling Header (Baris 1)
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_style
        ws.row_dimensions[1].height = 24

        # Styling Baris Data & Highlight Baris Prediksi
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

        # Auto-adjust lebar kolom & Freeze header
        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                df[col_name].astype(str).map(len).max(),
                len(str(col_name))
            ) + 4
            ws.column_dimensions[col_letter].width = max_len

        ws.freeze_panes = "A2"

        # Tambahkan Sheet Kesimpulan jika tersedia
        if daftar_kesimpulan:
            ws_kes = writer.book.create_sheet(title="Kesimpulan Analisis")
            ws_kes["A1"] = "RINGKASAN & KESIMPULAN ANALISIS PERAMALAN"
            ws_kes["A1"].font = Font(size=13, bold=True, color="1E3A5F")
            ws_kes.row_dimensions[1].height = 25

            for idx, poin in enumerate(daftar_kesimpulan, start=3):
                teks_bersih = poin.replace("**", "").replace("$", "")
                c_no = ws_kes.cell(row=idx, column=1, value=f"Poin {idx - 2}")
                c_no.font = Font(bold=True)
                c_no.alignment = Alignment(vertical="top")

                c_isi = ws_kes.cell(row=idx, column=2, value=teks_bersih)
                c_isi.alignment = Alignment(wrap_text=True, vertical="top")
                ws_kes.row_dimensions[idx].height = 42

            ws_kes.column_dimensions["A"].width = 12
            ws_kes.column_dimensions["B"].width = 95

    return output.getvalue()


NAMA_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

DAFTAR_PRODUK = ["WORKJACKET", "VARSITY", "COACH", "BOMBER", "LAINNYA"]


def hitung_tahun_bulan(periode: int):
    # Periode 1 = Maret 2024 (index bulan = 2 dari 0-based)
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


def hitung_regresi_linear(df_aktual: pd.DataFrame, kolom_target: str, bulan_prediksi: int):
    x = df_aktual["Periode"].to_numpy(dtype=float)
    y = df_aktual[kolom_target].to_numpy(dtype=float)

    # Regresi Linear Sederhana: Y = a + bX
    slope, intercept = np.polyfit(x, y, 1)
    y_hat = intercept + slope * x

    # Evaluasi Model
    mae = float(np.mean(np.abs(y - y_hat)))
    rmse = float(np.sqrt(np.mean((y - y_hat) ** 2)))
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1.0 - (ss_res / ss_tot) if ss_tot > 0 else 0.0

    # Proyeksi periode masa depan
    periode_terakhir = int(df_aktual["Periode"].max())
    rows_pred = []
    for step in range(1, bulan_prediksi + 1):
        p_baru = periode_terakhir + step
        thn, bln = hitung_tahun_bulan(p_baru)
        nilai_pred = max(0, int(round(intercept + slope * p_baru)))
        rows_pred.append({
            "Tahun": thn,
            "Bulan": bln,
            "Periode": p_baru,
            kolom_target: nilai_pred,
            "Status Data": "Prediksi"
        })

    df_pred = pd.DataFrame(rows_pred)
    df_hasil = pd.concat(
        [df_aktual[["Tahun", "Bulan", "Periode", kolom_target, "Status Data"]], df_pred],
        ignore_index=True
    )
    df_hasil["Garis Tren"] = np.maximum(0, np.round(intercept + slope * df_hasil["Periode"], 1))

    info_model = {
        "intercept": intercept,
        "slope": slope,
        "r2": r2,
        "mae": mae,
        "rmse": rmse
    }
    return df_hasil, info_model


def buat_penjelasan_kesimpulan(df_hasil: pd.DataFrame, info_model: dict, kolom_target: str, nama_objek: str):
    df_akt = df_hasil[df_hasil["Status Data"] == "Aktual"]
    df_prd = df_hasil[df_hasil["Status Data"] == "Prediksi"]

    slope = info_model["slope"]
    intercept = info_model["intercept"]
    rata_akt = df_akt[kolom_target].mean()
    total_akt = df_akt[kolom_target].sum()

    row_max = df_akt.loc[df_akt[kolom_target].idxmax()]
    row_min = df_akt.loc[df_akt[kolom_target].idxmin()]

    if slope > 5:
        arah = "mengalami tren **kenaikan (positif)**"
        saran = (
            "Kebutuhan produksi diproyeksikan terus bertumbuh. Disarankan menyiapkan stok bahan baku kain & aksesoris "
            "lebih awal, mengatur kapasitas mesin jahit, serta mengantisipasi jadwal lembur/penambahan penjahit."
        )
    elif slope < -5:
        arah = "mengalami tren **penurunan (negatif)**"
        saran = (
            "Permintaan cenderung melandai. Disarankan mengendalikan pembelian stok bahan baku agar tidak terjadi "
            "penumpukan (overstock) serta memfokuskan pemasaran pada artikel yang masih tinggi peminatnya."
        )
    else:
        arah = "cenderung **stabil (mendatar)**"
        saran = (
            "Volume pesanan relatif stabil dari bulan ke bulan. Fokuskan strategi pada efisiensi biaya produksi, "
            "pengendalian kualitas (QC), dan menjaga ketepatan waktu pengiriman."
        )

    kesimpulan = [
        f"**Persamaan Regresi & Arah Tren:** Model menghasilkan persamaan $Y = {intercept:,.2f} + ({slope:,.2f})X$. Secara keseluruhan, {nama_objek} {arah} dengan perubahan rata-rata **{abs(slope):,.1f} pcs per periode**.",
        f"**Analisis Data Historis ({len(df_akt)} Bulan):** Total pesanan tercatat sebesar **{total_akt:,.0f} pcs** dengan rata-rata **{rata_akt:,.0f} pcs/bulan**. Permintaan tertinggi terjadi pada **{row_max['Bulan']} {row_max['Tahun']} ({row_max[kolom_target]:,.0f} pcs)** dan terendah pada **{row_min['Bulan']} {row_min['Tahun']} ({row_min[kolom_target]:,.0f} pcs)**."
    ]

    if not df_prd.empty:
        rata_prd = df_prd[kolom_target].mean()
        total_prd = df_prd[kolom_target].sum()
        awal_prd = f"{df_prd['Bulan'].iloc[0]} {df_prd['Tahun'].iloc[0]}"
        akhir_prd = f"{df_prd['Bulan'].iloc[-1]} {df_prd['Tahun'].iloc[-1]}"
        selisih_pct = ((rata_prd - rata_akt) / rata_akt * 100) if rata_akt > 0 else 0

        kesimpulan.append(
            f"**Proyeksi Peramalan ({len(df_prd)} Bulan ke Depan):** Pada periode **{awal_prd} – {akhir_prd}**, total {nama_objek} diperkirakan mencapai **{total_prd:,.0f} pcs** dengan rata-rata **{rata_prd:,.0f} pcs/bulan** (**{selisih_pct:+.1f}%** dibandingkan rata-rata historis)."
        )

    kesimpulan.append(f"**Kesimpulan & Rekomendasi Keputusan:** {saran}")
    return kesimpulan


# ==========================================
# TAMPILAN UTAMA DASHBOARD
# ==========================================
inisialisasi_data()

st.title("Dashboard Peramalan Purchase Order (PO) Konveksi")
st.caption("Analisis Tren Historis & Proyeksi Kebutuhan Produksi Menggunakan Metode Regresi Linear Sederhana")

with st.sidebar:
    st.header("Pengaturan Peramalan")
    jumlah_prediksi = st.slider("Jumlah Periode Prediksi (Bulan)", min_value=1, max_value=12, value=6)
    st.divider()
    if st.button("Reset ke Data Awal", use_container_width=True):
        del st.session_state["df_bulanan"]
        del st.session_state["df_produk"]
        st.rerun()

tab1, tab2, tab3 = st.tabs([
    "Peramalan Total PO",
    "Peramalan Per Produk",
    "Kelola & Edit Data"
])

# --- TAB 1: TOTAL PO ---
with tab1:
    df_hasil_total, info_total = hitung_regresi_linear(
        st.session_state.df_bulanan,
        kolom_target="Jumlah PO",
        bulan_prediksi=jumlah_prediksi
    )
    kesimpulan_total = buat_penjelasan_kesimpulan(
        df_hasil_total, info_total, kolom_target="Jumlah PO", nama_objek="Total PO"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Slope (Tren / Bulan)", f"{info_total['slope']:+.2f} pcs")
    m2.metric("Intercept (a)", f"{info_total['intercept']:,.2f}")
    m3.metric("R-Squared (R²)", f"{info_total['r2']:.4f}")
    m4.metric("MAE (Rata-rata Error)", f"{info_total['mae']:,.1f} pcs")

    # Grafik Tren & Peramalan
    st.subheader("Grafik Tren & Hasil Peramalan Total PO")
    df_chart_total = df_hasil_total.copy()
    df_chart_total["Periode Label"] = df_chart_total["Periode"].astype(str).str.zfill(2) + " - " + df_chart_total["Bulan"].str[:3] + " " + df_chart_total["Tahun"].astype(str)
    df_chart_total["Data Aktual"] = np.where(df_chart_total["Status Data"] == "Aktual", df_chart_total["Jumlah PO"], np.nan)
    df_chart_total["Data Prediksi"] = np.where(df_chart_total["Status Data"] == "Prediksi", df_chart_total["Jumlah PO"], np.nan)
    st.line_chart(
        df_chart_total.set_index("Periode Label")[["Data Aktual", "Data Prediksi", "Garis Tren"]],
        height=380
    )

    # Panel Penjelasan & Kesimpulan
    with st.container(border=True):
        st.subheader("Penjelasan & Kesimpulan Analisis (Total PO)")
        for poin in kesimpulan_total:
            st.markdown(f"- {poin}")

    col_kiri, col_kanan = st.columns([3, 1])
    with col_kiri:
        st.subheader("Tabel Data Aktual & Prediksi")
    with col_kanan:
        excel_total = buat_file_excel(
            df_hasil_total,
            nama_sheet="Peramalan_Total_PO",
            kolom_angka=["Jumlah PO", "Garis Tren"],
            daftar_kesimpulan=kesimpulan_total
        )
        st.download_button(
            label="Download Excel (.xlsx)",
            data=excel_total,
            file_name="Peramalan_Total_PO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.dataframe(
        df_hasil_total.style.apply(
