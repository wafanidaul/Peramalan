import io
import streamlit as st
import pandas as pd
import numpy as np
from statsmodels.tsa.ar_model import AutoReg, ar_select_order
from statsmodels.tsa.api import VAR
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard Peramalan PO Konveksi", layout="wide")

NAMA_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

def hitung_tahun_bulan(periode: int):
    total_bulan = (2024 * 12 + 2) + (periode - 1)
    tahun = total_bulan // 12
    bulan = NAMA_BULAN[total_bulan % 12]
    return tahun, bulan

def inisialisasi_data():
    if "df_bulanan" not in st.session_state:
        po_bulanan = [
            974, 924, 1622, 1301, 736, 710, 1611, 975, 1075, 972,
            1420, 1071, 1137, 1213, 1630, 1411, 877, 1657, 2175, 1539, 1135, 798,
            1207, 1231, 807, 1380, 886, 807, 942, 1052, 2043, 1728
        ]
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
        data_prod = {
            "WORKJACKET": [
                281, 207, 575, 301, 193, 36, 345, 210, 0, 0,
                579, 313, 237, 862, 217, 455, 228, 259, 391, 511, 547, 339,
                687, 931, 265, 519, 329, 582, 720, 393, 1503, 1396
            ],
            "VARSITY": [
                251, 544, 422, 818, 310, 560, 1203, 441, 594, 300,
                523, 658, 790, 351, 562, 597, 509, 312, 432, 254, 491, 303,
                333, 160, 100, 104, 204, 75, 70, 378, 100, 87
            ],
            "COACH": [
                330, 90, 573, 134, 233, 101, 26, 151, 0, 122,
                37, 64, 110, 0, 0, 0, 0, 0, 89, 197, 97, 71,
                73, 24, 0, 15, 0, 0, 22, 150, 82, 137
            ]
        }
        rows = []
        for idx in range(1, 33):
            thn, bln = hitung_tahun_bulan(idx)
            rows.append({
                "Tahun": thn,
                "Bulan": bln,
                "Periode": idx,
                "WORKJACKET": data_prod["WORKJACKET"][idx - 1],
                "VARSITY": data_prod["VARSITY"][idx - 1],
                "COACH": data_prod["COACH"][idx - 1],
                "Status Data": "Aktual"
            })
        st.session_state.df_produk = pd.DataFrame(rows)
        st.session_state.daftar_produk = ["WORKJACKET", "VARSITY", "COACH"]

    if "mode_edit_bulanan" not in st.session_state:
        st.session_state.mode_edit_bulanan = False
    if "mode_edit_produk" not in st.session_state:
        st.session_state.mode_edit_produk = False

def prediksi_autoreg_series(y_series: np.ndarray, langkah: int = 3) -> np.ndarray:
    y = np.asarray(y_series, dtype=float)
    n = len(y)

    if n == 0:
        return np.zeros(langkah, dtype=int)
    if n < 4 or np.allclose(y, y[0]):
        return np.maximum(0, np.round(np.repeat(y[-1], langkah))).astype(int)

    max_lag = min(6, max(1, (n - 1) // 2))
    try:
        sel = ar_select_order(y, maxlag=max_lag, ic="aic", trend="c")
        lags = sel.ar_lags if sel.ar_lags else [1]
        model = AutoReg(y, lags=lags, trend="c")
        fit = model.fit()
        pred = fit.predict(start=n, end=n + langkah - 1)
    except Exception:
        fallback_lag = min(1, n - 1)
        fit = AutoReg(y, lags=fallback_lag, trend="c").fit()
        pred = fit.predict(start=n, end=n + langkah - 1)

    return np.maximum(0, np.round(pred)).astype(int)

def prediksi_var_multivariate(df_urut: pd.DataFrame, kolom_target: list, langkah: int = 3):
    """Memprediksi semua kolom produk secara simultan dengan VAR, dengan fallback AR jika gagal."""
    k = len(kolom_target)
    n = len(df_urut)

    # Jika hanya 1 produk, VAR identik dengan AR univariate
    if k == 1:
        col = kolom_target[0]
        y = df_urut[col].fillna(0).values
        return {col: prediksi_autoreg_series(y, langkah=langkah)}, "AR(p) (1 produk)"

    # Butuhkan cukup observasi agar parameter VAR tidak over-parameterized
    if n < max(6, k * 2 + 2):
        hasil = {
            col: prediksi_autoreg_series(df_urut[col].fillna(0).values, langkah=langkah)
            for col in kolom_target
        }
        return hasil, "Fallback AR(p) (observasi terbatas)"

    df_mv = df_urut[kolom_target].fillna(0).astype(float).copy()
    max_lag = min(4, max(1, (n - 2) // (k + 1)))

    try:
        model = VAR(df_mv)
        sel = model.select_order(maxlags=max_lag)
        p = sel.aic if sel.aic and sel.aic >= 1 else 1
        fit = model.fit(p)
        y_input = df_mv.values[-p:]
        fc = fit.forecast(y=y_input, steps=langkah)

        hasil = {}
        for idx, col in enumerate(kolom_target):
            hasil[col] = np.maximum(0, np.round(fc[:, idx])).astype(int)
        return hasil, f"VAR({p}) (AIC)"
    except Exception:
        hasil = {
            col: prediksi_autoreg_series(df_urut[col].fillna(0).values, langkah=langkah)
            for col in kolom_target
        }
        return hasil, "Fallback AR(p) (matriks VAR singular/tidak stabil)"

def buat_prediksi_3_bulan(df_aktual: pd.DataFrame, kolom_target: list, metode: str = "AR"):
    df_urut = df_aktual.sort_values("Periode").reset_index(drop=True)
    periode_terakhir = int(df_urut["Periode"].max())
    periode_depan = np.arange(periode_terakhir + 1, periode_terakhir + 4)

    rows_pred = []
    for p in periode_depan:
        thn, bln = hitung_tahun_bulan(int(p))
        rows_pred.append({
            "Tahun": thn,
            "Bulan": bln,
            "Periode": int(p),
            "Status Data": "Prediksi"
        })

    df_pred = pd.DataFrame(rows_pred)
    info_model = ""

    if metode == "VAR":
        hasil_dict, info_model = prediksi_var_multivariate(df_urut, kolom_target, langkah=3)
        for col in kolom_target:
            df_pred[col] = hasil_dict[col]
    else:
        for col in kolom_target:
            y = df_urut[col].fillna(0).values
            df_pred[col] = prediksi_autoreg_series(y, langkah=3)
        info_model = "AR(p) per kolom"

    kolom_urut = ["Tahun", "Bulan", "Periode"] + kolom_target + ["Status Data"]
    return df_pred[kolom_urut], info_model

def warnai_baris_prediksi(row):
    if row["Status Data"] == "Prediksi":
        return ["background-color: #FEF9C3; color: #1E293B; font-weight: bold"] * len(row)
    return [""] * len(row)

def buat_file_excel(df: pd.DataFrame, nama_sheet: str, kolom_angka: list) -> bytes:
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

    return output.getvalue()

inisialisasi_data()

st.sidebar.title("📌 Menu Navigasi")
halaman = st.sidebar.radio(
    "Pilih Halaman:",
    ["1. PO Bulanan", "2. Peramalan Produk Utama"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Import Data dari Excel / CSV")
uploaded_file = st.sidebar.file_uploader(
    "Upload file template Excel (.xlsx) atau CSV",
    type=["xlsx", "csv"]
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_up = pd.read_csv(uploaded_file)
        else:
            df_up = pd.read_excel(uploaded_file)

        df_up.columns = [str(c).strip() for c in df_up.columns]
        if "Tahun" in df_up.columns:
            df_up["Tahun"] = df_up["Tahun"].ffill()

        if "Periode" in df_up.columns:
            df_up = df_up.dropna(subset=["Periode"]).copy()
            df_up["Periode"] = df_up["Periode"].astype(int)

        mode_import = st.sidebar.radio(
            "Pilih cara memproses file:",
            ["Tambahkan data baru ke dashboard", "Ganti seluruh data dengan file ini"]
        )

        if st.sidebar.button("🚀 Proses File Excel", use_container_width=True):
            if "Jumlah PO" in df_up.columns:
                df_b_baru = df_up.dropna(subset=["Jumlah PO"]).copy()
                df_b_baru["Tahun"] = df_b_baru["Tahun"].astype(int)
                df_b_baru["Jumlah PO"] = df_b_baru["Jumlah PO"].astype(int)
                df_b_baru["Status Data"] = "Aktual"
                df_b_baru = df_b_baru[["Tahun", "Bulan", "Periode", "Jumlah PO", "Status Data"]]

                if mode_import == "Tambahkan data baru ke dashboard":
                    gabung = pd.concat([st.session_state.df_bulanan, df_b_baru], ignore_index=True)
                    gabung = gabung.drop_duplicates(subset=["Periode"], keep="last").sort_values("Periode").reset_index(drop=True)
                    st.session_state.df_bulanan = gabung
                else:
                    st.session_state.df_bulanan = df_b_baru.sort_values("Periode").reset_index(drop=True)

            kolom_meta = {"Tahun", "Bulan", "Periode", "Jumlah PO", "Status Data"}
            kolom_prod_file = [c for c in df_up.columns if c not in kolom_meta]

            if kolom_prod_file:
                df_p_baru = df_up.dropna(subset=kolom_prod_file, how="all").copy()
                df_p_baru["Tahun"] = df_p_baru["Tahun"].astype(int)
                for c in kolom_prod_file:
                    df_p_baru[c] = df_p_baru[c].fillna(0).astype(int)
                df_p_baru["Status Data"] = "Aktual"
                df_p_baru = df_p_baru[["Tahun", "Bulan", "Periode"] + kolom_prod_file + ["Status Data"]]

                if mode_import == "Tambahkan data baru ke dashboard":
                    gabung_p = pd.concat([st.session_state.df_produk, df_p_baru], ignore_index=True)
                    gabung_p = gabung_p.drop_duplicates(subset=["Periode"], keep="last").sort_values("Periode").reset_index(drop=True)
                    for c in kolom_prod_file:
                        if c not in st.session_state.daftar_produk:
                            st.session_state.daftar_produk.append(c)
                        gabung_p[c] = gabung_p[c].fillna(0).astype(int)
                    st.session_state.df_produk = gabung_p[["Tahun", "Bulan", "Periode"] + st.session_state.daftar_produk + ["Status Data"]]
                else:
                    st.session_state.daftar_produk = kolom_prod_file
                    st.session_state.df_produk = df_p_baru.sort_values("Periode").reset_index(drop=True)

            st.sidebar.success("Data dari file berhasil diproses ke dashboard!")
            st.rerun()

    except Exception as e:
        st.sidebar.error(f"Gagal membaca file: {e}")

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Reset ke Data Awal (Periode 1–32)"):
    for key in ["df_bulanan", "df_produk", "daftar_produk", "mode_edit_bulanan", "mode_edit_produk"]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

# ==============================================================================
# HALAMAN 1: PO BULANAN
# ==============================================================================
if halaman == "1. PO Bulanan":
    st.title("📦 Rekap & Peramalan PO Bulanan")
    st.caption("Model Autoregressive (AR) memproyeksikan 3 bulan ke depan berdasarkan lag historis bulanan.")

    df_akt = st.session_state.df_bulanan.copy()
    df_pred, _ = buat_prediksi_3_bulan(df_akt, ["Jumlah PO"], metode="AR")
    df_gabungan = pd.concat([df_akt, df_pred], ignore_index=True)

    baris_pred_1 = df_pred.iloc[0]
    p_target = int(baris_pred_1["Periode"])
    bln_target = baris_pred_1["Bulan"]
    thn_target = int(baris_pred_1["Tahun"])
    val_pred_awal = int(baris_pred_1["Jumlah PO"])

    with st.container(border=True):
        st.subheader(f"✏️ Input Realisasi Aktual: {bln_target} {thn_target} (Periode {p_target})")
        st.write(
            f"Angka prediksi Autoregressive saat ini untuk **{bln_target} {thn_target}** adalah **{val_pred_awal:,} pcs**. "
            "Jika data realisasi bulan tersebut sudah keluar, masukkan angkanya di bawah ini:"
        )
        col_in, col_btn = st.columns([2, 1])
        with col_in:
            nilai_aktual_baru = st.number_input(
                f"Jumlah PO Aktual ({bln_target} {thn_target}):",
                min_value=0,
                value=val_pred_awal,
                step=10,
                key=f"input_bulanan_{p_target}"
            )
        with col_btn:
            st.write("")
            st.write("")
            if st.button("✅ Simpan & Ubah Jadi Aktual", type="primary", use_container_width=True):
                baris_baru = pd.DataFrame([{
                    "Tahun": thn_target,
                    "Bulan": bln_target,
                    "Periode": p_target,
                    "Jumlah PO": int(nilai_aktual_baru),
                    "Status Data": "Aktual"
                }])
                st.session_state.df_bulanan = pd.concat([df_akt, baris_baru], ignore_index=True)
                st.success(f"Data {bln_target} {thn_target} berhasil diubah menjadi Aktual! Prediksi baru otomatis ditambahkan.")
                st.rerun()

    col_hdr, col_edit_btn = st.columns([3, 1])
    with col_hdr:
        st.subheader("📋 Tabel Data Historis & Prediksi (Autoregressive)")
    with col_edit_btn:
        st.write("")
        if not st.session_state.mode_edit_bulanan:
            if st.button("✏️ Edit Data Historis", use_container_width=True):
                st.session_state.mode_edit_bulanan = True
                st.rerun()
        else:
            if st.button("❌ Batal Edit", use_container_width=True):
                st.session_state.mode_edit_bulanan = False
                st.rerun()

    if st.session_state.mode_edit_bulanan:
        st.info("Mode Edit Aktif: Silakan ubah angka pada kolom **Jumlah PO** di bawah ini jika ada kesalahan input, lalu klik **Simpan Perubahan**.")
        edited_bulanan = st.data_editor(
            df_akt,
            disabled=["Tahun", "Bulan", "Periode", "Status Data"],
            hide_index=True,
            use_container_width=True,
            height=450,
            key="editor_hist_bulanan"
        )
        if st.button("💾 Simpan Perubahan Data Historis", type="primary"):
            edited_bulanan["Jumlah PO"] = edited_bulanan["Jumlah PO"].fillna(0).astype(int)
            st.session_state.df_bulanan = edited_bulanan
            st.session_state.mode_edit_bulanan = False
            st.success("Perubahan data historis berhasil disimpan! Prediksi dan grafik telah diperbarui.")
            st.rerun()
    else:
        styled_df = df_gabungan.style.apply(warnai_baris_prediksi, axis=1).format({"Jumlah PO": "{:,}"})
        st.dataframe(styled_df, use_container_width=True, height=520, hide_index=True)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            csv_data = df_gabungan.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Tabel PO Bulanan (CSV)",
                data=csv_data,
                file_name="rekap_dan_prediksi_po_bulanan.csv",
                mime="text/csv",
                use_container_width=True
            )
        with col_dl2:
            excel_data = buat_file_excel(df_gabungan, "PO Bulanan", ["Jumlah PO"])
            st.download_button(
                "📊 Download Tabel PO Bulanan (Excel)",
                data=excel_data,
                file_name="rekap_dan_prediksi_po_bulanan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown("---")

    st.subheader("📈 Grafik Tren PO Bulanan (Aktual vs Prediksi AR)")
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df_akt["Periode"],
        y=df_akt["Jumlah PO"],
        mode="lines+markers",
        name="Aktual",
        line=dict(color="#1E3A5F", width=3),
        marker=dict(size=6),
        customdata=df_akt[["Bulan", "Tahun"]],
        hovertemplate="<b>%{customdata[0]} %{customdata[1]}</b> (Periode %{x})<br>Jumlah PO: %{y:,} pcs<extra></extra>"
    ))

    df_sambung_pred = pd.concat([df_akt.tail(1), df_pred], ignore_index=True)
    fig.add_trace(go.Scatter(
        x=df_sambung_pred["Periode"],
        y=df_sambung_pred["Jumlah PO"],
        mode="lines+markers",
        name="Prediksi AR (3 Bulan ke Depan)",
        line=dict(color="#D97706", width=3, dash="dash"),
        marker=dict(size=8, symbol="diamond"),
        customdata=df_sambung_pred[["Bulan", "Tahun"]],
        hovertemplate="<b>%{customdata[0]} %{customdata[1]}</b> (Periode %{x})<br>Proyeksi PO: %{y:,} pcs<extra></extra>"
    ))

    fig.update_layout(
        xaxis_title="Periode (Bulan ke-)",
        yaxis_title="Jumlah PO (pcs)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=460,
        margin=dict(l=20, r=20, t=40, b=20)
    )
    st.plotly_chart(fig, use_container_width=True)

# ==============================================================================
# HALAMAN 2: PERAMALAN PRODUK UTAMA (VAR)
# ==============================================================================
elif halaman == "2. Peramalan Produk Utama":
    st.title("🧥 Peramalan Produk Utama (VAR)")
    daftar_prod = st.session_state.daftar_produk
    df_akt_prod = st.session_state.df_produk.copy()
    df_pred_prod, info_var = buat_prediksi_3_bulan(df_akt_prod, daftar_prod, metode="VAR")
    df_gabungan_prod = pd.concat([df_akt_prod, df_pred_prod], ignore_index=True)

    st.caption(f"Proyeksi 3 bulan ke depan antarproduk secara simultan menggunakan Vector Autoregression (VAR). Status model aktif: **{info_var}**.")

    with st.expander("➕ Tambah Produk Baru untuk Diprediksi"):
        st.write("Masukkan nama produk baru dan isi data historis bulanannya (isi 0 jika belum ada pesanan di bulan tersebut):")
        nama_prod_baru = st.text_input("Nama Jenis PO Baru (contoh: HOODIE, T-SHIRT, PDH):").strip().upper()

        if nama_prod_baru:
            if nama_prod_baru in daftar_prod:
                st.warning(f"Produk **{nama_prod_baru}** sudah ada di dalam tabel.")
            else:
                df_input_hist = df_akt_prod[["Tahun", "Bulan", "Periode"]].copy()
                df_input_hist[nama_prod_baru] = 0
                edited_hist = st.data_editor(
                    df_input_hist,
                    disabled=["Tahun", "Bulan", "Periode"],
                    hide_index=True,
                    use_container_width=True,
                    height=260,
                    key="editor_produk_baru"
                )
                if st.button("💾 Tambahkan Produk ke Tabel & Grafik"):
                    st.session_state.df_produk[nama_prod_baru] = edited_hist[nama_prod_baru].fillna(0).astype(int)
                    st.session_state.daftar_produk.append(nama_prod_baru)
                    st.success(f"Produk **{nama_prod_baru}** berhasil ditambahkan!")
                    st.rerun()

    baris_pred_p1 = df_pred_prod.iloc[0]
    p_target = int(baris_pred_p1["Periode"])
    bln_target = baris_pred_p1["Bulan"]
    thn_target = int(baris_pred_p1["Tahun"])

    with st.container(border=True):
        st.subheader(f"✏️ Input Realisasi Aktual Produk: {bln_target} {thn_target} (Periode {p_target})")
        st.write(f"Ubah angka prediksi bulan **{bln_target} {thn_target}** menjadi data realisasi aktual:")

        cols_input = st.columns(len(daftar_prod))
        nilai_baru_dict = {}
        for idx, prod in enumerate(daftar_prod):
            with cols_input[idx]:
                nilai_baru_dict[prod] = st.number_input(
                    f"{prod} (pcs):",
                    min_value=0,
                    value=int(baris_pred_p1[prod]),
                    step=5,
                    key=f"in_{prod}_{p_target}"
                )

        if st.button("✅ Simpan Realisasi & Update Prediksi", type="primary"):
            new_row = {
                "Tahun": thn_target,
                "Bulan": bln_target,
                "Periode": p_target,
                "Status Data": "Aktual"
            }
            for prod in daftar_prod:
                new_row[prod] = int(nilai_baru_dict[prod])

            st.session_state.df_produk = pd.concat([df_akt_prod, pd.DataFrame([new_row])], ignore_index=True)
            st.success(f"Realisasi produk bulan {bln_target} {thn_target} tersimpan sebagai Aktual! Prediksi baru otomatis ditambahkan.")
            st.rerun()

    col_hdr_p, col_edit_btn_p = st.columns([3, 1])
    with col_hdr_p:
        st.subheader("📋 Tabel Historis & Proyeksi Produk Utama (VAR)")
    with col_edit_btn_p:
        st.write("")
        if not st.session_state.mode_edit_produk:
            if st.button("✏️ Edit Data Historis", key="btn_edit_prod", use_container_width=True):
                st.session_state.mode_edit_produk = True
                st.rerun()
        else:
            if st.button("❌ Batal Edit", key="btn_batal_prod", use_container_width=True):
                st.session_state.mode_edit_produk = False
                st.rerun()

    if st.session_state.mode_edit_produk:
        st.info("Mode Edit Aktif: Silakan koreksi angka historis produk di bawah ini jika ada kesalahan input, lalu klik **Simpan Perubahan**.")
        edited_prod = st.data_editor(
            df_akt_prod,
            disabled=["Tahun", "Bulan", "Periode", "Status Data"],
            hide_index=True,
            use_container_width=True,
            height=450,
            key="editor_hist_produk"
        )
        if st.button("💾 Simpan Perubahan Data Produk", type="primary"):
            for prod in daftar_prod:
                edited_prod[prod] = edited_prod[prod].fillna(0).astype(int)
            st.session_state.df_produk = edited_prod
            st.session_state.mode_edit_produk = False
            st.success("Perubahan data historis produk berhasil disimpan!")
            st.rerun()
    else:
        fmt_dict = {prod: "{:,}" for prod in daftar_prod}
        styled_prod = df_gabungan_prod.style.apply(warnai_baris_prediksi, axis=1).format(fmt_dict)
        st.dataframe(styled_prod, use_container_width=True, height=480, hide_index=True)

        col_dlp1, col_dlp2 = st.columns(2)
        with col_dlp1:
            csv_prod = df_gabungan_prod.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Prediksi Produk (CSV)",
                data=csv_prod,
                file_name="prediksi_produk_utama_var.csv",
                mime="text/csv",
                use_container_width=True
            )
        with col_dlp2:
            excel_prod = buat_file_excel(df_gabungan_prod, "Produk Utama", daftar_prod)
            st.download_button(
                "📊 Download Prediksi Produk (Excel)",
                data=excel_prod,
                file_name="prediksi_produk_utama_var.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown("---")

    st.subheader("📈 Grafik Tren Aktual & Prediksi per Produk (VAR)")
    palette = ["#1E3A5F", "#D97706", "#059669", "#7C3AED", "#DC2626", "#0284C7"]
    fig_prod = go.Figure()

    for i, prod in enumerate(daftar_prod):
        warna = palette[i % len(palette)]
        fig_prod.add_trace(go.Scatter(
            x=df_akt_prod["Periode"],
            y=df_akt_prod[prod],
            mode="lines+markers",
            name=f"{prod} (Aktual)",
            line=dict(color=warna, width=2.5),
            marker=dict(size=5)
        ))
        df_sambung = pd.concat([df_akt_prod.tail(1), df_pred_prod], ignore_index=True)
        fig_prod.add_trace(go.Scatter(
            x=df_sambung["Periode"],
            y=df_sambung[prod],
            mode="lines+markers",
            name=f"{prod} (Prediksi VAR)",
            line=dict(color=warna, width=2.5, dash="dash"),
            marker=dict(size=8, symbol="diamond")
        ))

    fig_prod.update_layout(
        xaxis_title="Periode (Bulan ke-)",
        yaxis_title="Jumlah PO (pcs)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=480,
        margin=dict(l=20, r=20, t=40, b=20)
    )
    st.plotly_chart(fig_prod, use_container_width=True)
