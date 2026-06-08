import streamlit as st
import gspread
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard Project Viewer", layout="wide")
st.title("📊 Dashboard Manajemen Project (View Only)")
st.markdown("---")

# ═══════════════════════════════════════════════════════════
# KONFIGURASI
# ═══════════════════════════════════════════════════════════
ID_SPREADSHEET_UTAMA   = '1Jbe3pyoiLj_C3p2BwwyruQn1AEaWGCdz7aDaw9T55e0'
ID_SPREADSHEET_TRACKER = '198lmRiSC1VjZeEL9CK11miLLQlBMQBjXKubEJMX0MO0'

DAFTAR_TRACKER = [
    "Tracker KDS", "Tracker BGI", "Tracker Material",
    "Tracker Nova", "Tracker RTMD Kemenkes",
]

# ═══════════════════════════════════════════════════════════
# FUNGSI LOAD DATA (Hanya membaca, tanpa akses edit)
# ═══════════════════════════════════════════════════════════
@st.cache_data(ttl=60) # Cache data selama 1 menit agar tidak berat
def load_data_utama():
    # CATATAN: Nanti saat di-online-kan, bagian ini perlu disesuaikan menggunakan st.secrets
    gc = gspread.service_account(filename='credentials.json')
    sh = gc.open_by_key(ID_SPREADSHEET_UTAMA)
    worksheet = sh.get_worksheet(0)

    data = worksheet.get_all_values()
    df = pd.DataFrame(data)
    header_idx = df[df.eq('Nama Project').any(axis=1)].index[0]
    df.columns = df.iloc[header_idx]
    df = df.iloc[header_idx + 1:]
    df = df.loc[:, df.columns != '']
    df = df[df['Nama Project'] != '']
    df['Target']    = pd.to_numeric(df['Target'],    errors='coerce').fillna(0)
    df['Realisasi'] = pd.to_numeric(df['Realisasi'], errors='coerce').fillna(0)
    df = df.reset_index(drop=True)
    df.index = df.index + 1
    return df

@st.cache_data(ttl=60)
def load_tracker(nama_tracker: str) -> pd.DataFrame:
    gc = gspread.service_account(filename='credentials.json')
    sh = gc.open_by_key(ID_SPREADSHEET_TRACKER)
    ws = sh.worksheet(nama_tracker)
    raw = ws.get_all_values()

    if not raw: return pd.DataFrame()

    KATA_KUNCI_HEADER = ('no', 'nama', 'status', 'tanggal', 'kode', 'site', 'kategori', 'progres', 'target', 'realisasi', 'harga', 'cost', 'revenue', 'margin', 'period', 'remark', 'deadline', 'prioritas')
    best_idx, best_score = 0, -1
    for i, row in enumerate(raw[:12]):
        filled   = [str(c).strip() for c in row if str(c).strip()]
        kw_bonus = sum(1 for c in filled if any(k in str(c).lower() for k in KATA_KUNCI_HEADER))
        score    = len(filled) + kw_bonus * 2
        if score > best_score: best_score, best_idx = score, i

    headers = [str(h).strip() for h in raw[best_idx]]
    df = pd.DataFrame(raw[best_idx + 1:], columns=headers)
    df = df.loc[:, [c for c in df.columns if c]]
    df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
    return df.reset_index(drop=True)

# Helper untuk memformat angka (disembunyikan untuk menyingkat kode, gunakan fungsi render_tracker_ringkasan yang sama persis seperti kode sebelumnya)
# -- Masukkan fungsi parse_rupiah, _cari_kolom, dan render_tracker_ringkasan dari file sebelumnya di sini --

# (Untuk contoh ini, saya melompati fungsi render_tracker agar kode tidak kepanjangan, 
# Anda tinggal copy-paste fungsi tersebut dari app.py ke sini)

# ═══════════════════════════════════════════════════════════
# MAIN APP - TAMPILAN DASHBOARD SAJA
# ═══════════════════════════════════════════════════════════
try:
    df = load_data_utama()

    # ── Ringkasan Eksekutif ─────────────────────────────────────────
    st.subheader("Ringkasan Eksekutif")
    c1, c2, c3 = st.columns(3)
    total_target       = df['Target'].sum()
    total_realisasi    = df['Realisasi'].sum()
    persen_keseluruhan = (total_realisasi / total_target * 100) if total_target > 0 else 0
    c1.metric("Total Project",          len(df))
    c2.metric("Target Rata-rata",       "100%")
    c3.metric("Pencapaian Keseluruhan", f"{persen_keseluruhan:.1f}%")
    st.markdown("---")

    # ── Visualisasi Analytics ───────────────────────────────────────
    st.subheader("Visualisasi Analytics")
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.markdown("**Status Keseluruhan (Total)**")
        data_pie = pd.DataFrame({
            'Status': ['Sudah Terealisasi', 'Sisa Target Belum Tercapai'],
            'Nilai':  [total_realisasi, total_target - total_realisasi],
        })
        fig_pie = px.pie(data_pie, values='Nilai', names='Status', hole=0.4,
                         color='Status', color_discrete_map={'Sudah Terealisasi': '#2ecc71', 'Sisa Target Belum Tercapai': '#e74c3c'})
        fig_pie.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5), margin=dict(t=50, b=0, l=0, r=0))
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_chart2:
        st.markdown("**Progress Spesifik per Project (On Progress & Pending)**")
        df_filtered = df[df['Status'].isin(['On Progress', 'Pending'])].copy()
        if not df_filtered.empty:
            priority_map = {'High': 1, 'Medium': 2, 'Low': 3}
            df_filtered['Priority_Rank'] = df_filtered['Prioritas'].map(priority_map).fillna(4)
            df_filtered = df_filtered.sort_values(by=['Priority_Rank', 'Nama Project'], ascending=[False, False])
            
            df_melt = df_filtered.melt(id_vars='Nama Project', value_vars=['Target', 'Realisasi'], var_name='Kategori_Stat', value_name='Nilai (%)')
            fig_bar = px.bar(df_melt, y='Nama Project', x='Nilai (%)', color='Kategori_Stat', barmode='group', orientation='h',
                             color_discrete_map={'Target': '#bdc3c7', 'Realisasi': '#3498db'})
            fig_bar.update_layout(yaxis_title="", xaxis_title="Persentase (%)", legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5), margin=dict(t=50, b=0, l=0, r=0))
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("👍 Tidak ada project yang On Progress atau Pending.")

    # ── Detail Status Project ───────────────────────────────────────
    st.markdown("---")
    st.subheader("Detail Status Project")
    priority_map = {'High': 1, 'Medium': 2, 'Low': 3}
    df_display = df.copy()
    df_display['Priority_Rank'] = df_display['Prioritas'].map(priority_map).fillna(4)
    df_display = df_display.sort_values(by=['Priority_Rank', 'Nama Project']).drop(columns=['Priority_Rank'])
    st.dataframe(df_display, use_container_width=True)

    # ── Tracker Khusus ──────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🗂️ Lihat Data Tracker Lainnya")
    selected_tracker = st.selectbox("Pilih Tracker yang ingin dilihat:", DAFTAR_TRACKER)
    df_tracker = load_tracker(selected_tracker)
    if not df_tracker.empty:
        st.dataframe(df_tracker, use_container_width=True, height=420)
    else:
        st.warning("⚠️ Data tracker kosong.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")