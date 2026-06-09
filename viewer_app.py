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
# FUNGSI LOAD DATA (Menggunakan st.secrets untuk keamanan)
# ═══════════════════════════════════════════════════════════
@st.cache_data(ttl=60)
def load_data_utama():
    credentials = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(credentials)
    
    sh = gc.open_by_key(ID_SPREADSHEET_UTAMA)
    worksheet = sh.get_worksheet(0)

    data = worksheet.get_all_values()
    df = pd.DataFrame(data)
    header_idx = df[df.eq('Nama Project').any(axis=1)].index[0]
    df.columns = df.iloc[header_idx]
    df = df.iloc[header_idx + 1:]
    
    # Menghapus kolom kosong dan duplikat
    df = df.loc[:, [bool(c) for c in df.columns]]
    df = df.loc[:, ~df.columns.duplicated()]
    
    df = df[df['Nama Project'] != '']
    df['Target']    = pd.to_numeric(df['Target'],    errors='coerce').fillna(0)
    df['Realisasi'] = pd.to_numeric(df['Realisasi'], errors='coerce').fillna(0)
    df = df.reset_index(drop=True)
    df.index = df.index + 1
    return df

@st.cache_data(ttl=60)
def load_tracker(nama_tracker: str) -> pd.DataFrame:
    credentials = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(credentials)
    
    sh = gc.open_by_key(ID_SPREADSHEET_TRACKER)
    ws = sh.worksheet(nama_tracker)
    raw = ws.get_all_values()

    if not raw: return pd.DataFrame()

    # Penambahan kata kunci deteksi untuk Tracker Nova, RTMD, & project telekomunikasi lainnya
    KATA_KUNCI_HEADER = (
        'no', 'nama', 'status', 'tanggal', 'kode', 'site', 'kategori', 'po', 'plan cost', 'actual cost',
        'progres', 'progress', 'target', 'realisasi', 'harga', 'cost', 'persentase',
        'revenue', 'margin', 'period', 'remark', 'deadline', 'prioritas',
        'customer', 'spk', 'pekerjaan', 'tagihan', 'id', 'pic'
    )
    best_idx, best_score = 0, -1
    for i, row in enumerate(raw[:15]):
        filled   = [str(c).strip() for c in row if str(c).strip()]
        kw_bonus = sum(1 for c in filled if any(k in str(c).lower() for k in KATA_KUNCI_HEADER))
        score    = len(filled) + kw_bonus * 2
        if score > best_score: best_score, best_idx = score, i

    headers = [str(h).strip() for h in raw[best_idx]]
    df = pd.DataFrame(raw[best_idx + 1:], columns=headers)
    
    # ── FIX BUG UTAMA TRACKER ──
    # 1. Buang kolom tanpa nama
    df = df.loc[:, [bool(c) for c in df.columns]]
    # 2. Buang duplikasi kolom (mencegah error value_counts)
    df = df.loc[:, ~df.columns.duplicated()]
    # 3. Buang baris yang isinya kosong semua
    df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
    
    return df.reset_index(drop=True)

# ═══════════════════════════════════════════════════════════
# HELPER VISUALISASI TRACKER KHUSUS
# ═══════════════════════════════════════════════════════════
def parse_rupiah(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
              .str.replace(r'[Rp\s\.]', '', regex=True)
              .str.replace(',', '', regex=False),
        errors='coerce'
    ).fillna(0)

def _cari_kolom(cols: list, *keywords: str):
    for kw in keywords:
        for c in cols:
            if kw.lower() in c.lower():
                return c
    return None

def render_tracker_ringkasan(df: pd.DataFrame, tracker_name: str):
    cols = df.columns.tolist()

    if any(k in tracker_name for k in ('RTMD', 'Kemenkes')):
        plan_site   = _cari_kolom(cols, 'Plan Site')
        actual_site = _cari_kolom(cols, 'Actual Site')
        tag_ent     = _cari_kolom(cols, 'Tag Entities', 'Entitas')
        site_name   = _cari_kolom(cols, 'Site Name')
        progres_col = _cari_kolom(cols, 'Progres', 'Progress')
        stat_doc    = _cari_kolom(cols, 'Stat Doc', 'Status Doc')

        po_col      = next((c for c in cols if c.strip().lower() == 'po'), None)
        ac_col      = next((c for c in cols if c.strip().lower() == 'actual cost'), None)
        margin_col  = next((c for c in cols if c.strip().lower() == 'margin'), None)

        if plan_site: df[plan_site]   = pd.to_numeric(df[plan_site], errors='coerce').fillna(0)
        if actual_site: df[actual_site] = pd.to_numeric(df[actual_site], errors='coerce').fillna(0)

        tot_plan   = int(df[plan_site].sum()) if plan_site else 0
        tot_actual = int(df[actual_site].sum()) if actual_site else 0
        pct_txt    = f"{tot_actual / tot_plan * 100:.1f}%" if tot_plan else "N/A"

        tot_po     = parse_rupiah(df[po_col]).sum() if po_col else 0
        tot_ac     = parse_rupiah(df[ac_col]).sum() if ac_col else 0
        tot_margin = parse_rupiah(df[margin_col]).sum() if margin_col else 0

        st.markdown("**Ringkasan Finansial**")
        cf1, cf2, cf3, cf4 = st.columns(4)
        cf1.metric("🏢 Total Site / Pekerjaan", len(df))
        cf2.metric("💰 Total PO", f"Rp {tot_po:,.0f}" if tot_po else "—")
        cf3.metric("📉 Total Actual Cost", f"Rp {tot_ac:,.0f}" if tot_ac else "—")
        cf4.metric("📈 Total Margin", f"Rp {tot_margin:,.0f}" if tot_margin else "—")
        
        # Grafik Finansial Total Saja (Bukan per Site)
        if po_col and ac_col and margin_col:
            st.markdown("<br>", unsafe_allow_html=True)
            fin_data = pd.DataFrame({
                'Kategori': ['Total PO', 'Total Actual Cost', 'Total Margin'],
                'Nilai': [tot_po, tot_ac, tot_margin]
            })
            
            fig_fin = px.bar(fin_data, x='Kategori', y='Nilai', text_auto='.2s', color='Kategori',
                             color_discrete_map={'Total PO': '#3498db', 'Total Actual Cost': '#e74c3c', 'Total Margin': '#2ecc71'})
            fig_fin.update_traces(textfont_size=12, textangle=0, textposition="outside", cliponaxis=False)
            fig_fin.update_layout(yaxis_title="Nilai (Rp)", showlegend=False, margin=dict(t=20, b=0), height=350)
            st.plotly_chart(fig_fin, use_container_width=True)

        st.markdown("---")
        st.markdown("**Ringkasan Kunjungan**")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 Plan Kunjungan", tot_plan if tot_plan else "—")
        c2.metric("✅ Actual Kunjungan", tot_actual if tot_actual else "—")
        c3.metric("📈 Progress Kunjungan", pct_txt)
        
        st.markdown("<br>", unsafe_allow_html=True)

        left, right = st.columns(2)
        if plan_site and actual_site and site_name:
            with left:
                st.markdown("**Plan vs Actual Kunjungan per Site**")
                viz = df[[site_name, plan_site, actual_site]].copy()
                viz = viz[(viz[plan_site] > 0) | (viz[actual_site] > 0)]
                if not viz.empty:
                    melted = viz.melt(id_vars=site_name, value_vars=[plan_site, actual_site], var_name='Tipe', value_name='Kunjungan')
                    fig = px.bar(melted, y=site_name, x='Kunjungan', color='Tipe', orientation='h', barmode='group',
                                 color_discrete_map={plan_site: '#bdc3c7', actual_site: '#2ecc71'}, height=max(350, len(viz) * 35))
                    fig.update_layout(yaxis_title="", legend_title="", margin=dict(t=20, b=0))
                    st.plotly_chart(fig, use_container_width=True)

        if tag_ent:
            with right:
                st.markdown("**Distribusi Entitas**")
                tc = (df[tag_ent].replace('', 'Lainnya').fillna('Lainnya').value_counts().reset_index())
                tc.columns = ['Entitas', 'Jumlah']
                fig = px.pie(tc, values='Jumlah', names='Entitas', hole=0.4)
                fig.update_layout(margin=dict(t=20, b=0))
                st.plotly_chart(fig, use_container_width=True)

        if progres_col or stat_doc:
            l2, r2 = st.columns(2)
            if progres_col:
                with l2:
                    st.markdown("**Status Progress**")
                    pc = (df[progres_col].replace('', 'Belum').fillna('Belum').value_counts().reset_index())
                    pc.columns = ['Progress', 'Jumlah']
                    fig = px.bar(pc, x='Progress', y='Jumlah', color='Progress')
                    fig.update_layout(showlegend=False, margin=dict(t=20, b=0))
                    st.plotly_chart(fig, use_container_width=True)
            if stat_doc:
                with r2:
                    st.markdown("**Status Dokumen**")
                    dc = (df[stat_doc].replace('', 'Belum Ada').fillna('Belum Ada').value_counts().reset_index())
                    dc.columns = ['Status', 'Jumlah']
                    fig = px.pie(dc, values='Jumlah', names='Status', hole=0.4)
                    fig.update_layout(margin=dict(t=20, b=0))
                    st.plotly_chart(fig, use_container_width=True)

    elif 'KDS' in tracker_name:
        rev_col      = _cari_kolom(cols, 'Revenue')
        cost_col     = _cari_kolom(cols, 'Cost')
        margin_col   = next((c for c in cols if c.strip() == 'Margin'), None) or _cari_kolom(cols, 'Margin')
        st_tagihan   = _cari_kolom(cols, 'Status Tagihan')
        st_pekerjaan = _cari_kolom(cols, 'Status Pekerjaan')
        period_col   = _cari_kolom(cols, 'Period')

        tot_rev    = parse_rupiah(df[rev_col]).sum() if rev_col else 0
        tot_cost   = parse_rupiah(df[cost_col]).sum() if cost_col else 0
        tot_margin = parse_rupiah(df[margin_col]).sum() if margin_col and 'Persentase' not in margin_col else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 Total Pekerjaan", len(df))
        c2.metric("💰 Total Revenue", f"Rp {tot_rev:,.0f}" if tot_rev else "—")
        c3.metric("📉 Total Cost", f"Rp {tot_cost:,.0f}" if tot_cost else "—")
        c4.metric("📈 Total Margin", f"Rp {tot_margin:,.0f}" if tot_margin else "—")

        left, right = st.columns(2)
        if st_tagihan:
            with left:
                st.markdown("**Status Tagihan**")
                tc = (df[st_tagihan].replace('', 'N/A').fillna('N/A').value_counts().reset_index())
                tc.columns = ['Status', 'Jumlah']
                fig = px.pie(tc, values='Jumlah', names='Status', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(margin=dict(t=20, b=0))
                st.plotly_chart(fig, use_container_width=True)
        if st_pekerjaan:
            with right:
                st.markdown("**Status Pekerjaan**")
                sc = (df[st_pekerjaan].replace('', 'N/A').fillna('N/A').value_counts().reset_index())
                sc.columns = ['Status', 'Jumlah']
                fig = px.bar(sc, x='Status', y='Jumlah', color='Status', color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(showlegend=False, margin=dict(t=20, b=0))
                st.plotly_chart(fig, use_container_width=True)

        if rev_col and period_col:
            st.markdown("**Revenue per Period**")
            dp = df[[period_col, rev_col]].copy()
            dp[rev_col] = parse_rupiah(dp[rev_col])
            dp = dp.groupby(period_col)[rev_col].sum().reset_index()
            dp.columns = ['Period', 'Revenue']
            fig = px.bar(dp, x='Period', y='Revenue', color_discrete_sequence=['#3498db'])
            fig.update_layout(xaxis_title="Period", yaxis_title="Revenue (Rp)", margin=dict(t=20, b=0))
            st.plotly_chart(fig, use_container_width=True)

    else:
        # ── GENERIC DASHBOARD UTK NOVA, DLL ──
        c1, c2 = st.columns(2)
        c1.metric("📋 Total Baris", len(df))
        c2.metric("🔢 Total Kolom", len(cols))
        
        # Fix Bug: Mengabaikan kolom "Tanggal" atau "Date" agar tidak dipaksa jadi pie chart
        status_cols = [c for c in cols if any(
            k in c.lower() for k in ('status', 'progres', 'progress', 'kondisi', 'state')
        ) and not any(
            x in c.lower() for x in ('date', 'tanggal', 'waktu', 'time', 'periode', 'period')
        )]
        
        if status_cols:
            n = min(len(status_cols), 2)
            chart_cols = st.columns(n)
            for i, sc in enumerate(status_cols[:n]):
                with chart_cols[i]:
                    st.markdown(f"**{sc}**")
                    try:
                        # Standardisasi ke string lalu ganti blank text
                        vc = df[sc].astype(str).replace(r'^\s*$', 'Kosong', regex=True).replace('nan', 'Kosong').value_counts().reset_index()
                        vc.columns = ['Label', 'Jumlah']
                        
                        fig = px.pie(vc, values='Jumlah', names='Label', hole=0.4)
                        fig.update_layout(margin=dict(t=20, b=0))
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception:
                        st.info(f"Visualisasi untuk kolom {sc} tidak dapat dimuat.")


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
        # Menjalankan fungsi visualisasi ringkasan
        render_tracker_ringkasan(df_tracker, selected_tracker)
        st.markdown("---")
        st.markdown("**Tabel Detail Tracker**")
        st.dataframe(df_tracker, use_container_width=True, height=420)
    else:
        st.warning("⚠️ Data tracker kosong.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")
