import streamlit as st
import gspread
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Dashboard Project Viewer", layout="wide")

# ═══════════════════════════════════════════════════════════
# KONFIGURASI
# ═══════════════════════════════════════════════════════════
ID_SPREADSHEET_UTAMA   = '1Jbe3pyoiLj_C3p2BwwyruQn1AEaWGCdz7aDaw9T55e0'
ID_SPREADSHEET_TRACKER = '198lmRiSC1VjZeEL9CK11miLLQlBMQBjXKubEJMX0MO0'

DAFTAR_TRACKER = [
    "Tracker KDS", "Tracker BGI", "Tracker Material",
    "Tracker Nova", "Tracker RTMD Kemenkes", "Tracker KCS Batam"
]

# ═══════════════════════════════════════════════════════════
# FUNGSI LOAD DATA (Menggunakan st.secrets & Caching)
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

    KATA_KUNCI_HEADER = (
        'no', 'nama', 'status', 'tanggal', 'kode', 'site', 'kategori', 'po', 'plan cost', 'actual cost',
        'progres', 'progress', 'target', 'realisasi', 'harga', 'cost', 'persentase',
        'revenue', 'margin', 'period', 'remark', 'deadline', 'prioritas',
        'customer', 'spk', 'pekerjaan', 'tagihan', 'id', 'pic', 'assign',
        'fase proyek', 'deskripsi', 'sistem', 'catatan', 'checklist'
    )
    best_idx, best_score = 0, -1
    for i, row in enumerate(raw[:15]):
        filled   = [str(c).strip() for c in row if str(c).strip()]
        kw_bonus = sum(1 for c in filled if any(k in str(c).lower() for k in KATA_KUNCI_HEADER))
        score    = len(filled) + kw_bonus * 2
        if score > best_score: best_score, best_idx = score, i

    headers = [str(h).strip() for h in raw[best_idx]]
    df = pd.DataFrame(raw[best_idx + 1:], columns=headers)
    
    df = df.loc[:, [bool(c) for c in df.columns]]
    df = df.loc[:, ~df.columns.duplicated()]
    df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
    
    return df.reset_index(drop=True)

@st.cache_data(ttl=60)
def load_timeline_dssp():
    credentials = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(credentials)
    sh = gc.open_by_key(ID_SPREADSHEET_TRACKER)
    try:
        ws = sh.worksheet("Timeline")
    except gspread.exceptions.WorksheetNotFound:
        return pd.DataFrame()

    raw = ws.get_all_values()
    if not raw:
        return pd.DataFrame()

    KATA_KUNCI = ('sub-task', 'sub task', 'start')
    header_idx = -1
    for i, row in enumerate(raw[:15]): 
        if any(k in str(c).lower() for c in row for k in KATA_KUNCI):
            header_idx = i
            break

    if header_idx == -1:
        return pd.DataFrame()

    headers = [str(h).strip() for h in raw[header_idx]]
    df = pd.DataFrame(raw[header_idx + 1:], columns=headers)

    core_cols = ['Task', 'Sub-Task', 'Start', 'Deadline Finish', 'Actual Finish Date', 'Remark Status', 'PIC']
    available_cols = [c for c in core_cols if c in df.columns]
    df = df[available_cols]

    if 'Sub-Task' in df.columns:
        df = df[df['Sub-Task'].astype(str).str.strip() != '']
    if 'Task' in df.columns:
        df['Task'] = df['Task'].replace(r'^\s*$', pd.NA, regex=True).ffill()

    return df.reset_index(drop=True)

# ═══════════════════════════════════════════════════════════
# HELPER VISUALISASI
# ═══════════════════════════════════════════════════════════
def parse_rupiah(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
              .str.replace(r'[Rp\s\.]', '', regex=True)
              .str.replace(',', '', regex=False),
        errors='coerce'
    ).fillna(0)

def _cari_kolom(cols: list, *keywords: str):
    """Mencari nama kolom aktual di dataframe berdasarkan kata kunci"""
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

    # ====================================================================================
    # UPDATE BARU: Visualisasi Khusus Tracker KCS Batam (Pipeline / Task Management)
    # ====================================================================================
    elif 'KCS Batam' in tracker_name:
        col_fase   = _cari_kolom(cols, 'fase', 'phase')
        col_status = _cari_kolom(cols, 'status')
        col_prog   = _cari_kolom(cols, 'progres', 'progress')
        col_pic    = _cari_kolom(cols, 'pic', 'penanggung jawab')
        col_sistem = _cari_kolom(cols, 'sistem', 'system')

        # 1. Row Metrik (Highlight Angka Penting)
        c1, c2, c3 = st.columns(3)
        c1.metric("📋 Total Task Tersedia", len(df))
        
        if col_status:
            done_count = len(df[df[col_status].astype(str).str.lower().isin(['done', 'selesai'])])
            c2.metric("✅ Task Selesai", done_count)
        else:
            c2.metric("✅ Task Selesai", "N/A")

        if col_prog:
            # Hilangkan simbol '%' dan konversi ke numeric untuk mendapatkan rata-rata
            prog_series = pd.to_numeric(df[col_prog].astype(str).str.replace('%', '', regex=False), errors='coerce').fillna(0)
            avg_prog = prog_series.mean()
            c3.metric("📈 Rata-rata Progress", f"{avg_prog:.1f}%")
        else:
            c3.metric("📈 Rata-rata Progress", "N/A")

        st.markdown("---")
        
        # 2. Visualisasi Grafik Status
        left, right = st.columns(2)
        color_map_status = {
            'Done': '#2ecc71', 'Selesai': '#2ecc71',
            'On Progress': '#3498db', 'In Progress': '#3498db', 'Progress': '#3498db',
            'Not Started': '#e74c3c', 'Belum Mulai': '#e74c3c'
        }

        if col_status:
            with left:
                st.markdown("**Status Keseluruhan Task**")
                sc = df[col_status].replace(r'^\s*$', 'Unknown', regex=True).fillna('Unknown').value_counts().reset_index()
                sc.columns = ['Status', 'Jumlah']
                
                fig_stat = px.pie(sc, values='Jumlah', names='Status', hole=0.4, color='Status', color_discrete_map=color_map_status)
                fig_stat.update_layout(margin=dict(t=20, b=0))
                st.plotly_chart(fig_stat, use_container_width=True)

        if col_fase and col_status:
            with right:
                st.markdown("**Status Task per Fase Proyek**")
                df_fase = df.groupby([col_fase, col_status]).size().reset_index(name='Jumlah')
                fig_fase = px.bar(df_fase, x=col_fase, y='Jumlah', color=col_status, barmode='stack', color_discrete_map=color_map_status)
                fig_fase.update_layout(xaxis_title="", yaxis_title="Jumlah Task", legend_title="", margin=dict(t=20, b=0))
                st.plotly_chart(fig_fase, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # 3. Visualisasi Pendukung: PIC & Sistem
        l2, r2 = st.columns(2)
        if col_pic:
            with l2:
                st.markdown("**Distribusi Beban Kerja PIC**")
                pc = df[col_pic].replace(r'^\s*$', 'Unassigned', regex=True).fillna('Unassigned').value_counts().reset_index()
                pc.columns = ['PIC', 'Jumlah Task']
                fig_pic = px.bar(pc, x='PIC', y='Jumlah Task', color='PIC')
                fig_pic.update_layout(showlegend=False, margin=dict(t=20, b=0), xaxis_title="")
                st.plotly_chart(fig_pic, use_container_width=True)
                
        if col_sistem:
            with r2:
                st.markdown("**Distribusi Kategori Sistem**")
                sys_c = df[col_sistem].replace(r'^\s*$', 'Unknown', regex=True).fillna('Unknown').value_counts().reset_index()
                sys_c.columns = ['Sistem', 'Jumlah Task']
                fig_sys = px.pie(sys_c, values='Jumlah Task', names='Sistem', hole=0.4)
                fig_sys.update_layout(margin=dict(t=20, b=0))
                st.plotly_chart(fig_sys, use_container_width=True)


    else:
        c1, c2 = st.columns(2)
        c1.metric("📋 Total Baris", len(df))
        c2.metric("🔢 Total Kolom", len(cols))
        
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
                        vc = df[sc].astype(str).replace(r'^\s*$', 'Kosong', regex=True).replace('nan', 'Kosong').value_counts().reset_index()
                        vc.columns = ['Label', 'Jumlah']
                        
                        fig = px.pie(vc, values='Jumlah', names='Label', hole=0.4)
                        fig.update_layout(margin=dict(t=20, b=0))
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception:
                        st.info(f"Visualisasi untuk kolom {sc} tidak dapat dimuat.")


# ═══════════════════════════════════════════════════════════
# MAIN APP - TAMPILAN DASHBOARD SAJA (VIEWER)
# ═══════════════════════════════════════════════════════════
try:
    # ── HEADER & TOMBOL REFRESH GLOBAL ──
    col_title, col_btn = st.columns([5, 1])
    with col_title:
        st.title("📊 Dashboard Manajemen Project (View Only)")
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Muat Ulang Data", use_container_width=True):
            load_data_utama.clear()
            load_tracker.clear()
            load_timeline_dssp.clear()
            st.rerun()
            
    st.markdown("---")

    df = load_data_utama()

    # ── TABS UNTUK VIEWER ──
    tab1, tab2, tab3 = st.tabs([
        "📊 Main Dashboard", 
        "🗂️ Lihat Tracker", 
        "📅 Timeline DSSP"
    ])

    # ── TAB 1 : MAIN DASHBOARD ──
    with tab1:
        st.subheader("Ringkasan Eksekutif")
        c1, c2, c3 = st.columns(3)
        total_target       = df['Target'].sum()
        total_realisasi    = df['Realisasi'].sum()
        persen_keseluruhan = (total_realisasi / total_target * 100) if total_target > 0 else 0
        c1.metric("Total Project",          len(df))
        c2.metric("Target Rata-rata",       "100%")
        c3.metric("Pencapaian Keseluruhan", f"{persen_keseluruhan:.1f}%")
        st.markdown("---")

        st.subheader("Visualisasi Analytics")
        col_chart1, col_chart2 = st.columns(2)

        with col_chart1:
            st.markdown("**Status Keseluruhan (Total Target vs Realisasi)**")
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
                if 'Prioritas' in df_filtered.columns:
                    df_filtered['Priority_Rank'] = df_filtered['Prioritas'].map(priority_map).fillna(4)
                    df_filtered = df_filtered.sort_values(by=['Priority_Rank', 'Nama Project'], ascending=[False, False])
                else:
                    df_filtered = df_filtered.sort_values(by=['Nama Project'], ascending=[False])
                
                df_melt = df_filtered.melt(id_vars='Nama Project', value_vars=['Target', 'Realisasi'], var_name='Kategori_Stat', value_name='Nilai (%)')
                fig_bar = px.bar(df_melt, y='Nama Project', x='Nilai (%)', color='Kategori_Stat', barmode='group', orientation='h',
                                 color_discrete_map={'Target': '#bdc3c7', 'Realisasi': '#3498db'})
                fig_bar.update_layout(yaxis_title="", xaxis_title="Persentase (%)", legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5), margin=dict(t=50, b=0, l=0, r=0))
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("👍 Tidak ada project yang On Progress atau Pending.")

        st.markdown("---")
        
        cols_main = df.columns.tolist()
        col_company = _cari_kolom(cols_main, 'company', 'perusahaan', 'klien', 'client')
        col_status  = _cari_kolom(cols_main, 'status') or 'Status'
        col_remark  = _cari_kolom(cols_main, 'remark', 'keterangan', 'catatan')
        col_pic     = _cari_kolom(cols_main, 'pic', 'penanggung jawab', 'assign', 'in charge')
        
        st.subheader("🏢 Distribusi Status Project per Company")
        if col_company and col_status:
            df_comp = df.copy()
            df_comp[col_company] = df_comp[col_company].replace(r'^\s*$', 'Lainnya/Kosong', regex=True).fillna('Lainnya/Kosong')
            df_comp[col_status]  = df_comp[col_status].replace(r'^\s*$', 'Unknown', regex=True).fillna('Unknown')
            
            df_dist = df_comp.groupby([col_company, col_status]).size().reset_index(name='Jumlah Project')
            
            color_map_status = {
                'Done': '#2ecc71', 'Selesai': '#2ecc71', 'Closed': '#2ecc71',
                'On Progress': '#3498db', 'In Progress': '#3498db', 'Progress': '#3498db',
                'Pending': '#e74c3c', 'Hold': '#e74c3c', 'Belum Mulai': '#f1c40f'
            }
            
            fig_comp = px.bar(df_dist, x=col_company, y='Jumlah Project', color=col_status, 
                              barmode='group', text='Jumlah Project', color_discrete_map=color_map_status)
            fig_comp.update_traces(textposition='outside')
            fig_comp.update_layout(xaxis_title="Company", yaxis_title="Jumlah Project", legend_title="Status", margin=dict(t=20, b=0))
            st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.info("⚠️ Kolom 'Company' atau sejenisnya tidak ditemukan di sheet utama untuk membuat grafik distribusi.")

        st.markdown("---")
        
        st.subheader("📋 Pantauan Khusus: Project On Progress & Pending")
        st.write("Daftar project aktif beserta *Remark* dan *PIC* untuk pemantauan.")
        
        status_aktif = [s for s in df[col_status].dropna().unique() if any(k in str(s).lower() for k in ('progress', 'pending', 'hold', 'jalan'))]
        if not status_aktif:
            status_aktif = ['On Progress', 'Pending'] 
            
        df_active = df[df[col_status].isin(status_aktif)].copy()
        
        if not df_active.empty:
            show_cols = ['Nama Project']
            if col_company: show_cols.append(col_company)
            show_cols.append(col_status)
            show_cols.append('Realisasi')
            if col_pic: show_cols.append(col_pic)
            if col_remark: show_cols.append(col_remark)
            
            col_prio = _cari_kolom(cols_main, 'prioritas', 'priority')
            if col_prio:
                show_cols.insert(1, col_prio) 
                priority_map = {'High': 1, 'Medium': 2, 'Low': 3}
                df_active['Priority_Rank'] = df_active[col_prio].map(priority_map).fillna(4)
                df_active = df_active.sort_values(by=['Priority_Rank', 'Nama Project']).drop(columns=['Priority_Rank'])
            
            st.dataframe(df_active[show_cols], use_container_width=True, hide_index=True)
        else:
            st.success("✅ Hebat! Tidak ada project yang *On Progress* atau *Pending* saat ini (Semua sudah selesai).")

        with st.expander("Lihat Seluruh Tabel Project Lengkap"):
            st.dataframe(df, use_container_width=True)

    # ── TAB 2 : LIHAT TRACKER LAINNYA ──
    with tab2:
        st.subheader("🗂️ Detail Tracker Terpilih")
        selected_tracker = st.selectbox("Pilih Tracker yang ingin dilihat:", DAFTAR_TRACKER)
        
        df_tracker = load_tracker(selected_tracker)
        
        if not df_tracker.empty:
            render_tracker_ringkasan(df_tracker, selected_tracker)
            st.markdown("---")
            st.markdown("**Tabel Detail Data**")
            st.dataframe(df_tracker, use_container_width=True, height=420)
        else:
            st.warning("⚠️ Data tracker kosong.")

    # ── TAB 3 : TIMELINE DSSP ──
    with tab3:
        st.subheader("📅 Timeline Interaktif DSSP")
        st.write("Visualisasi Gantt Chart interaktif berdasarkan Sheet `Timeline`.")
        
        with st.spinner("Memuat Timeline..."):
            df_timeline = load_timeline_dssp()
        
        if not df_timeline.empty:
            end_col = 'Deadline Finish' if 'Deadline Finish' in df_timeline.columns else 'Deadline'
            
            if 'Start' in df_timeline.columns and end_col in df_timeline.columns:
                
                def parse_timeline_date(d):
                    try:
                        s = str(d).strip()
                        if not s or s.lower() in ['nan', 'none', '-']: return pd.NaT
                        if len(s) <= 6 and '-' in s: s += f"-{datetime.now().year}"
                        return pd.to_datetime(s)
                    except:
                        return pd.NaT

                df_timeline['Start_Parsed'] = df_timeline['Start'].apply(parse_timeline_date)
                df_timeline['End_Parsed']   = df_timeline[end_col].apply(parse_timeline_date)
                
                df_plot = df_timeline.dropna(subset=['Start_Parsed', 'End_Parsed']).copy()
                
                if not df_plot.empty:
                    def format_ylabel(row):
                        task = str(row['Task']).strip()
                        sub = str(row['Sub-Task']).strip()
                        teks = f"{task} - {sub}" if sub and sub not in ['-', 'nan'] else task
                        return teks[:40] + "..." if len(teks) > 40 else teks

                    df_plot['Y_Label'] = df_plot.apply(format_ylabel, axis=1)
                    
                    today_date = pd.Timestamp(datetime.now().date())
                    
                    def tentukan_status(row):
                        remark = str(row.get('Remark Status', '')).lower()
                        end_date = row['End_Parsed']
                        
                        if any(k in remark for k in ['done', 'selesai', 'complete']):
                            return "✅ Selesai"
                        elif end_date < today_date:
                            return "🚨 Overdue (Terlambat)"
                        elif any(k in remark for k in ['progress', 'jalan']):
                            return "🔄 On Progress"
                        else:
                            return "⏳ Pending / Belum Mulai"
                    
                    df_plot['Status Kategori'] = df_plot.apply(tentukan_status, axis=1)

                    df_plot['Task Lengkap'] = df_plot['Task']
                    df_plot['Sub-Task Lengkap'] = df_plot['Sub-Task']
                    
                    color_map = {
                        "✅ Selesai": "#2ecc71",
                        "🔄 On Progress": "#3498db",
                        "🚨 Overdue (Terlambat)": "#e74c3c",
                        "⏳ Pending / Belum Mulai": "#f1c40f"
                    }

                    dynamic_height = max(400, len(df_plot) * 35)

                    fig_tl = px.timeline(
                        df_plot, 
                        x_start="Start_Parsed", 
                        x_end="End_Parsed", 
                        y="Y_Label", 
                        color="Status Kategori",
                        color_discrete_map=color_map,
                        hover_data={
                            "Task Lengkap": True, "Sub-Task Lengkap": True, "Remark Status": True, "PIC": True,
                            "Start_Parsed": "|%d %b %Y", "End_Parsed": "|%d %b %Y", "Y_Label": False, "Status Kategori": False
                        }, 
                        title="Jadwal Pelaksanaan Project DSSP",
                        height=dynamic_height
                    )
                    
                    fig_tl.update_yaxes(autorange="reversed", title="")
                    fig_tl.update_xaxes(title="Tanggal Pengerjaan")
                    
                    fig_tl.add_vline(x=datetime.now(), line_width=2, line_dash="dash", line_color="red", annotation_text="Hari Ini", annotation_position="top left")
                    fig_tl.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=""), margin=dict(l=0, r=0, t=50, b=0))
                    
                    st.plotly_chart(fig_tl, use_container_width=True)
                    
                    st.markdown("### ⚠️ Pantauan Keterlambatan Task")
                    overdue_tasks = df_plot[df_plot['Status Kategori'] == "🚨 Overdue (Terlambat)"]
                    
                    if not overdue_tasks.empty:
                        st.error(f"🚨 Terdapat {len(overdue_tasks)} Sub-Task yang melewati deadline!")
                        st.dataframe(overdue_tasks[['Task', 'Sub-Task', end_col, 'Remark Status', 'PIC']].reset_index(drop=True), use_container_width=True)
                    else:
                        st.success("✅ Tepat Waktu! Seluruh pengerjaan masih dalam batas jadwal.")
                else:
                    st.warning("⚠️ Data tanggal pada Sheet Timeline tidak valid (Tidak dapat dibaca sistem).")
            else:
                st.warning("⚠️ Kolom 'Start' dan/atau 'Deadline Finish' tidak ditemukan di Sheet Timeline.")
        else:
            st.warning("⚠️ Sheet bernama 'Timeline' tidak ditemukan atau datanya kosong.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")
