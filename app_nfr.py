#!/usr/bin/env python3
"""
app_nfr.py — App lokal (Streamlit) untuk filter & siapkan proposal NFR lahan.

Cara pakai:
    streamlit run app_nfr.py

Upload file:
  1. NFR Master Lahan (export dari website) — wajib
  2. Template (enrichment: alamat, koordinat, SES, Red/Blue Ocean, Harga) — opsional tapi disarankan
  3. PIC (kontak per No SPBU) — opsional

Lalu: filter di sidebar, lihat tabel/peta/chart, klik Export Excel.
"""
import html
import io, os, sys, datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import folium
from folium.plugins import MarkerCluster, FastMarkerCluster

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from prepare_proposal import (load_master, load_template, load_pic,
                              build_dataframe, filter_df, build_proposal_workbook,
                              build_proposal_workbook_fast,
                              REGIONS, REGION_DIGIT)

# Region name normalization: Pengajuan file uses different naming than Master
REGION_NORM = {
    'SUMBAGUT': 'Sumbagut', 'SUMBAGSEL': 'Sumbagsel',
    'JBB': 'JBB', 'JBT': 'JBT',
    'JTM BALINUS': 'Jatimbalinus', 'JATIMBALINUS': 'Jatimbalinus',
    'KALIMANTAN': 'Kalimantan', 'SULAWESI': 'Sulawesi',
    'MALUKU PAPUA': 'Maluku Papua',
}

# Dummy SPBU codes — placeholder/test data, BUKAN SPBU asli. Selalu difilter.
# Format non-standar lainnya (mis. panjang bukan 7-9 digit) tetap diterima sebagai SPBU asli.
DUMMY_SPBU = {'5999995', '5999997', '5999999'}

DEFAULT_COLS = ['No SPBU', 'Kode Lahan', 'Nama Perusahaan', 'Provinsi', 'Kota/Kab',
                'Kecamatan', 'Desa/Kelurahan', 'Alamat', 'Tipe SPBU', 'Luas (M2)',
                'Latitude', 'Longitude', 'Road Type', 'SES', 'Red Ocean', 'Blue Ocean',
                'Status Sewa']


def _arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalisasi kolom mixed-type (campur int & str) jadi string.

    Kolom object yang isinya campuran tipe bikin pyarrow/st.dataframe error &
    lambat (fallback Arrow serialization tiap render ribuan baris). Kolom yang
    semua-str atau semua-numerik dibiarkan (tetap cepat).
    """
    for _c in df.columns:
        if df[_c].dtype == object:
            _nn = df[_c].dropna()
            if not _nn.empty:
                _has_str = any(isinstance(v, str) for v in _nn)
                _has_non = any(not isinstance(v, str) for v in _nn)
                if _has_str and _has_non:
                    df[_c] = df[_c].astype(str).str.strip()
    return df


def to_bytes(upload):
    """Baca isi file upload sekali -> bytes (None bila tidak ada)."""
    if upload is None:
        return None
    upload.seek(0)
    return upload.read()


@st.cache_data
def cached_load_master(data: bytes):
    return load_master(io.BytesIO(data))


@st.cache_data
def cached_load_template(data: bytes):
    return load_template(io.BytesIO(data))


@st.cache_data
def cached_load_pic(data: bytes):
    return load_pic(io.BytesIO(data))


@st.cache_data
def cached_build_dataframe(master_bytes, templ_bytes, pic_bytes):
    master = cached_load_master(master_bytes) if master_bytes else {}
    templ  = cached_load_template(templ_bytes) if templ_bytes else {}
    pic    = cached_load_pic(pic_bytes) if pic_bytes else {}
    return build_dataframe(master, templ, pic)


@st.cache_data
def occ_load_master(data: bytes) -> pd.DataFrame:
    """Load NFR Master Lahan untuk tab Occupancy — derive Region dari digit-1 No SPBU.
    Format non-standar diterima (SPBU asli); hanya kode dummy (5999995/5999997/5999999) yang difilter.
    """
    df = pd.read_excel(io.BytesIO(data), sheet_name='NFR Master Lahan')
    df.columns = df.columns.str.strip()

    # Find SPBU column
    _spbu_col = next((c for c in ['No SPBU', 'No. SPBU', 'SPBU'] if c in df.columns), None)
    if not _spbu_col:
        st.warning('Kolom "No SPBU" tidak ditemukan di NFR Master Lahan. Mengembalikan dataframe kosong.')
        return pd.DataFrame()

    # Normalize column name to 'No SPBU' for consistency
    if _spbu_col != 'No SPBU':
        df.rename(columns={_spbu_col: 'No SPBU'}, inplace=True)
    # Kunci No SPBU JADI STRING: file campur int (5999995) & str (5999997-, 4A50612).
    # Kolom mixed-type bikin st.dataframe/pyarrow lambat (fallback Arrow tiap render).
    df['No SPBU'] = df['No SPBU'].astype(str).str.strip()
    _spbu_str = df['No SPBU']
    # Filter hanya kode dummy (placeholder, bukan SPBU asli); format lain diterima apa adanya
    df = df[~_spbu_str.isin(DUMMY_SPBU)].copy()
    # Derive Region dari digit-1 No SPBU
    digit_to_region = {str(i + 1): r for i, r in enumerate(REGIONS)}
    df['Region'] = df['No SPBU'].str[0].map(digit_to_region)
    return _arrow_safe(df)


@st.cache_data
def occ_load_pengajuan(data: bytes) -> pd.DataFrame:
    """Load Pengajuan NFR Izin Prinsip — derive Region from No. SPBU if missing."""
    df = pd.read_excel(io.BytesIO(data))
    df.columns = df.columns.str.strip()

    # Find SPBU column
    _spbu_col = next((c for c in ['No. SPBU', 'No SPBU', 'SPBU'] if c in df.columns), None)

    if 'Region' not in df.columns and _spbu_col:
        # Derive Region from first digit of SPBU number (same logic as master)
        digit_to_region = {str(i + 1): r for i, r in enumerate(REGIONS)}
        df['Region'] = df[_spbu_col].astype(str).str.strip().str[0].map(digit_to_region)
    elif 'Region' in df.columns:
        # Normalize existing Region column
        df['Region'] = (df['Region'].astype(str).str.upper().str.strip()
                        .map(REGION_NORM)
                        .fillna(df['Region']))

    return _arrow_safe(df)


@st.cache_data
def occ_load_berakhir(data: bytes) -> pd.DataFrame:
    """Load Pengajuan NFR Telah Berakhir — derive Region from No. SPBU if missing."""
    df = pd.read_excel(io.BytesIO(data))
    df.columns = df.columns.str.strip()

    # Find SPBU column
    _spbu_col = next((c for c in ['No. SPBU', 'No SPBU', 'SPBU'] if c in df.columns), None)

    if 'Region' not in df.columns and _spbu_col:
        # Derive Region from first digit of SPBU number (same logic as master)
        digit_to_region = {str(i + 1): r for i, r in enumerate(REGIONS)}
        df['Region'] = df[_spbu_col].astype(str).str.strip().str[0].map(digit_to_region)
    elif 'Region' in df.columns:
        # Normalize existing Region column
        df['Region'] = (df['Region'].astype(str).str.upper().str.strip()
                        .map(REGION_NORM)
                        .fillna(df['Region']))

    return _arrow_safe(df)


@st.cache_data
def detect_duplicates(master_bytes: bytes) -> pd.DataFrame:
    """Deteksi duplikat Kode Lahan dari Master Lahan FULL.

    Sinyal A: Kode Lahan muncul >1× → klasifikasi:
      - Identik       : semua baris punya Luas+Harga+Status Sewa sama (pure copy/bug)
      - Beda Nilai    : kode sama tapi Luas/Harga berbeda (reinput tanpa hapus lama)
      - Mixed Status  : 1 kode muncul sbg Tersedia sekaligus Tersewa (bug logika paling berbahaya)
    Sinyal B: Status Approval mengandung 'Ubah' → edit pending (versi baru tanpa hapus lama)

    Mengembalikan DataFrame semua baris duplikat + kolom Tipe Duplikat /
    Jumlah Kemunculan / Catatan / Sinyal. Baris diurutkan per Kode Lahan.
    """
    df = occ_load_master(master_bytes)  # derive Region, No SPBU apa adanya

    # hanya baris dengan Kode Lahan (exclude pending '-' yang belum punya kode)
    _kode_col = next((c for c in ['Kode Lahan'] if c in df.columns), None)
    _luas_col = next((c for c in ['Luas (M2)'] if c in df.columns), None)
    _harga_col = next((c for c in ['Harga (Rp)'] if c in df.columns), None)
    _status_col = next((c for c in ['Status Sewa'] if c in df.columns), None)
    _appr_col = next((c for c in ['Status Approval'] if c in df.columns), None)

    if not _kode_col:
        return pd.DataFrame(columns=['Tipe Duplikat', 'Jumlah Kemunculan', 'Sinyal', 'Catatan'])

    has_kode = df[_kode_col].astype(str).str.strip().ne('') & \
               df[_kode_col].notna() & \
               df[_kode_col].astype(str).str.strip().ne('nan')
    df_k = df[has_kode].copy()

    # Kode yang dibersihkan — dipakai untuk grouping & klasifikasi (O(n), tanpa scan per-kode)
    _strip_s = df_k[_kode_col].astype(str).str.strip()
    df_k['_kode_strip'] = _strip_s

    # --- Sinyal A: Kode Lahan (stripped) muncul >1 ---
    counts = df_k.groupby('_kode_strip')['_kode_strip'].transform('size')
    dup_mask_a = counts > 1
    # --- Sinyal B: Status Approval mengandung 'Ubah' ---
    ubah_mask = (_appr_col and
                 df_k[_appr_col].astype(str).str.contains('Ubah', case=False, na=False))
    dup_mask = dup_mask_a | (ubah_mask.fillna(False))

    dup = df_k[dup_mask].copy()
    if dup.empty:
        return dup.assign(**{
            'Tipe Duplikat': pd.Series(dtype=str),
            'Jumlah Kemunculan': pd.Series(dtype=int),
            'Sinyal': pd.Series(dtype=str),
            'Catatan': pd.Series(dtype=str),
        })

    dup['Jumlah Kemunculan'] = dup.groupby('_kode_strip')['_kode_strip'].transform('size')

    # --- klasifikasi per kode: SATU groupby (bukan scan 31k baris per kode) ---
    agg = {}
    if _status_col:
        agg['_status_n'] = (_status_col, 'nunique')
    if _luas_col:
        agg['_luas_set'] = (_luas_col, lambda s: frozenset(s.fillna('__NULL__').astype(str).str.strip()))
    if _harga_col:
        agg['_harga_set'] = (_harga_col, lambda s: frozenset(s.fillna('__NULL__').astype(str).str.strip()))
    agg['_size'] = ('_kode_strip', 'size')
    grp = dup.groupby('_kode_strip').agg(**agg)

    def _tipe_for(kode):
        g = grp.loc[kode]
        if g['_size'] <= 1:
            return 'Ubah Pending'  # hanya muncul 1× tapi ada edit pending
        if _status_col and g['_status_n'] > 1:
            return 'Mixed Status'
        l_uniq = len(g['_luas_set']) if _luas_col else 0
        h_uniq = len(g['_harga_set']) if _harga_col else 0
        if l_uniq <= 1 and h_uniq <= 1:
            return 'Identik'
        return 'Beda Nilai'

    dup['Tipe Duplikat'] = dup['_kode_strip'].map(_tipe_for)

    # Sinyal: A (Kode >1×) dan/atau B (Ubah Pending)
    _a_flags = dup_mask_a.reindex(dup.index).fillna(False).astype(bool)
    if ubah_mask is not None and bool(ubah_mask.any()):
        _b_flags = ubah_mask.reindex(dup.index).fillna(False).astype(bool)
        def _sinyal(a, b):
            parts = []
            if a: parts.append('A: Kode >1×')
            if b: parts.append('B: Ubah Pending')
            return ' + '.join(parts) if parts else 'A: Kode >1×'
        dup['Sinyal'] = [_sinyal(a, b) for a, b in zip(_a_flags, _b_flags)]
    else:
        dup['Sinyal'] = _a_flags.map(lambda a: 'A: Kode >1×' if a else 'B: Ubah Pending')
    dup['Sinyal'] = dup['Sinyal'].replace('', 'A: Kode >1×')

    # Catatan: sort unik Luas & Harga juga dari satu groupby (precompute), tanpa scan ulang
    _notes_luas = {}   # kode -> sorted list Luas
    _notes_harga = {}  # kode -> sorted list Harga
    if _luas_col:
        _luas_by_k = dup.groupby('_kode_strip')[_luas_col].apply(
            lambda s: sorted(set(s.fillna('(kosong)').astype(str).str.strip()))).to_dict()
        _notes_luas = _luas_by_k
    if _harga_col:
        _harga_by_k = dup.groupby('_kode_strip')[_harga_col].apply(
            lambda s: sorted(set(s.fillna('(kosong)').astype(str).str.strip()))).to_dict()
        _notes_harga = _harga_by_k

    def _catatan(row):
        tipe = row['Tipe Duplikat']
        kode = row['_kode_strip']
        if tipe == 'Mixed Status':
            return ('Bug logika: 1 Kode Lahan muncul sebagai Tersedia + Tersewa. '
                    'Prioritas cleansing tertinggi.')
        if tipe == 'Identik':
            return 'Kemungkinan duplikat murni (copy/sistem). Simpan 1 baris, hapus sisanya.'
        if tipe == 'Beda Nilai':
            luas = _notes_luas.get(kode, [])
            harga = _notes_harga.get(kode, [])
            return (f'Reinput: Luas bervariasi {luas} | Harga bervariasi {harga}. '
                    f'Verifikasi mana yang valid.')
        if tipe == 'Ubah Pending':
            return 'Edit pending (Ubah): versi baru dibuat tanpa hapus lama. Cek approval.'
        return ''

    dup['Catatan'] = dup.apply(_catatan, axis=1)
    # bungkus ulang agar kolom '_kode_strip' tidak bocor ke output
    dup = dup.drop(columns=['_kode_strip'], errors='ignore')
    # urutkan per Kode Lahan, Identik/Mixed lebih dulu
    tipe_order = {'Mixed Status': 0, 'Beda Nilai': 1, 'Identik': 2, 'Ubah Pending': 3}
    dup['_sort_tipe'] = dup['Tipe Duplikat'].map(tipe_order).fillna(9)
    dup = dup.sort_values(['_sort_tipe', _kode_col]).drop(columns=['_sort_tipe'])
    return dup


@st.cache_data
def load_ddms(data: bytes) -> pd.DataFrame:
    """Load DDMS / List Lembaga Penyalur — master semua SPBU se-Indonesia.

    Auto-detect kolom: SPBU (AgenNo/No. SPBU), Nama, Provinsi, Kota, Tipe SPBU,
    ProgressInput, Status Operasi. Filter SPBU Reguler saja. Derive Region dari digit-1.
    Output hanya kolom standar (hindari kolom duplikat code+name dari DDMS).
    """
    df = pd.read_excel(io.BytesIO(data))
    df.columns = df.columns.str.strip()

    _spbu_col = next((c for c in ['AgenNo', 'No. SPBU', 'No SPBU', 'SPBU'] if c in df.columns), None)
    if _spbu_col is None:
        return pd.DataFrame()

    # Filter hanya kode dummy (placeholder); format non-standar lain tetap diterima (SPBU asli)
    _spbu_str = df[_spbu_col].astype(str).str.strip()
    df = df[(_spbu_str.str.len() > 0) & (~_spbu_str.isin(DUMMY_SPBU))].copy()

    # Filter SPBU Reguler jika ada kolom tipe
    _tipe_col = next((c for c in ['NamaTipeLembagaPenyalur', 'Type Lembaga Penyalur', 'Tipe SPBU']
                      if c in df.columns), None)
    if _tipe_col:
        df = df[df[_tipe_col].astype(str).str.strip().str.lower() == 'spbu reguler'].copy()

    # Derive Region dari digit-1 No SPBU
    digit_to_region = {str(i + 1): r for i, r in enumerate(REGIONS)}
    _region_s = df[_spbu_col].astype(str).str[0].map(digit_to_region)

    # Build output bersih — pilih source column terbaik per nama standar
    # (DDMS punya pasangan code+name: Provinsi(code)+NamaProvinsi(name),
    #  Kota(code)+NamaKota(name) — ambil yang NAME agar tidak duplikat)
    def _pick(opts):
        for c in opts:
            if c in df.columns:
                return df[c]
        return pd.Series(dtype=object, index=df.index)

    out = pd.DataFrame(index=df.index)
    out['No SPBU']            = df[_spbu_col]
    out['Nama SPBU']          = _pick(['NamaLembagaPenyalur', 'Nama SPBU', 'Nama Perusahaan'])
    out['Provinsi']           = _pick(['NamaProvinsi', 'Provinsi'])
    out['Kota/Kab']          = _pick(['NamaKota', 'Kota/Kabupaten', 'Kota/Kab', 'Kota'])
    out['Region']             = _region_s
    out['Progress Input (%)'] = _pick(['ProgressInput', 'Progress Input (%)', 'Progress Input'])
    out['Status Operasi']    = _pick(['StatusOperasi', 'Status Operasi', 'Status Oper'])
    out['Latitude']          = _pick(['Coordinate_Y', 'Koordinat Y', 'Latitude'])
    out['Longitude']         = _pick(['Coordinate_X', 'Koordinat X', 'Longitude'])
    return out


@st.cache_data
def detect_tenant_issues(master_bytes: bytes) -> pd.DataFrame:
    """Deteksi masalah data tenant di Master Lahan FULL (hanya baris Tersewa).

    Lima tipe temuan (non-eksklusif — 1 baris bisa punya beberapa):
      - Kontrak Expired  : Tgl Berakhir < hari ini (stale — masih ditagih tersewa)
      - Expiring ≤90 hari: hari ini ≤ Tgl Berakhir ≤ hari ini+90 (perlu perpanjangan)
      - Tenant Kosong    : Tenant '-' atau null padahal Status Tersewa (inkonsistensi)
      - Harga Sewa Kosong: Harga Sewa (Rp) 0/null (nilai kontrak tidak diketahui)
      - Tgl Sewa Kosong  : Tgl Mulai atau Tgl Berakhir null (tanggal sewa hilang)

    Mengembalikan DataFrame baris Tersewa + kolom Tipe Temuan (comma-joined) /
    Status Temuan (prioritas tertinggi) / Catatan. Diurutkan per prioritas.
    """
    df = occ_load_master(master_bytes)  # derive Region, No SPBU apa adanya

    _status_col = next((c for c in ['Status Sewa'] if c in df.columns), None)
    _tenant_col = next((c for c in ['Tenant'] if c in df.columns), None)
    _hargasewa_col = next((c for c in ['Harga Sewa (Rp)'] if c in df.columns), None)
    _tgl_mulai_col = next((c for c in ['Tanggal Mulai Sewa', 'Tgl Mulai Sewa Lahan',
                                        'Tgl Mulai Sewa'] if c in df.columns), None)
    _tgl_akhir_col = next((c for c in ['Tanggal Berakhir Sewa', 'Tgl Akhir Sewa Lahan',
                                        'Tgl Berakhir Sewa'] if c in df.columns), None)
    _kode_col = next((c for c in ['Kode Lahan'] if c in df.columns), None)

    if not _status_col:
        return pd.DataFrame(columns=['Tipe Temuan', 'Status Temuan', 'Catatan'])

    # hanya baris Tersewa (punya tenant)
    tersewa_mask = df[_status_col].astype(str).str.strip() == 'Tersewa'
    df_t = df[tersewa_mask].copy()
    if df_t.empty:
        return df_t.assign(**{
            'Tipe Temuan': pd.Series(dtype=str),
            'Status Temuan': pd.Series(dtype=str),
            'Catatan': pd.Series(dtype=str),
        })

    today = datetime.date.today()
    horizon = today + datetime.timedelta(days=90)

    # parse tanggal
    _akhir_dates = pd.to_datetime(df_t[_tgl_akhir_col], errors='coerce') if _tgl_akhir_col else None
    _mulai_dates = pd.to_datetime(df_t[_tgl_mulai_col], errors='coerce') if _tgl_mulai_col else None

    # --- flag per tipe ---
    f_expired = pd.Series(False, index=df_t.index)
    f_expiring = pd.Series(False, index=df_t.index)
    f_tgl_kosong = pd.Series(False, index=df_t.index)
    if _akhir_dates is not None:
        _akhir_d = _akhir_dates.dt.date
        f_expired = (_akhir_d < today) & _akhir_dates.notna()
        f_expiring = _akhir_dates.notna() & (_akhir_d >= today) & (_akhir_d <= horizon)
    # tgl sewa kosong: mulai atau akhir null
    if _tgl_mulai_col is not None and _tgl_akhir_col is not None:
        f_tgl_kosong = _mulai_dates.isna() | _akhir_dates.isna()

    f_tenant_kosong = pd.Series(False, index=df_t.index)
    if _tenant_col:
        _t = df_t[_tenant_col].astype(str).str.strip()
        f_tenant_kosong = _t.isin(['-', '', 'nan', 'None'])

    f_harga_kosong = pd.Series(False, index=df_t.index)
    if _hargasewa_col:
        _h = df_t[_hargasewa_col]
        f_harga_kosong = _h.isna() | (_h.fillna(0).astype(float) == 0)

    # --- bangun Tipe Temuan (comma-joined) — lewat matriks bool, tanpa iterrows ---
    _flag_cols = {
        'Kontrak Expired': f_expired,
        'Expiring ≤90 hari': f_expiring,
        'Tenant Kosong': f_tenant_kosong,
        'Harga Sewa Kosong': f_harga_kosong,
        'Tgl Sewa Kosong': f_tgl_kosong,
    }
    _labels = list(_flag_cols)
    _m = pd.concat([s.rename(lab) for lab, s in _flag_cols.items()], axis=1).fillna(False).to_numpy(dtype=bool)
    _tipe_arr = [
        ', '.join(_labels[i] for i in range(len(_labels)) if row[i]) if row.any() else 'Lainnya'
        for row in _m
    ]
    df_t['Tipe Temuan'] = _tipe_arr

    # --- Status Temuan: prioritas tertinggi ---
    prio_order = ['Kontrak Expired', 'Expiring ≤90 hari', 'Tenant Kosong',
                  'Harga Sewa Kosong', 'Tgl Sewa Kosong', 'Lainnya']
    prio_rank = {t: i for i, t in enumerate(prio_order)}
    def _status(tipes):
        for t in prio_order:
            if t in tipes:
                return t
        return 'Lainnya'
    df_t['Status Temuan'] = df_t['Tipe Temuan'].map(_status)

    # --- Catatan (vectorized: numpy array posisi, tanpa apply/loc) ---
    _expired_v  = f_expired.values
    _expiring_v = f_expiring.values
    _tenant_v   = f_tenant_kosong.values
    _harga_v    = f_harga_kosong.values
    _tgl_v      = f_tgl_kosong.values
    # Pakai Series + .iloc (bukan .values) — .values memberi numpy.datetime64 yang
    # tak punya .date(); .iloc memberi pandas Timestamp seperti .loc semula.
    _akhir_s    = _akhir_dates if _akhir_dates is not None else None
    _n = len(f_expired)
    _notes_arr = []
    for i in range(_n):
        notes = []
        if _expired_v[i]:
            d = _akhir_s.iloc[i] if _akhir_s is not None else None
            notes.append(f"Kontrak berakhir {d.date() if pd.notna(d) else '?'} — sudah lewat, seharusnya Tersedia.")
        if _expiring_v[i]:
            d = _akhir_s.iloc[i] if _akhir_s is not None else None
            notes.append(f"Kontrak berakhir {d.date() if pd.notna(d) else '?'} — ≤90 hari, perlu perpanjangan.")
        if _tenant_v[i]:
            notes.append("Status Tersewa tapi nama Tenant kosong — inkonsistensi data.")
        if _harga_v[i]:
            notes.append("Harga Sewa (Rp) kosong/0 — nilai kontrak tidak diketahui.")
        if _tgl_v[i]:
            notes.append("Tanggal Mulai/Berakhir Sewa kosong — data sewa tidak lengkap.")
        _notes_arr.append(' '.join(notes) if notes else '')
    df_t['Catatan'] = _notes_arr

    # urutkan per prioritas lalu Tgl Berakhir
    df_t['_rank'] = df_t['Status Temuan'].map(prio_rank).fillna(99)
    _sort_by = ['_rank']
    if _tgl_akhir_col:
        df_t['_tgl_sort'] = pd.to_datetime(df_t[_tgl_akhir_col], errors='coerce')
        _sort_by.append('_tgl_sort')
    df_t = df_t.sort_values(_sort_by).drop(columns=['_rank'] + (['_tgl_sort'] if '_tgl_sort' in df_t.columns else []))
    return df_t


@st.cache_data
def search_beroperasi_keyword(beroperasi_bytes: bytes, keyword: str) -> pd.DataFrame:
    """Search keyword across Nama Tenant, Nama Brand, Kategori, Sub Kategori in Beroperasi.

    Returns DataFrame of matching contracts with all columns.
    Case-insensitive search, handles null values.
    """
    df = occ_load_pengajuan(beroperasi_bytes)

    keyword_lower = keyword.lower().strip()
    if not keyword_lower:
        return pd.DataFrame()

    # Search across 4 columns
    _tenant_col = next((c for c in ['Nama Tenant', 'Tenant'] if c in df.columns), None)
    _brand_col = next((c for c in ['Nama Brand', 'Brand'] if c in df.columns), None)
    _kategori_col = next((c for c in ['Kategori', 'Category'] if c in df.columns), None)
    _subkat_col = next((c for c in ['Sub Kategori', 'Sub Category'] if c in df.columns), None)
    _spbu_col = next((c for c in ['No. SPBU', 'No SPBU', 'SPBU'] if c in df.columns), None)

    matches = pd.Series(False, index=df.index)

    if _tenant_col:
        matches |= df[_tenant_col].fillna('').astype(str).str.lower().str.contains(keyword_lower, na=False)
    if _brand_col:
        matches |= df[_brand_col].fillna('').astype(str).str.lower().str.contains(keyword_lower, na=False)
    if _kategori_col:
        matches |= df[_kategori_col].fillna('').astype(str).str.lower().str.contains(keyword_lower, na=False)
    if _subkat_col:
        matches |= df[_subkat_col].fillna('').astype(str).str.lower().str.contains(keyword_lower, na=False)

    result = df[matches].copy()

    # Standardize column names for easier access
    if _spbu_col and _spbu_col != 'No SPBU':
        result['No SPBU'] = result[_spbu_col]

    return result


@st.cache_data
def detect_space_conflicts(master_bytes: bytes, beroperasi_bytes: bytes, keyword: str) -> dict:
    """Cross-reference Master Lahan Tersedia spaces vs Beroperasi active contracts.

    Returns dict with:
        - clean: DataFrame of clean available spaces (no conflicts)
        - conflicts: DataFrame of conflicted spaces (Tersedia but has active contract)
        - stats: dict with counts and percentages
        - beroperasi_matches: DataFrame of matching Beroperasi contracts
    """
    # Load Master Lahan
    df_master = occ_load_master(master_bytes)

    # Filter for Tersedia spaces only
    _status_col = next((c for c in ['Status Sewa'] if c in df_master.columns), None)
    if not _status_col:
        return {'clean': pd.DataFrame(), 'conflicts': pd.DataFrame(),
                'stats': {}, 'beroperasi_matches': pd.DataFrame()}

    tersedia = df_master[df_master[_status_col].astype(str).str.strip() == 'Tersedia'].copy()

    # Search Beroperasi for keyword matches
    beroperasi_matches = search_beroperasi_keyword(beroperasi_bytes, keyword)

    if beroperasi_matches.empty:
        # No matches in Beroperasi - all Tersedia spaces are clean
        return {
            'clean': tersedia,
            'conflicts': pd.DataFrame(),
            'stats': {
                'total_tersedia': len(tersedia),
                'clean_count': len(tersedia),
                'conflict_count': 0,
                'clean_pct': 100.0,
                'keyword': keyword
            },
            'beroperasi_matches': beroperasi_matches
        }

    # Aggregate Beroperasi contracts by SPBU
    _spbu_col = 'No SPBU'
    _brand_col = next((c for c in ['Nama Brand', 'Brand'] if c in beroperasi_matches.columns), None)
    _kategori_col = next((c for c in ['Kategori'] if c in beroperasi_matches.columns), None)
    _tgl_akhir_col = next((c for c in ['Tgl Akhir Sewa Lahan', 'Tanggal Berakhir Sewa']
                                     if c in beroperasi_matches.columns), None)
    _harga_col = next((c for c in ['Total Harga Sewa', 'Harga Sewa']
                                 if c in beroperasi_matches.columns), None)

    agg_dict = {'_count': (_spbu_col, 'size')}
    if _brand_col:
        agg_dict['Brands'] = (_brand_col, lambda x: ', '.join(x.dropna().astype(str).unique()[:3]))
    if _kategori_col:
        agg_dict['Kategori'] = (_kategori_col, 'first')
    if _tgl_akhir_col:
        agg_dict['Latest_End_Date'] = (_tgl_akhir_col, 'max')
    if _harga_col:
        agg_dict['Total_Value'] = (_harga_col, 'sum')

    spbu_contracts = beroperasi_matches.groupby(_spbu_col).agg(**agg_dict).reset_index()
    spbu_contracts.rename(columns={'_count': 'Contract_Count'}, inplace=True)

    # Check if contracts are expired (Option B: include but flag)
    if _tgl_akhir_col and 'Latest_End_Date' in spbu_contracts.columns:
        today = datetime.date.today()
        end_dates = pd.to_datetime(spbu_contracts['Latest_End_Date'], errors='coerce')
        spbu_contracts['Is_Expired'] = end_dates.dt.date < today
    else:
        spbu_contracts['Is_Expired'] = False

    # Join with Tersedia spaces to find conflicts
    tersedia['No SPBU'] = tersedia['No SPBU'].astype(str).str.strip()
    spbu_contracts['No SPBU'] = spbu_contracts['No SPBU'].astype(str).str.strip()

    conflicts = tersedia.merge(
        spbu_contracts,
        on='No SPBU',
        how='inner'
    )

    # Clean spaces = Tersedia spaces NOT in conflict list
    clean = tersedia[~tersedia['No SPBU'].isin(conflicts['No SPBU'])].copy()

    # Calculate stats
    total = len(tersedia)
    clean_count = len(clean)
    conflict_count = len(conflicts)
    clean_pct = (clean_count / total * 100) if total > 0 else 0.0

    return {
        'clean': clean,
        'conflicts': conflicts,
        'stats': {
            'total_tersedia': total,
            'clean_count': clean_count,
            'conflict_count': conflict_count,
            'clean_pct': clean_pct,
            'keyword': keyword,
            'expired_contracts': int(spbu_contracts['Is_Expired'].sum()) if 'Is_Expired' in spbu_contracts.columns else 0
        },
        'beroperasi_matches': beroperasi_matches
    }


@st.cache_data
def to_excel_bytes(records, data_cols, reg_records=None, kota_records=None, spbu_records=None):
    """Cache on records (list of dicts) + cols tuple — both hashable."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    df_export = pd.DataFrame(records)

    # Data besar (ribuan baris): pakai xlsxwriter (~2x lbh cepat) — sheet Data saja,
    # tanpa Charts/Peta/styling per-sel/summary yang bikin openpyxl lambat 5-40 dtk.
    if len(df_export) > 6000:
        return build_proposal_workbook_fast(df_export, list(data_cols))

    buf = io.BytesIO()
    wb = build_proposal_workbook(df_export, list(data_cols))

    # styling helpers for summary sheets
    BLUE_HDR  = '0070C0'
    hdr_fill  = PatternFill('solid', fgColor=BLUE_HDR)
    hdr_font  = Font(color='FFFFFF', bold=True, size=11, name='Calibri')
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    dat_font  = Font(size=11, name='Calibri')
    dat_align = Alignment(horizontal='left', vertical='center')
    thin      = Side(border_style='thin', color='BDD7EE')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _write_summary_sheet(wb, title, records_list):
        if not records_list:
            return
        ws = wb.create_sheet(title)
        ws.sheet_view.showGridLines = False
        cols = list(records_list[0].keys())
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=col)
            c.fill = hdr_fill; c.font = hdr_font
            c.alignment = hdr_align; c.border = border
            # auto-width: header length or max data length, cap at 40
            max_len = max(len(str(col)), max((len(str(r.get(col, ''))) for r in records_list), default=0))
            ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 40)
        ws.row_dimensions[1].height = 28
        for row in records_list:
            ws.append([row.get(col) for col in cols])
        # Style per-cell data hanya bila baris sedikit (sama seperti sheet Data).
        # Sheet ringkasan SPBU bisa ribuan baris -> styling per sel sangat lambat.
        if len(records_list) <= 6000:
            for ri in range(2, len(records_list) + 2):
                for ci in range(1, len(cols) + 1):
                    c = ws.cell(row=ri, column=ci)
                    c.font = dat_font; c.alignment = dat_align; c.border = border
                ws.row_dimensions[ri].height = 15
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions

    if reg_records:
        _write_summary_sheet(wb, 'Ringkasan Region', reg_records)
    if kota_records:
        _write_summary_sheet(wb, 'Ringkasan Kota', kota_records)
    if spbu_records:
        _write_summary_sheet(wb, 'Ringkasan SPBU', spbu_records)

    wb.save(buf)
    return buf.getvalue()


@st.cache_data
def space_check_to_excel(clean_records, conflict_records, keyword):
    """Simple Excel export for Space Check results."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)

    # Styling
    hdr_fill = PatternFill('solid', fgColor='0070C0')
    hdr_font = Font(color='FFFFFF', bold=True, size=11, name='Calibri')
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    dat_font = Font(size=11, name='Calibri')
    dat_align = Alignment(horizontal='left', vertical='center')
    thin = Side(border_style='thin', color='BDD7EE')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_sheet(wb, title, records_list):
        if not records_list:
            ws = wb.create_sheet(title)
            ws.append(['No data'])
            return

        ws = wb.create_sheet(title)
        ws.sheet_view.showGridLines = False

        cols = list(records_list[0].keys())
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=col)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = hdr_align
            c.border = border
            max_len = max(len(str(col)), max((len(str(r.get(col, ''))) for r in records_list), default=0))
            ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 40)

        ws.row_dimensions[1].height = 28

        for ri, row in enumerate(records_list, 2):
            for ci, col in enumerate(cols, 1):
                c = ws.cell(row=ri, column=ci, value=row.get(col))
                c.font = dat_font
                c.alignment = dat_align
                c.border = border
            ws.row_dimensions[ri].height = 15

        ws.freeze_panes = 'A2'
        if len(records_list) > 0:
            ws.auto_filter.ref = ws.dimensions

    write_sheet(wb, f'Clean ({len(clean_records)})', clean_records)
    write_sheet(wb, f'Conflicts ({len(conflict_records)})', conflict_records)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_export_filename(region, statuses, provinsi, kota):
    parts = ['available_space_nfr']
    if region:
        parts.append('-'.join(region[:2]))
    if provinsi:
        parts.append('-'.join(p.replace(' ', '_') for p in provinsi[:2]))
    if kota:
        parts.append('-'.join(k.replace(' ', '_') for k in kota[:2]))
    if statuses and set(statuses) != {'Tersedia'}:
        parts.append('_'.join(statuses))
    parts.append(datetime.date.today().strftime('%Y-%m-%d'))
    return '_'.join(parts) + '.xlsx'


def fmt_rupiah(v):
    if v is None or (isinstance(v, float) and v != v):
        return '-'
    try:
        return 'Rp {:,.0f}'.format(float(v)).replace(',', '.')
    except (ValueError, TypeError):
        return '-'


def _parse_num(s):
    if s is None:
        return None
    cleaned = str(s).replace('.', '').replace(',', '').strip()
    try:
        return float(cleaned) if cleaned else None
    except ValueError:
        return None


FILTER_KEYS = ['f_region', 'f_statuses', 'f_min_luas', 'f_max_luas',
               'f_min_harga', 'f_max_harga', 'f_usaha', 'f_provinsi', 'f_kota', 'f_tipe_spbu',
               'f_throughput_col', 'f_min_throughput', 'f_transaksi_col', 'f_min_transaksi',
               'f_show_cols', 'f_sort_cols', 'f_sort_desc']

def reset_filters():
    for k in FILTER_KEYS:
        if k in st.session_state:
            del st.session_state[k]


# ---- methodology expander helper ----
METODOLOGI = {
    'occupancy': {
        'judul': 'Metodologi Occupancy Rate',
        'rumus': 'Occupancy = Tersewa / (Tersedia + Tersewa) × 100',
        'sumber': [
            '**NFR Master Lahan FULL** — kolom `Status Sewa` (Tersedia / Tersewa)',
            'Region di-derive dari digit-1 No SPBU (1=Sumbagut, 2=Sumbagsel, 3=JBB, 4=JBT, 5=Jatimbalinus, 6=Kalimantan, 7=Sulawesi, 8=Maluku Papua)',
        ],
        'edge': [
            'Dummy SPBU (5999995/5999997/5999999) difilter — bukan SPBU asli. Format non-standar lain diterima.',
            'Lahan dengan Status Sewa "-" (pending, belum punya Kode Lahan) dilaporkan terpisah, tidak masuk occupancy.',
            'Mode "Disesuaikan": lahan Tersewa dengan Tgl Berakhir sudah lewat dianggap Tersedia (data belum diupdate di master).',
            'Klasifikasi SPBU memerlukan file Telah Berakhir untuk membedakan Baru vs Kembali Tersedia.',
        ],
    },
    'cleansing': {
        'judul': 'Metodologi Deteksi Duplikat',
        'rumus': 'Dua sinyal independen: (A) Kode Lahan muncul >1×, (B) Status Approval mengandung "Ubah"',
        'sumber': [
            '**NFR Master Lahan FULL** — kolom `Kode Lahan`, `Luas (M2)`, `Harga (Rp)`, `Status Sewa`, `Status Approval`',
            'Dummy SPBU (5999995/5999997/5999999) difilter. Format non-standar lain diterima (SPBU asli).',
        ],
        'edge': [
            'Identik = Kode sama, Luas & Harga identik (termasuk null = null). Kemungkinan copy/bug sistem.',
            'Beda Nilai = Kode sama, Luas atau Harga berbeda (null ≠ terisi). Kemungkinan reinput tanpa hapus lama.',
            'Mixed Status = 1 Kode muncul sebagai Tersedia + Tersewa. Bug logika — prioritas tertinggi.',
            'Ubah Pending = Status Approval "Ubah - Menunggu/Ditolak/Revisi" tapi Kode muncul 1× saja. Sistem buat versi baru tanpa hapus lama.',
            'Baris dengan Kode Lahan null (Status "-") tidak dianalisis — belum punya identitas lahan.',
        ],
    },
    'progress': {
        'judul': 'Metodologi Progress Input SPBU',
        'rumus': 'Belum Input = SPBU Reguler di DDMS yang tidak ada di NFR Master Lahan',
        'sumber': [
            '**DDMS / List Lembaga Penyalur** — kolom `AgenNo`/`No. SPBU`, `NamaTipeLembagaPenyalur`, `ProgressInput`',
            '**NFR Master Lahan** — kolom `No SPBU` (untuk perbandingan)',
        ],
        'edge': [
            'Hanya SPBU dengan Tipe "SPBU Reguler" yang dianalisis.',
            'Dummy SPBU (5999995/5999997/5999999) difilter. Format non-standar lain diterima (SPBU asli).',
            'Progress Input (%) dari DDMS = kelengkapan data SPBU secara umum, bukan khusus lahan NFR.',
            'Gap "Belum Input" = SPBU Reguler ada di DDMS tapi tidak muncul di Master Lahan NFR sama sekali.',
        ],
    },
    'tenant': {
        'judul': 'Metodologi Cleansing Data Tenant',
        'rumus': 'Analisis hanya baris Status Sewa = Tersewa. 5 tipe temuan (non-eksklusif): Kontrak Expired (Tgl Berakhir < hari ini), Expiring ≤90 hari, Tenant Kosong, Harga Sewa Kosong, Tgl Sewa Kosong',
        'sumber': [
            '**NFR Master Lahan FULL** — kolom `Status Sewa`, `Tenant`, `Tanggal Mulai Sewa`, `Tanggal Berakhir Sewa`, `Harga Sewa (Rp)`',
            'Tanggal di-parse dengan `pd.to_datetime(errors="coerce")` — NaT dikecualikan.',
        ],
        'edge': [
            'Hanya baris Tersewa yang dianalisis (baris Tersedia/pending tidak punya tenant).',
            'Kontrak Expired = Tgl Berakhir < hari ini DAN Status masih Tersewa (data belum diupdate — seharusnya Tersedia).',
            'Expiring ≤90 hari = Tgl Berakhir antara hari ini dan 90 hari ke depan (perlu perpanjangan).',
            'Tenant Kosong = Status Tersewa tapi nama Tenant kosong/"-" (inkonsistensi data).',
            '1 baris bisa punya beberapa tipe (mis. Expired + Harga Kosong) → Tipe Temuan comma-joined, Status Temuan = prioritas tertinggi.',
            'Dummy SPBU (5999995/5999997/5999999) difilter. Format non-standar lain diterima (SPBU asli).',
        ],
    },
    'space_check': {
        'judul': 'Metodologi Space Check — Deteksi Konflik Ketersediaan',
        'rumus': 'Conflict = Lahan berstatus "Tersedia" di Master Lahan DAN SPBU punya kontrak aktif matching keyword di Beroperasi',
        'sumber': [
            '**NFR Master Lahan FULL** — kolom `No SPBU`, `Kode Lahan`, `Status Sewa` (filter "Tersedia")',
            '**Pengajuan Beroperasi** — kolom `No. SPBU`, `Nama Tenant`, `Nama Brand`, `Kategori`, `Sub Kategori`, `Tgl Akhir Sewa Lahan`',
        ],
        'edge': [
            'Search case-insensitive across 4 kolom di Beroperasi: Nama Tenant, Nama Brand, Kategori, Sub Kategori.',
            'Join pada No SPBU (one-to-many: satu SPBU bisa punya banyak lahan dan banyak kontrak).',
            'Jika SPBU punya 3 lahan Tersedia dan 1 kontrak matching keyword → SEMUA 3 lahan di-flag (tidak bisa tentukan lahan spesifik).',
            'Kontrak expired (Tgl Akhir < hari ini) tetap dianalisis dan di-flag sebagai "Expired contract - data issue".',
            'Clean Available = lahan Tersedia di SPBU yang TIDAK punya kontrak matching keyword.',
            'Dummy SPBU (5999995/5999997/5999999) difilter. Format non-standar lain diterima (SPBU asli).',
        ],
    },
}


def show_metodologi(key: str):
    """Tampilkan expander metodologi untuk tab analitik."""
    m = METODOLOGI.get(key)
    if not m:
        return
    with st.expander(f'ℹ️ {m["judul"]} — klik untuk lihat cara perhitungan', expanded=False):
        st.markdown(f"**Rumus:** `{m['rumus']}`")
        st.markdown('**Sumber data:**')
        for s in m['sumber']:
            st.markdown(f'- {s}')
        st.markdown('**Penanganan edge case:**')
        for e in m['edge']:
            st.markdown(f'- {e}')


# ---------------------------------------------------------------- UI
_ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon PPN.webp')
st.set_page_config(page_title='NFR Lahan Tool — Pertamina Patra Niaga', layout='wide',
                   page_icon=_ICON_PATH if os.path.exists(_ICON_PATH) else '⛽')

# Pertamina brand theme injection
st.markdown("""
<style>
/* ---- Brand variables ---- */
:root {
    --ptm-red:   #E31B23;
    --ptm-blue:  #0072C6;
    --ptm-green: #5A8A00;
    --ptm-dark:  #1A1A1A;
    --ptm-light: #F4F6F9;
    --ptm-border: #DDE3EC;
}

/* ---- Top header bar ---- */
[data-testid="stAppViewContainer"] > .main > .block-container {
    padding-top: 1.5rem;
}
header[data-testid="stHeader"] {
    background: #F4F6F9;
    height: 4px;
}

/* ---- Sidebar ---- */
[data-testid="stSidebar"] {
    background-color: var(--ptm-light) !important;
    border-right: 2px solid var(--ptm-border);
}

/* ---- Sidebar header text: dark label, red underline ---- */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: var(--ptm-dark) !important;
    font-size: 0.85rem !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    border-bottom: 2px solid var(--ptm-red);
    padding-bottom: 4px;
    margin-top: 1rem;
}

/* ---- Page title ---- */
h1 { color: var(--ptm-dark) !important; }
h1::after {
    content: '';
    display: block;
    width: 48px;
    height: 3px;
    background: var(--ptm-red);
    border-radius: 2px;
    margin-top: 6px;
}

/* ---- Primary buttons (Export, Reset) ---- */
[data-testid="stButton"] > button,
[data-testid="stDownloadButton"] > button {
    background-color: var(--ptm-red) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 4px !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em;
    transition: background 0.15s;
}
[data-testid="stButton"] > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    background-color: #b01019 !important;
}

/* ---- Metric cards ---- */
[data-testid="stMetric"] {
    background: #fff;
    border: 1px solid var(--ptm-border);
    border-top: 3px solid var(--ptm-red);
    border-radius: 6px;
    padding: 0.75rem 1rem !important;
}
[data-testid="stMetricValue"] {
    color: var(--ptm-dark) !important;
    font-weight: 700;
}
[data-testid="stMetricLabel"] {
    color: #666 !important;
    font-size: 0.78rem !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

/* ---- Tab bar ---- */
[data-testid="stTabs"] [role="tab"] {
    font-weight: 500;
    color: #555;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: var(--ptm-red) !important;
    border-bottom-color: var(--ptm-red) !important;
}

/* ---- Divider ---- */
hr { border-color: var(--ptm-border) !important; }

/* ---- Info / Warning banners ---- */
[data-testid="stAlert"] {
    border-radius: 4px;
}
</style>
""", unsafe_allow_html=True)

# Logo in sidebar header — hidden when collapsed, custom height via CSS
_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pertamina patra niaga logo.webp')
if os.path.exists(_LOGO_PATH):
    st.logo(_LOGO_PATH, size='large', icon_image=None)

st.markdown("""
<style>
/* Logo: 3rem tinggi, auto lebar, centered horizontal di sidebar header */
[data-testid="stLogoSidebar"],
[data-testid="stLogoSidebar"] img,
[data-testid="stLogo"],
[data-testid="stLogo"] img,
div[class*="logo"] img,
a[data-testid="stLogoLink"] img {
    height: 3rem !important;
    max-height: 3rem !important;
    min-height: 3rem !important;
    width: auto !important;
    display: block !important;
    margin-left: auto !important;
    margin-right: auto !important;
}
/* Container logo: center horizontal */
[data-testid="stLogoSidebar"],
[data-testid="stSidebar"] [data-testid="stLogo"],
div[class*="logo"] {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
}

/* ---- Desktop: sembunyikan tombol collapse sidebar (disable collapse) ---- */
/* Hanya target di layar lebar (desktop) — biarkan hamburger mobile tetap hidup */
@media (min-width: 769px) {
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stLogoCollapsed"] {
        display: none !important;
    }
}

/* ---- Mobile: top bar dengan icon kiri + hamburger kanan ---- */
/* Saat sidebar collapse di mobile, Streamlit render top bar dengan hamburger.
   Kita biarkan default (jangan display:none) supaya bisa toggle filter. */
@media (max-width: 768px) {
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {
        display: flex !important;
        align-items: center !important;
    }
    [data-testid="stLogoIcon"],
    [data-testid="stLogoCollapsed"] {
        display: flex !important;
        height: 2.5rem !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="margin-bottom:0.25rem;">
    <div style="font-size:1.5rem;font-weight:700;color:#1A1A1A;line-height:1.2;">
        NFR Lahan — Filter &amp; Proposal
    </div>
    <div style="font-size:0.82rem;color:#666;margin-top:2px;">
        Pertamina Patra Niaga &nbsp;·&nbsp; Upload export NFR Master Lahan, filter, lihat tabel/peta/chart, lalu Export Excel.
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header('File')
    master_up = st.file_uploader('1. NFR Master Lahan (xlsx)', type=['xlsx'],
                                  help='Export FULL dari website — harus berisi Tersedia + Tersewa. '
                                       'File ini dipakai SEMUA tab (proposal + occupancy + analytics).')
    templ_up = st.file_uploader('2. Template (opsional)', type=['xlsx'])
    pic_up = st.file_uploader('3. PIC (opsional)', type=['xlsx'],
                              help='Kolom: No SPBU / AgentNo, Nama (boleh kosong), No HP, Email, Phone Office. '
                                   'Nilai "NULL" dianggap kosong.')

    st.header('Filter')
    region = st.multiselect('Region', REGIONS, key='f_region',
                            help='Kosongkan = semua region.')
    statuses = st.multiselect('Status Sewa', ['Tersedia', 'Tersewa'], default=['Tersedia'], key='f_statuses')
    col1, col2 = st.columns(2)
    min_luas = col1.number_input('Luas min (m²)', 0, 100000, 0, step=10, key='f_min_luas')
    if min_luas:
        col1.caption(f"{min_luas:,.0f} m²".replace(',', '.'))
    max_luas = col2.number_input('Luas max (m²)', 0, 100000, 100000, step=10, key='f_max_luas')
    col2.caption(f"{max_luas:,.0f} m²".replace(',', '.'))
    col3, col4 = st.columns(2)
    min_harga = col3.number_input('Harga min (Rp)', 0, 10 ** 12, 0, step=500000, key='f_min_harga')
    if min_harga:
        col3.caption(fmt_rupiah(min_harga))
    max_harga = col4.number_input('Harga max (Rp)', 0, 10 ** 12, 10 ** 12, step=500000, key='f_max_harga')
    col4.caption(fmt_rupiah(max_harga))
    usaha = st.text_input('Jenis usaha (mis. coffee, grocery, atm)', key='f_usaha')

    with st.expander('Filter Lanjutan (Throughput & Transaksi)'):
        st.markdown('**Throughput BBM**')
        throughput_col = st.selectbox('Kolom throughput', ['Total', 'PSO', 'Non PSO'],
                                      key='f_throughput_col',
                                      help='PSO = BBM bersubsidi, Non PSO = BBM non-subsidi, Total = keduanya')
        min_throughput = st.number_input('Min throughput (kL/hari)', 0.0, 5000.0, 0.0,
                                         step=5.0, key='f_min_throughput',
                                         help='Data bulanan dibagi 30. Req umum: > 20 kL/hari.')
        if min_throughput:
            st.caption(f'= min {min_throughput * 30:,.0f} kL/bulan'.replace(',', '.'))

        st.markdown('**Transaksi / Bulan**')
        transaksi_col = st.selectbox('Kolom transaksi', ['Total2', 'PSO2', 'Non PSO2'],
                                     key='f_transaksi_col',
                                     help='PSO2/Non PSO2/Total2 = jumlah transaksi per bulan')
        min_transaksi = st.number_input('Min transaksi/bulan', 0, 1000000, 0,
                                        step=100, key='f_min_transaksi')
        if min_transaksi:
            st.caption(f'{min_transaksi:,} transaksi/bulan'.replace(',', '.'))

    st.button('Reset Filter', on_click=reset_filters, width='stretch')

# Tabs always rendered so Occupancy is visible even without master upload
tabs = st.tabs(['Dashboard', 'Tabel', 'Peta', 'Chart', 'Ringkasan', 'Space Check', 'Occupancy', 'Cleansing', 'Progress Input', 'Tenant'])

# html-escape helper used in map popup
def _e(v):
    s = str(v) if v is not None else ''
    return html.escape(s) if s and s != 'nan' else '-'


# ---- map performance helpers ----
MAP_MARKER_CAP = 10_000  # max markers rendered on the Peta tab; above this we sample

def _stratified_sample(df: pd.DataFrame, cap: int = MAP_MARKER_CAP) -> pd.DataFrame:
    """Deterministic proportional sample across Status Sewa x Region to respect MAP_MARKER_CAP.

    Uses groupby().head(n) (no Math.random / Date) so it is Streamlit-cache safe.
    Keeps the sample representative — tidak drop region / status secara sepihak.
    """
    if len(df) <= cap:
        return df
    total = len(df)
    # Bucket by Status Sewa + Region jika kolom ada; fallback ke single bucket
    keys = [c for c in ['Status Sewa', 'Region'] if c in df.columns]
    if not keys:
        return df.head(cap)
    parts = []
    for _, g in df.groupby(keys, dropna=False, sort=False):
        n = max(1, int(round(len(g) * cap / total)))
        parts.append(g.head(n))
    out = pd.concat(parts)
    return out.head(cap) if len(out) > cap else out

show_cols = []  # set inside tabs[1] when master data is available

if not master_up:
    for i in range(5):
        with tabs[i]:
            st.markdown("""
<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
            padding:3rem 1rem;text-align:center;">
  <div style="font-size:3rem;margin-bottom:1rem;">📂</div>
  <div style="font-size:1.1rem;font-weight:600;color:#1A1A1A;margin-bottom:0.5rem;">
    Belum ada data
  </div>
  <div style="font-size:0.875rem;color:#666;max-width:340px;line-height:1.6;">
    Upload <b>NFR Master Lahan</b> di sidebar kiri untuk mulai.<br>
    Template dan PIC bersifat opsional tapi disarankan.
  </div>
</div>
""", unsafe_allow_html=True)


if master_up:
    master_bytes = to_bytes(master_up)
    # Seed shared master for ALL analytics tabs (Occupancy, Space Check, Cleansing,
    # Progress Input, Tenant). Sidebar runs before st.tabs() so every tab sees this.
    st.session_state['_dash_occ_master'] = master_bytes
    templ_bytes  = to_bytes(templ_up)
    pic_bytes    = to_bytes(pic_up)

    if not templ_bytes:
        st.warning('Template belum diupload — kolom alamat/koordinat/harga/SES akan kosong. Disarankan upload Template.')

    with st.spinner('Memuat data...'):
        df = cached_build_dataframe(master_bytes, templ_bytes, pic_bytes)
    st.sidebar.markdown(f'**Total lahan (master):** {len(df):,}')

    # Cascading filter Provinsi + Kota/Kab
    with st.sidebar:
        if region:
            digits = [REGION_DIGIT[r] for r in region]
            if 'No SPBU' in df.columns:
                df_scoped = df[df['No SPBU'].astype(str).str[0].isin(digits)]
            else:
                df_scoped = df
        else:
            df_scoped = df

        if 'Provinsi' in df_scoped.columns:
            provinsi_options = sorted(df_scoped['Provinsi'].dropna().astype(str).unique())
        else:
            provinsi_options = []
        prev_prov = st.session_state.get('f_provinsi', [])
        if any(p not in provinsi_options for p in prev_prov):
            st.session_state['f_provinsi'] = [p for p in prev_prov if p in provinsi_options]

        provinsi = st.multiselect(
            f'Provinsi ({len(provinsi_options)} tersedia)',
            provinsi_options,
            help='Otomatis menyesuaikan Region. Kosongkan = semua.',
            key='f_provinsi',
        )

        if provinsi:
            df_kota_scope = df_scoped[df_scoped['Provinsi'].isin(provinsi)]
        else:
            df_kota_scope = df_scoped

        if 'Kota/Kab' in df_kota_scope.columns:
            kota_options = sorted(df_kota_scope['Kota/Kab'].dropna().astype(str).unique())
        else:
            kota_options = []
        prev_kota = st.session_state.get('f_kota', [])
        if any(k not in kota_options for k in prev_kota):
            st.session_state['f_kota'] = [k for k in prev_kota if k in kota_options]

        kota = st.multiselect(
            f'Kota/Kab ({len(kota_options)} tersedia)',
            kota_options,
            help='Otomatis menyesuaikan Provinsi & Region. Kosongkan = semua.',
            key='f_kota',
        )

        tipe_options = sorted(df['Tipe SPBU'].dropna().astype(str).unique())
        tipe_spbu = st.multiselect(
            f'Tipe SPBU ({len(tipe_options)} tersedia)',
            tipe_options,
            help='Mis. DODO, COCO, DODO PERTAMINI, dll. Kosongkan = semua tipe.',
            key='f_tipe_spbu',
        )

    result = filter_df(df, region, statuses, kota,
                       min_luas if min_luas else None, max_luas if max_luas < 100000 else None,
                       min_harga if min_harga else None, max_harga,
                       usaha, provinsi=provinsi if provinsi else None,
                       tipe_spbu=tipe_spbu if tipe_spbu else None,
                       throughput_col=throughput_col,
                       min_throughput_hari=min_throughput if min_throughput else None,
                       transaksi_col=transaksi_col,
                       min_transaksi_bulan=min_transaksi if min_transaksi else None)


    st.subheader(f'Hasil filter: {len(result):,} lokasi')

    _result_empty = result.empty
    if _result_empty:
        active = []
        if region:        active.append(f'Region: {", ".join(region)}')
        if provinsi:      active.append(f'Provinsi: {", ".join(provinsi)}')
        if kota:          active.append(f'Kota: {", ".join(kota)}')
        if statuses:      active.append(f'Status: {", ".join(statuses)}')
        if min_luas:      active.append(f'Luas min: {min_luas} m²')
        if min_harga:     active.append(f'Harga min: {fmt_rupiah(min_harga)}')
        if min_throughput: active.append(f'Throughput min: {min_throughput} kL/hari')
        if min_transaksi:  active.append(f'Transaksi min: {min_transaksi}/bulan')
        hint = ' · '.join(active) if active else 'semua filter aktif'
        st.markdown(f"""
<div style="background:#FFF8E1;border:1px solid #FFD54F;border-radius:6px;
            padding:1.25rem 1.5rem;margin:0.5rem 0;">
  <div style="font-size:1.5rem;margin-bottom:0.5rem;">🔍</div>
  <div style="font-weight:600;color:#1A1A1A;margin-bottom:0.25rem;">Tidak ada lokasi yang cocok</div>
  <div style="font-size:0.875rem;color:#555;">Filter aktif: {hint}</div>
</div>
""", unsafe_allow_html=True)
        st.button('Reset semua filter', on_click=reset_filters, type='primary')

    if not _result_empty:
        # ----- metrics -----
        m1, m2, m3, m4 = st.columns(4)
        m1.metric('Total Lokasi', f"{len(result):,}")
        avg_luas = result['Luas (M2)'].dropna()
        m2.metric('Rata-rata Luas', f"{avg_luas.mean():.0f} m²" if not avg_luas.empty else '-')
        avg_harga = result['Harga (Rp)'].dropna()
        m3.metric('Rata-rata Harga', fmt_rupiah(avg_harga.mean()) if not avg_harga.empty else '-')
        m4.metric('Kota/Kab Unik', f"{result['Kota/Kab'].nunique():,}")

        # Column list computed once — used in Tabel tab and export
        all_cols = [c for c in result.columns if c not in ('Latitude', 'Longitude')]
        if 'f_show_cols' not in st.session_state:
            st.session_state['f_show_cols'] = [c for c in DEFAULT_COLS if c in all_cols]
        st.session_state['f_show_cols'] = [c for c in st.session_state['f_show_cols'] if c in all_cols]

    with tabs[0]:
        # ---- Dashboard: full-dataset overview (unfiltered) ----
        DIGIT_REGION_D = {str(i + 1): r for i, r in enumerate(REGIONS)}
        df_dash = df.assign(Region=df['No SPBU'].astype(str).str[0]
                            .map(DIGIT_REGION_D).fillna('Lainnya'))

        total_lahan    = len(df_dash)
        tersedia_total = int((df_dash['Status Sewa'] == 'Tersedia').sum())
        tersewa_total  = int((df_dash['Status Sewa'] == 'Tersewa').sum())
        pct_tersedia   = round(tersedia_total / total_lahan * 100, 1) if total_lahan else 0
        avg_luas_all   = df_dash['Luas (M2)'].dropna().mean()
        avg_harga_all  = df_dash['Harga (Rp)'].dropna().mean()
        region_count   = df_dash['Region'].nunique()

        st.markdown('#### Overview Keseluruhan Master Lahan')
        # Row 1: count metrics
        d1, d2, d3, d4 = st.columns(4)
        d1.metric('Total Lahan', f'{total_lahan:,}')
        d2.metric('Tersedia', f'{tersedia_total:,} ({pct_tersedia}%)')
        d3.metric('Tersewa', f'{tersewa_total:,} ({100-pct_tersedia:.1f}%)')
        d4.metric('Jumlah Region', str(region_count))

        # Row 2: value metrics — pakai 2 kolom lebar + markdown biar Rupiah tidak terpotong
        harga_series = df_dash['Harga (Rp)'].dropna()
        min_harga_all = harga_series.min() if not harga_series.empty else None
        max_harga_all = harga_series.max() if not harga_series.empty else None

        def _metric_html(label, value):
            """Card metric tanpa truncation — nilai panjang tetap tampil penuh (wrap)."""
            return f"""
            <div style="padding: 0.5rem 0;">
                <p style="font-size: 0.875rem; color: rgba(49,51,63,0.6); margin: 0;">{label}</p>
                <p style="font-size: 1.6rem; font-weight: 600; margin: 0; word-wrap: break-word;
                          line-height: 1.2; overflow-wrap: anywhere;">{value}</p>
            </div>
            """

        r1a, r1b = st.columns(2)
        with r1a:
            st.markdown(_metric_html('Avg Luas',
                        f'{avg_luas_all:,.0f} m²' if avg_luas_all == avg_luas_all else '-'),
                        unsafe_allow_html=True)
        with r1b:
            st.markdown(_metric_html('Avg Harga',
                        fmt_rupiah(avg_harga_all) if avg_harga_all == avg_harga_all else '-'),
                        unsafe_allow_html=True)

        r2a, r2b = st.columns(2)
        with r2a:
            st.markdown(_metric_html('Harga Terendah',
                        fmt_rupiah(min_harga_all) if min_harga_all is not None else '-'),
                        unsafe_allow_html=True)
        with r2b:
            st.markdown(_metric_html('Harga Tertinggi',
                        fmt_rupiah(max_harga_all) if max_harga_all is not None else '-'),
                        unsafe_allow_html=True)

        st.divider()

        with st.expander('🔍 Debug: nilai unik Status Sewa di master', expanded=False):
            st.dataframe(df_dash['Status Sewa'].value_counts(dropna=False).reset_index()
                         .rename(columns={'Status Sewa': 'Nilai', 'count': 'Jumlah'}),
                         hide_index=True)

        da, db = st.columns(2)
        with da:
            st.markdown('**Status Sewa — Semua Lahan**')
            status_vc = df_dash['Status Sewa'].fillna('(tidak diketahui)').value_counts()
            st.bar_chart(status_vc)

        with db:
            st.markdown('**Lahan per Region**')
            reg_vc = df_dash.groupby('Region').agg(
                Tersedia=('Status Sewa', lambda x: (x == 'Tersedia').sum()),
                Tersewa=('Status Sewa', lambda x: (x == 'Tersewa').sum()),
            )
            st.bar_chart(reg_vc)

        st.divider()

        dc, dd = st.columns(2)
        with dc:
            st.markdown('**Top 15 Kota/Kab — Jumlah Lahan Tersedia**')
            top_kota = (df_dash[df_dash['Status Sewa'] == 'Tersedia']['Kota/Kab']
                        .fillna('(tanpa kota)').value_counts().head(15))
            st.bar_chart(top_kota)

        with dd:
            st.markdown('**Sebaran Tipe SPBU**')
            tipe_vc = df_dash['Tipe SPBU'].fillna('(lainnya)').value_counts().head(10)
            st.bar_chart(tipe_vc)

        st.divider()

        st.markdown('**Occupancy Rate per Region**')
        # Occupancy computed directly from Master FULL (Status Sewa: Tersedia/Tersewa)
        _occ_master_bytes = st.session_state.get('_dash_occ_master')
        if _occ_master_bytes:
            _df_m = occ_load_master(_occ_master_bytes)
            _status_m = _df_m['Status Sewa'].astype(str).str.strip()
            _tersedia_reg = (_df_m[_status_m == 'Tersedia'].groupby('Region').size().rename('Tersedia'))
            _tersewa_reg  = (_df_m[_status_m == 'Tersewa'].groupby('Region').size().rename('Tersewa'))
            occ_reg = pd.concat([_tersedia_reg, _tersewa_reg], axis=1).fillna(0).astype(int).reset_index()
            occ_reg['Total'] = occ_reg['Tersedia'] + occ_reg['Tersewa']
            occ_reg['Occupancy %'] = (occ_reg['Tersewa'] / occ_reg['Total'].replace(0, 1) * 100).round(1)
            occ_reg = occ_reg.sort_values('Occupancy %', ascending=False)
            st.caption('Sumber: Master Lahan FULL — Occupancy = Tersewa / (Tersedia + Tersewa)')
            st.dataframe(occ_reg[['Region', 'Tersedia', 'Tersewa', 'Total', 'Occupancy %']],
                         width='stretch', hide_index=True,
                         column_config={
                             'Tersedia':    st.column_config.NumberColumn(format='%d'),
                             'Tersewa':     st.column_config.NumberColumn(format='%d'),
                             'Total':       st.column_config.NumberColumn(format='%d'),
                             'Occupancy %': st.column_config.ProgressColumn(
                                 'Occupancy %', format='%.1f%%', min_value=0, max_value=100),
                         })
        else:
            st.info('Upload **NFR Master Lahan FULL** di **sidebar** untuk melihat occupancy rate per region.')

        st.divider()
        st.markdown('**Proven Locations (Lokasi Terbukti)**')
        st.caption('Lokasi dengan throughput tinggi yang bisa dijadikan benchmark untuk area serupa')

        # Filter proven locations: must have valid coordinates and throughput data
        df_proven = df_dash.copy()
        df_proven = df_proven[df_proven['Latitude'].notna() & df_proven['Longitude'].notna()]

        # Check which throughput column exists
        throughput_cols = ['Total2', 'PSO2', 'Non PSO2']
        available_col = None
        for col in throughput_cols:
            if col in df_proven.columns and df_proven[col].notna().any():
                available_col = col
                break

        if available_col is not None:
            df_proven = df_proven[df_proven[available_col].notna()]
            df_proven = df_proven[df_proven[available_col] > 0]

            # Calculate percentile threshold for "high performers"
            percentile_75 = df_proven[available_col].quantile(0.75)
            df_proven_high = df_proven[df_proven[available_col] >= percentile_75].copy()

            if not df_proven_high.empty:
                # Add ranking
                df_proven_high = df_proven_high.sort_values(available_col, ascending=False).head(50)
                df_proven_high['Rank'] = range(1, len(df_proven_high) + 1)

                # Display summary metrics
                p1, p2, p3, p4 = st.columns(4)
                p1.metric('Proven Locations', f'{len(df_proven_high):,}')
                p2.metric('Avg Throughput', f'{df_proven_high[available_col].mean():,.0f}')
                p3.metric('Min Threshold (P75)', f'{percentile_75:,.0f}')
                p4.metric('Top Performer', f'{df_proven_high[available_col].max():,.0f}')

                # Show table with key columns
                display_cols = ['Rank', 'No SPBU', 'Provinsi', 'Kota/Kab', 'Tipe SPBU',
                               available_col, 'Latitude', 'Longitude']
                display_cols = [c for c in display_cols if c in df_proven_high.columns]

                st.dataframe(
                    df_proven_high[display_cols],
                    width='stretch',
                    hide_index=True,
                    column_config={
                        'Rank': st.column_config.NumberColumn('Rank', format='%d'),
                        available_col: st.column_config.NumberColumn(
                            available_col,
                            format='%d',
                            help='Transaksi per bulan'
                        ),
                        'Latitude': st.column_config.NumberColumn('Lat', format='%.6f'),
                        'Longitude': st.column_config.NumberColumn('Lon', format='%.6f'),
                    }
                )

                st.caption(f'Menampilkan top 50 lokasi dengan {available_col} ≥ P75 ({percentile_75:,.0f}). '
                          'Gunakan data ini untuk analisis kompetitor atau pemilihan site baru.')
            else:
                st.info('Tidak ada lokasi dengan throughput tinggi yang memenuhi kriteria (≥ P75).')
        else:
            st.info('Data throughput (Total2/PSO2/Non PSO2) tidak tersedia. Upload Template untuk melihat Proven Locations.')

        if result is not None and not result.empty:
            st.divider()
            st.markdown(f'**Filter aktif: {len(result):,} dari {total_lahan:,} lahan** '
                        f'({len(result)/total_lahan*100:.1f}%) — lihat tab Tabel / Peta / Chart untuk detail.')

    with tabs[1]:
        if _result_empty:
            st.info('Tidak ada lokasi yang cocok dengan filter saat ini.')
        else:
            sortable = [c for c in result.columns if c not in ('Latitude', 'Longitude')]
            sc1, sc2 = st.columns([1, 1])
            with sc1:
                sort_cols = st.multiselect('Urutkan (urutan pilihan = prioritas)', sortable,
                                           default=[c for c in ['No SPBU', 'Kota/Kab'] if c in sortable],
                                           key='f_sort_cols')
            with sc2:
                sort_desc = st.multiselect('Urutkan menurun untuk kolom', sort_cols, key='f_sort_desc')
            cari = st.text_input('Cari teks di semua kolom (min 2 karakter)')

            show_cols = st.multiselect('Kolom yang ditampilkan', all_cols, key='f_show_cols')

            view = result.copy()
            if sort_cols:
                view = view.sort_values(by=sort_cols,
                                        ascending=[c not in sort_desc for c in sort_cols])
            if cari and len(cari) >= 2:
                view = view[view.astype(str).apply(
                    lambda col: col.str.contains(cari, case=False, na=False)).any(axis=1)]

            st.caption(f'Menampilkan **{len(view):,} lokasi** '
                       f'— klik header kolom untuk sort, ikon corong untuk filter per kolom')

            col_cfg = {}
            for c in show_cols:
                cl = c.lower()
                if 'harga' in cl or ('sewa' in cl and 'rp' in cl):
                    col_cfg[c] = st.column_config.NumberColumn(c, format='Rp %,.0f')
                elif 'luas' in cl and 'm2' in cl.replace(' ', '').replace('(', '').replace(')', ''):
                    col_cfg[c] = st.column_config.NumberColumn(c, format='%.0f m²')

            valid_show = [c for c in show_cols if c in view.columns]
            st.dataframe(view[valid_show], width='stretch', height=420,
                         column_config=col_cfg if col_cfg else None)

    with tabs[2]:
        if _result_empty:
            st.info('Tidak ada lokasi yang cocok dengan filter saat ini.')
        else:
            geo = result.dropna(subset=['Latitude', 'Longitude'])
            geo = geo[geo['Latitude'].between(-11, 6) & geo['Longitude'].between(95, 141)]
            if geo.empty:
                st.info('Tidak ada data koordinat (upload Template untuk peta).')
            else:
                total_markers = len(geo)
                geo = _stratified_sample(geo)              # cap ke MAP_MARKER_CAP bila perlu
                shown_markers = len(geo)

                # Compact data: [lat, lon, kode, spbu, provinsi, kota, luas, harga, status]
                # Popup dibangun di JavaScript (bukan HTML per baris di Python) —
                # payload ~10x lebih kecil (7MB -> 0.7MB), build ~10x lebih cepat.
                def _col(name, default=None):
                    return geo[name].tolist() if name in geo.columns else [default] * len(geo)

                rows = list(zip(
                    geo['Latitude'].round(5).tolist(),
                    geo['Longitude'].round(5).tolist(),
                    _col('Kode Lahan'),
                    _col('No SPBU'),
                    _col('Provinsi'),
                    _col('Kota/Kab'),
                    geo['Luas (M2)'].round(1).tolist() if 'Luas (M2)' in geo.columns else [None] * len(geo),
                    geo['Harga (Rp)'].round(0).tolist() if 'Harga (Rp)' in geo.columns else [None] * len(geo),
                    _col('Status Sewa'),
                ))

                callback = """\
function (row) {
    var color = row[8] === 'Tersedia' ? '#5A8A00' : (row[8] === 'Tersewa' ? '#E31B23' : '#888');
    var fmt = function(v) { return 'Rp ' + Number(v).toLocaleString('id-ID'); };
    var esc = function(s) { var d = document.createElement('div'); d.textContent = s == null ? '-' : String(s); return d.innerHTML; };
    var html = '<div style="font-family:sans-serif;font-size:13px;min-width:240px;">'
        + '<div style="font-weight:700;font-size:14px;margin-bottom:4px;">' + esc(row[2]) + '</div>'
        + '<div style="color:#555;margin-bottom:6px;">No SPBU: <b>' + esc(row[3]) + '</b></div>'
        + '<div style="margin-bottom:2px;">' + esc(row[4]) + ' &mdash; ' + esc(row[5]) + '</div>'
        + '<div style="margin-bottom:4px;">Luas: <b>' + esc(row[6]) + ' m²</b> &nbsp; Harga: <b>' + fmt(row[7]) + '</b></div>'
        + '<span style="background:' + color + ';color:#fff;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">' + esc(row[8]) + '</span>'
        + '</div>';
    var marker = L.marker([row[0], row[1]]);
    marker.bindPopup(html, {maxWidth: 340});
    return marker;
}
"""
                m = folium.Map(location=[geo['Latitude'].mean(), geo['Longitude'].mean()],
                               zoom_start=5, tiles='OpenStreetMap')
                m.add_child(FastMarkerCluster(rows, callback=callback))

                # Static render — tanpa round-trip websocket st_folium.
                # Peta display-only: interaksi (pan/zoom/cluster/popup) jalan di browser,
                # TIDAK memicu rerun Streamlit (sumber hang sebelumnya).
                components.html(m._repr_html_(), height=520, scrolling=False)

                if shown_markers < total_markers:
                    st.caption(f'🗺️ Menampilkan **{shown_markers:,} dari {total_markers:,}** lokasi '
                               f'(sampel proporsional per status × region). '
                               f'Persempit filter (region/provinsi) untuk detail penuh.')

    with tabs[3]:
        if _result_empty:
            st.info('Tidak ada lokasi yang cocok dengan filter saat ini.')
        else:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown('**Jumlah per Kota**')
                vc = result['Kota/Kab'].fillna('(tanpa kota)').value_counts().head(15)
                st.bar_chart(vc)
            with c2:
                st.markdown('**Sebaran Luas (M2)**')
                bins = pd.cut(result['Luas (M2)'].dropna(), bins=[0, 20, 40, 60, 100, 200, 10000],
                              labels=['<20', '20-40', '40-60', '60-100', '100-200', '200+'])
                st.bar_chart(bins.value_counts())
            with c3:
                st.markdown('**Harga vs Luas**')
                s = result.dropna(subset=['Luas (M2)', 'Harga (Rp)'])
                st.scatter_chart(s[['Luas (M2)', 'Harga (Rp)']], x='Luas (M2)', y='Harga (Rp)')

    with tabs[4]:
        st.markdown('**Ringkasan per Region & Kota/Kab**')
        if _result_empty:
            st.info('Tidak ada data untuk diringkas.')
        else:
            DIGIT_REGION = {v: k for k, v in REGION_DIGIT.items()}
            tmp = result.copy()
            tmp['Region'] = tmp['No SPBU'].astype(str).str[0].map(DIGIT_REGION).fillna('(lainnya)')

            reg_sum = tmp.groupby('Region').agg(
                Lokasi=('Kode Lahan', 'count'),
                Tersedia=('Status Sewa', lambda x: (x == 'Tersedia').sum()),
                Tersewa=('Status Sewa', lambda x: (x == 'Tersewa').sum()),
                Rata_Luas=('Luas (M2)', 'mean'),
                Rata_Harga=('Harga (Rp)', 'mean'),
            ).reset_index()
            reg_sum['Rata_Luas'] = reg_sum['Rata_Luas'].round(0)
            reg_sum['Rata_Harga'] = reg_sum['Rata_Harga'].round(0)
            reg_sum.columns = ['Region', 'Total Lokasi', 'Tersedia', 'Tersewa',
                               'Rata-rata Luas (m²)', 'Rata-rata Harga (Rp)']
            st.dataframe(reg_sum, width='stretch', hide_index=True,
                         column_config={
                             'Rata-rata Luas (m²)': st.column_config.NumberColumn(format='%.0f m²'),
                             'Rata-rata Harga (Rp)': st.column_config.NumberColumn(format='Rp %,.0f'),
                         })

            st.markdown('**Ringkasan per Kota/Kab**')
            kota_sum_full = tmp.groupby('Kota/Kab').agg(
                Lokasi=('Kode Lahan', 'count'),
                Tersedia=('Status Sewa', lambda x: (x == 'Tersedia').sum()),
                Tersewa=('Status Sewa', lambda x: (x == 'Tersewa').sum()),
                Rata_Luas=('Luas (M2)', 'mean'),
                Rata_Harga=('Harga (Rp)', 'mean'),
            ).reset_index().sort_values('Lokasi', ascending=False)
            kota_sum_full['Rata_Luas'] = kota_sum_full['Rata_Luas'].round(0)
            kota_sum_full['Rata_Harga'] = kota_sum_full['Rata_Harga'].round(0)
            kota_sum_full.columns = ['Kota/Kab', 'Total Lokasi', 'Tersedia', 'Tersewa',
                                     'Rata-rata Luas (m²)', 'Rata-rata Harga (Rp)']
            st.dataframe(kota_sum_full, width='stretch', hide_index=True,
                         column_config={
                             'Rata-rata Luas (m²)': st.column_config.NumberColumn(format='%.0f m²'),
                             'Rata-rata Harga (Rp)': st.column_config.NumberColumn(format='Rp %,.0f'),
                         })

            st.markdown('**Ringkasan per SPBU**')
            spbu_sum = tmp.groupby('No SPBU').agg(
                Region=('Region', 'first'),
                Provinsi=('Provinsi', 'first'),
                Kota=('Kota/Kab', 'first'),
                Total=('Kode Lahan', 'count'),
                Rata_Luas=('Luas (M2)', 'mean'),
                Rata_Harga=('Harga (Rp)', 'mean'),
            ).reset_index()
            spbu_sum['Rata_Luas']  = spbu_sum['Rata_Luas'].round(0)
            spbu_sum['Rata_Harga'] = spbu_sum['Rata_Harga'].round(0)
            spbu_sum = spbu_sum.sort_values('Total', ascending=False)
            spbu_sum.columns = ['No SPBU', 'Region', 'Provinsi', 'Kota/Kab',
                                'Total Lahan', 'Rata-rata Luas (m²)', 'Rata-rata Harga (Rp)']

            # Enrich with Tersewa count from Master FULL (Status Sewa) if uploaded in tab Occupancy
            _occ_m = st.session_state.get('_dash_occ_master')
            spbu_col_cfg = {
                'Total Lahan':          st.column_config.NumberColumn(format='%d'),
                'Rata-rata Luas (m²)':  st.column_config.NumberColumn(format='%.0f m²'),
                'Rata-rata Harga (Rp)': st.column_config.NumberColumn(format='Rp %,.0f'),
            }
            if _occ_m:
                _df_m2 = occ_load_master(_occ_m)
                _status_m2 = _df_m2['Status Sewa'].astype(str).str.strip()
                tersewa_per_spbu = (_df_m2[_status_m2 == 'Tersewa']
                                    .groupby('No SPBU').size()
                                    .rename('Tersewa').reset_index())
                tersewa_per_spbu['No SPBU'] = tersewa_per_spbu['No SPBU'].astype(str).str.strip()
                spbu_sum['No SPBU'] = spbu_sum['No SPBU'].astype(str).str.strip()
                spbu_sum = spbu_sum.merge(tersewa_per_spbu, on='No SPBU', how='left')
                spbu_sum['Tersewa'] = spbu_sum['Tersewa'].fillna(0).astype(int)
                spbu_sum['Occupancy %'] = (
                    spbu_sum['Tersewa'] / spbu_sum['Total Lahan'].replace(0, 1) * 100
                ).round(1)
                spbu_col_cfg['Tersewa']  = st.column_config.NumberColumn(format='%d')
                spbu_col_cfg['Occupancy %'] = st.column_config.ProgressColumn(
                    'Occupancy %', format='%.1f%%', min_value=0, max_value=100)
                st.caption('Occupancy % dari Master Lahan FULL — Tersewa / Total Lahan per SPBU.')
            else:
                st.caption('Upload **NFR Master Lahan FULL** di **sidebar** untuk melihat Occupancy % per SPBU.')

            st.dataframe(spbu_sum, width='stretch', hide_index=True, column_config=spbu_col_cfg)
            # store for export
            st.session_state['_spbu_sum'] = spbu_sum
            st.session_state['_reg_sum']  = reg_sum
            st.session_state['_kota_sum'] = kota_sum_full

    # ----- export -----
    if not _result_empty:
        st.divider()
        cA, cB = st.columns([1, 3])
        with cA:
            export_fname = build_export_filename(region, statuses, provinsi, kota)
            _raw_cols = tuple(c for c in show_cols if c in result.columns)
            export_cols = _raw_cols or tuple(c for c in DEFAULT_COLS if c in result.columns)
            n_export = len(result)

            st.markdown('**Export Excel (.xlsx)**')
            do_export = st.button(f'⬇️ Export Excel ({n_export:,} baris)',
                                  width='stretch', key='_export_btn',
                                  help='±3-5 detik utk ribuan baris.')
            if do_export:
                if n_export > 6000:
                    # Langsung DataFrame -> fast xlsxwriter. TANPA to_excel_bytes /
                    # to_dict / cache-data agar tak ada overhead tambahan di app.
                    # summary_df=df (FULL master) → Ringkasan tetap utuh walau Data terfilter.
                    with st.spinner('Menyiapkan Excel... ±3-5 dtk'):
                        excel_data = build_proposal_workbook_fast(
                            result, list(export_cols), summary_df=df)
                else:
                    with st.spinner('Menyiapkan Excel...'):
                        reg_rec  = st.session_state.get('_reg_sum',  pd.DataFrame()).to_dict('records')
                        kota_rec = st.session_state.get('_kota_sum', pd.DataFrame()).to_dict('records')
                        spbu_rec = st.session_state.get('_spbu_sum', pd.DataFrame()).to_dict('records')
                        excel_data = to_excel_bytes(
                            result.to_dict('records'), export_cols,
                            reg_records=reg_rec or None,
                            kota_records=kota_rec or None,
                            spbu_records=spbu_rec or None,
                        )
                st.session_state['_excel_data'] = excel_data
                st.session_state['_excel_fname'] = export_fname
            if st.session_state.get('_excel_data') is not None:
                st.download_button('⬇️ Download Excel',
                                   data=st.session_state['_excel_data'],
                                   file_name=st.session_state.get('_excel_fname', export_fname),
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   width='stretch')
        with cB:
            st.caption('**Excel** = ±3-5 dtk utk 31k baris (fast xlsxwriter + Charts). '
                       'Untuk ≤6000 baris: tetap ada Charts + Peta + Ringkasan.')


# ----- tab Space Check -----
with tabs[5]:
    try:
        st.markdown("""
<div style="margin-bottom:0.5rem;">
    <div style="font-size:1.1rem;font-weight:700;color:#1A1A1A;">Space Check — Deteksi Konflik Ketersediaan Lahan</div>
    <div style="font-size:0.82rem;color:#666;margin-top:2px;">
        Cross-reference lahan <b>Tersedia</b> di Master Lahan vs kontrak aktif di Beroperasi untuk
        deteksi data tidak sinkron. Cegah tawarkan lahan yang sebenarnya sudah terisi ke prospect.
    </div>
</div>
""", unsafe_allow_html=True)

        show_metodologi('space_check')

        # Check for required files in session state (uploaded in Occupancy tab)
        _sc_master_bytes = st.session_state.get('_dash_occ_master')
        _sc_beroperasi_bytes = st.session_state.get('_dash_occ_beroperasi')

        if not _sc_master_bytes or not _sc_beroperasi_bytes:
            st.warning('⚠️ **File yang dibutuhkan belum diupload**')
            st.info('Upload **NFR Master Lahan FULL** di **sidebar** dan '
                   '**Pengajuan Beroperasi** di tab **Occupancy** untuk menggunakan Space Check.')
            st.markdown("""
**Cara pakai:**
1. Upload **NFR Master Lahan FULL** (xlsx) di **sidebar**
2. Buka tab **Occupancy**, upload **Pengajuan Beroperasi** (xlsx)
3. Kembali ke tab ini untuk mulai checking
""")
        else:
            # Search input
            st.markdown('### 🔍 Cari Lahan Available')

            col1, col2 = st.columns([3, 1])
            with col1:
                keyword = st.text_input(
                    'Keyword Jenis Usaha',
                    placeholder='Contoh: nitrogen, coffee, atm, alfamart, indomaret',
                    key='_space_check_keyword',
                    help='Cari di kolom: Nama Tenant, Nama Brand, Kategori, Sub Kategori (case-insensitive)'
                )
            with col2:
                st.markdown('<div style="height:28px"></div>', unsafe_allow_html=True)  # Spacer
                check_btn = st.button('🔍 Check Availability', type='primary', width='stretch')

            if check_btn:
                if not keyword or not keyword.strip():
                    st.error('❌ Masukkan keyword terlebih dahulu')
                else:
                    with st.spinner(f'🔄 Menganalisis data untuk keyword "{keyword.strip()}"...'):
                        try:
                            results = detect_space_conflicts(
                                _sc_master_bytes,
                                _sc_beroperasi_bytes,
                                keyword.strip()
                            )
                            st.session_state['_space_check_results'] = results
                            st.success(f'✅ Analisis selesai untuk "{keyword.strip()}"')
                        except Exception as e:
                            st.error(f'Error saat analisis: {e}')
                            st.session_state['_space_check_results'] = None

            # Display results
            if st.session_state.get('_space_check_results'):
                results = st.session_state['_space_check_results']
                stats = results['stats']

                st.divider()

                # Summary metrics
                st.markdown('### 📊 Hasil Analisis')
                m1, m2, m3, m4 = st.columns(4)
                m1.metric('Total Lahan Tersedia', f"{stats['total_tersedia']:,}")
                m2.metric('🟢 Clean Available', f"{stats['clean_count']:,}")
                m3.metric('🔴 Conflicts Detected', f"{stats['conflict_count']:,}")
                m4.metric('Clean %', f"{stats['clean_pct']:.1f}%")

                if stats['expired_contracts'] > 0:
                    st.warning(f"⚠️ {stats['expired_contracts']} kontrak sudah expired tapi masih di Beroperasi — data quality issue")

                st.divider()

                # Results display
                if stats['conflict_count'] == 0:
                    st.success('✅ **Semua lahan Tersedia adalah CLEAN!** Tidak ada konflik data terdeteksi.')
                    st.info(f'Semua {stats["clean_count"]} lahan aman ditawarkan ke prospect dengan keyword "{stats["keyword"]}".')
                elif stats['clean_count'] == 0:
                    st.error(f'❌ **Semua lahan Tersedia memiliki konflik!** {stats["conflict_count"]} lahan tidak bisa ditawarkan.')
                    st.warning('Data Master Lahan perlu di-update untuk semua lahan ini.')

                # Tabbed results
                rtab1, rtab2 = st.tabs([
                    f'✅ Clean Available ({stats["clean_count"]})',
                    f'⚠️ Conflicts Detected ({stats["conflict_count"]})'
                ])

                with rtab1:
                    clean = results['clean']
                    if not clean.empty:
                        st.markdown(f'**{len(clean)} lahan aman ditawarkan ke prospect** — tidak ada kontrak aktif matching keyword.')

                        # Display columns
                        display_cols = ['No SPBU', 'Kode Lahan', 'Provinsi', 'Kota/Kab',
                                       'Luas (M2)', 'Harga (Rp)', 'Region']
                        display_cols = [c for c in display_cols if c in clean.columns]

                        st.dataframe(
                            clean[display_cols],
                            width='stretch',
                            hide_index=True,
                            column_config={
                                'Luas (M2)': st.column_config.NumberColumn(format='%.0f m²'),
                                'Harga (Rp)': st.column_config.NumberColumn(format='Rp %,.0f'),
                            }
                        )

                        st.markdown('**Export Space Check Results**')
                        if st.button('📥 Export All Results (Excel)', key='export_space_check'):
                            with st.spinner('Menyiapkan Excel...'):
                                clean_records = clean.to_dict('records')
                                conflicts_records = results['conflicts'].to_dict('records')
                                excel_data = space_check_to_excel(clean_records, conflicts_records, stats['keyword'])
                                fname = f'space_check_{stats["keyword"].replace(" ", "_")}_{datetime.date.today()}.xlsx'
                                st.download_button(
                                    '⬇️ Download Space Check Results',
                                    data=excel_data,
                                    file_name=fname,
                                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                    else:
                        st.info('Tidak ada lahan clean available.')

                with rtab2:
                    conflicts = results['conflicts']
                    if not conflicts.empty:
                        st.markdown(f'**⚠️ {len(conflicts)} lahan show konflik data** — Master says Tersedia tapi punya kontrak aktif.')
                        st.caption('Lahan ini TIDAK BOLEH ditawarkan ke prospect sampai data di-sync.')

                        # Display with Beroperasi details
                        display_cols = ['No SPBU', 'Kode Lahan', 'Provinsi', 'Kota/Kab',
                                       'Luas (M2)', 'Harga (Rp)', 'Region',
                                       'Contract_Count', 'Brands', 'Kategori', 'Latest_End_Date', 'Is_Expired']
                        display_cols = [c for c in display_cols if c in conflicts.columns]

                        st.dataframe(
                            conflicts[display_cols],
                            width='stretch',
                            hide_index=True,
                            column_config={
                                'Luas (M2)': st.column_config.NumberColumn(format='%.0f m²'),
                                'Harga (Rp)': st.column_config.NumberColumn(format='Rp %,.0f'),
                                'Contract_Count': st.column_config.NumberColumn('Contracts', format='%d'),
                                'Latest_End_Date': st.column_config.DateColumn('Contract End', format='YYYY-MM-DD'),
                                'Is_Expired': st.column_config.CheckboxColumn('Expired?'),
                            }
                        )

                    else:
                        st.success('✅ Tidak ada konflik terdeteksi!')

                st.divider()

                # Beroperasi matches info
                beroperasi_matches = results['beroperasi_matches']
                with st.expander(f'📋 Detail Kontrak Beroperasi Matching "{stats["keyword"]}" ({len(beroperasi_matches)} kontrak)'):
                    if not beroperasi_matches.empty:
                        show_cols = ['No SPBU', 'Nama Tenant', 'Nama Brand', 'Kategori', 'Sub Kategori',
                                    'Tgl Mulai Sewa Lahan', 'Tgl Akhir Sewa Lahan', 'Region']
                        show_cols = [c for c in show_cols if c in beroperasi_matches.columns]
                        st.dataframe(beroperasi_matches[show_cols], width='stretch', hide_index=True)

    except Exception as _sc_err:
        st.error(f'Error di tab Space Check: {_sc_err}')
        import traceback
        st.code(traceback.format_exc(), language='python')


# ----- tab Occupancy (always rendered — independent of master upload) -----
with tabs[6]:
    try:
        st.markdown("""
<div style="margin-bottom:0.5rem;">
    <div style="font-size:1.1rem;font-weight:700;color:#1A1A1A;">Occupancy — Perbandingan Lahan Tersedia vs Tersewa</div>
    <div style="font-size:0.82rem;color:#666;margin-top:2px;">
        Master Lahan NFR diambil otomatis dari <b>sidebar</b>. Occupancy dihitung langsung dari Status Sewa (Tersedia/Tersewa).
        Tambahkan <b>Telah Berakhir</b> untuk klasifikasi SPBU Baru vs Kembali Tersedia. <b>Beroperasi</b> opsional (enrichment tenant detail).
    </div>
</div>
""", unsafe_allow_html=True)

        show_metodologi('occupancy')

        # Master Lahan sekarang di-upload sekali di sidebar (dipakai semua tab).
        # Tab ini hanya butuh Berakhir (opsional) + Beroperasi (opsional).
        st.caption('Master Lahan NFR diambil otomatis dari **sidebar**. '
                   'Upload tambahan hanya untuk klasifikasi & enrichment.')

        oc1, oc2 = st.columns(2)
        occ_berakhir_up   = oc1.file_uploader('Pengajuan — Telah Berakhir (xlsx) · opsional',
                                               type=['xlsx'], key='occ_berakhir_up',
                                               help='Export Pengajuan Izin Prinsip filter Status = Telah Berakhir. '
                                                    'Opsional — aktifkan klasifikasi SPBU Baru vs Kembali Tersedia.')
        occ_beroperasi_up = oc2.file_uploader('Pengajuan — Beroperasi (xlsx) · opsional',
                                               type=['xlsx'], key='occ_beroperasi_up',
                                               help='Export Pengajuan Izin Prinsip filter Status = Beroperasi. '
                                                    'Opsional — untuk breakdown kategori & top SPBU tenant.')

        _occ_m_bytes = st.session_state.get('_dash_occ_master')
        if not _occ_m_bytes:
            st.info('Upload **NFR Master Lahan FULL** (berisi Tersedia + Tersewa) di **sidebar** untuk memulai analisis occupancy.')
        else:
            # Beroperasi shared ke Space Check — simpan di session state + rerun bila berubah
            _occ_b_bytes = to_bytes(occ_beroperasi_up) if occ_beroperasi_up else None
            _prev_b = st.session_state.get('_dash_occ_beroperasi')
            st.session_state['_dash_occ_beroperasi'] = _occ_b_bytes
            if _prev_b != _occ_b_bytes:
                st.rerun()

            df_m = occ_load_master(_occ_m_bytes)
            berakhir_uploaded = occ_berakhir_up is not None
            beroperasi_uploaded = occ_beroperasi_up is not None
            df_e = occ_load_berakhir(to_bytes(occ_berakhir_up)) if berakhir_uploaded else None
            df_b = occ_load_pengajuan(to_bytes(occ_beroperasi_up)) if beroperasi_uploaded else None

            # --- guard: master terfilter (Tersedia-only) → occupancy tidak akurat ---
            if 'Status Sewa' in df_m.columns:
                _ss_raw = df_m['Status Sewa'].astype(str).str.strip()
                _tersedia_n = int((_ss_raw == 'Tersedia').sum())
                _tersewa_n  = int((_ss_raw == 'Tersewa').sum())
                if _tersedia_n > 0 and _tersewa_n == 0:
                    st.warning('⚠️ Master yang di-upload hanya berisi status **Tersedia** (terfilter). '
                               'Occupancy akan terbaca **0%**. Upload export **FULL** (Tersedia + Tersewa) '
                               'di **sidebar** untuk occupancy yang akurat.')

            # --- handle stale Tersewa (Tgl Berakhir already past but still marked Tersewa) ---
            _tgl_akhir_col = next((c for c in ['Tanggal Berakhir Sewa', 'Tgl Akhir Sewa Lahan',
                                               'Tgl Berakhir Sewa'] if c in df_m.columns), None)
            today = datetime.date.today()
            stale_mask = pd.Series(False, index=df_m.index)
            if _tgl_akhir_col:
                _dates = pd.to_datetime(df_m[_tgl_akhir_col], errors='coerce')
                stale_mask = (_dates.dt.date < today) & (df_m['Status Sewa'].astype(str).str.strip() == 'Tersewa')
            stale_count = int(stale_mask.sum())

            view_mode = st.radio(
                'Mode tampilan occupancy',
                ['Aktual (sesuai Status Sewa)', 'Disesuaikan (exclude Tersewaexpired)'],
                horizontal=True,
                help='Sesuaikan: lahan berstatus Tersewa tapi Tgl Berakhir sudah lewat dianggap Tersedia '
                     '(data belum diupdate di master).',
                key='occ_view_mode',
            ) if stale_count > 0 else None
            adjusted = view_mode is not None and view_mode.startswith('Disesuaikan')

            # Working copy of Status Sewa (adjusted excludes stale Tersewa)
            if 'Status Sewa' not in df_m.columns:
                st.error('Kolom "Status Sewa" tidak ditemukan di Master.')
                st.stop()
            _status_eff = df_m['Status Sewa'].astype(str).str.strip().copy()
            if adjusted:
                _status_eff = _status_eff.mask(stale_mask, 'Tersedia')

            # --- summary per region (computed from Master FULL alone) ---
            if 'Region' not in df_m.columns:
                st.error('Kolom "Region" tidak ditemukan di Master.')
                st.stop()
            _work = df_m.copy()
            _work['_eff_status'] = _status_eff
            _tersedia_mask = _work['_eff_status'] == 'Tersedia'
            _tersewa_mask  = _work['_eff_status'] == 'Tersewa'
            tersedia_s = (_work[_tersedia_mask].groupby('Region').size().rename('Tersedia'))
            tersewa_s  = (_work[_tersewa_mask].groupby('Region').size().rename('Tersewa'))
            frames = [tersedia_s, tersewa_s]
            if berakhir_uploaded:
                berakhir_s = df_e.groupby('Region').size().rename('Telah Berakhir')
                frames.append(berakhir_s)
            if beroperasi_uploaded:
                beroperasi_s = df_b.groupby('Region').size().rename('Beroperasi')
                frames.append(beroperasi_s)
            summary = pd.concat(frames, axis=1).fillna(0).astype(int)
            # Occupancy from Master FULL: Tersewa / (Tersedia + Tersewa)
            summary['Total'] = summary['Tersedia'] + summary['Tersewa']
            summary['% Occupancy'] = (summary['Tersewa'] / summary['Total'].replace(0, 1) * 100).round(1)
            summary = summary.sort_values('% Occupancy', ascending=False).reset_index()

            # --- metric cards ---
            tot_tersedia = int(summary['Tersedia'].sum())
            tot_tersewa  = int(summary['Tersewa'].sum())
            tot_total    = int(summary['Total'].sum())
            overall_occ  = round(tot_tersewa / tot_total * 100, 1) if tot_total else 0.0
            # Lahan pending (Status Sewa = '-') excluded dari occupancy tapi dilaporkan
            tot_pending  = int((df_m['Status Sewa'].astype(str).str.strip() == '-').sum())
            highest      = summary.iloc[0] if not summary.empty else None

            # Row 1: Core metrics
            mcs1 = st.columns(4)
            mcs1[0].metric('Total Lahan', f'{tot_total:,}')
            mcs1[1].metric('Tersedia', f'{tot_tersedia:,}')
            mcs1[2].metric('Tersewa', f'{tot_tersewa:,}')
            mcs1[3].metric('Overall Occupancy', f'{overall_occ}%')

            # Row 2: Additional metrics
            cols_n2 = (1 if (adjusted and stale_count>0) else 0) + (1 if tot_pending>0 else 0) + (1 if berakhir_uploaded else 0) + (1 if beroperasi_uploaded else 0) + 1
            if cols_n2 > 1:
                mcs2 = st.columns(cols_n2)
                _ci = 0
                if adjusted and stale_count > 0:
                    mcs2[_ci].metric('Tersewa Stale (dikecualikan)', f'{stale_count:,}',
                                    help='Lahan berstatus Tersewa tapi Tgl Berakhir sudah lewat — dianggap Tersedia dalam mode ini.')
                    _ci += 1
                if tot_pending > 0:
                    mcs2[_ci].metric('Pending ("-")', f'{tot_pending:,}',
                                    help='Lahan dengan Status Sewa "-" — Kode Lahan belum ada, masih menunggu persetujuan. '
                                         'Tidak masuk perhitungan occupancy.')
                    _ci += 1
                if berakhir_uploaded:
                    mcs2[_ci].metric('Telah Berakhir', f'{int(summary["Telah Berakhir"].sum()):,}'); _ci += 1
                if beroperasi_uploaded:
                    mcs2[_ci].metric('Beroperasi', f'{int(summary["Beroperasi"].sum()):,}'); _ci += 1
                # Use markdown for Region Tertinggi to allow wrapping
                region_value = f"{highest['Region']} ({highest['% Occupancy']}%)" if highest is not None else '-'
                mcs2[_ci].markdown(f"""
                    <div style="padding: 0.5rem 0;">
                        <p style="font-size: 0.875rem; color: rgba(49, 51, 63, 0.6); margin: 0;">Region Tertinggi</p>
                        <p style="font-size: 1.875rem; font-weight: 600; margin: 0; word-wrap: break-word; line-height: 1.2;">{region_value}</p>
                    </div>
                """, unsafe_allow_html=True)
            else:
                # Use markdown for Region Tertinggi to allow wrapping
                region_value = f"{highest['Region']} ({highest['% Occupancy']}%)" if highest is not None else '-'
                st.markdown(f"""
                    <div style="padding: 0.5rem 0;">
                        <p style="font-size: 0.875rem; color: rgba(49, 51, 63, 0.6); margin: 0;">Region Tertinggi</p>
                        <p style="font-size: 1.875rem; font-weight: 600; margin: 0; word-wrap: break-word; line-height: 1.2;">{region_value}</p>
                    </div>
                """, unsafe_allow_html=True)

            if stale_count > 0 and not adjusted:
                st.warning(f'⚠️ {stale_count:,} lahan berstatus Tersewa tapi Tgl Berakhir sudah lewat (data belum diupdate). '
                           f'Occupancy aktual mungkin lebih rendah. Aktifkan mode "Disesuaikan" untuk melihat.')

            st.divider()

            # --- A: tabel per region ---
            st.markdown('**Per Region**')
            col_cfg_reg = {
                'Tersedia':    st.column_config.NumberColumn('Tersedia',   format='%d'),
                'Tersewa':     st.column_config.NumberColumn('Tersewa',     format='%d'),
                'Total':       st.column_config.NumberColumn('Total',      format='%d'),
                '% Occupancy': st.column_config.ProgressColumn(
                    '% Occupancy', format='%.1f%%', min_value=0, max_value=100),
            }
            if berakhir_uploaded:
                col_cfg_reg['Telah Berakhir'] = st.column_config.NumberColumn('Telah Berakhir', format='%d')
            if beroperasi_uploaded:
                col_cfg_reg['Beroperasi'] = st.column_config.NumberColumn('Beroperasi', format='%d')
            _summary_disp = summary.copy()
            _summary_disp['% Occupancy'] = _summary_disp['% Occupancy'].astype(float)
            st.dataframe(_summary_disp, width='stretch', hide_index=True, column_config=col_cfg_reg)

            # --- B: stacked bar chart ---
            st.markdown('**Komposisi per Region**')
            chart_cols = ['Tersedia', 'Tersewa']
            if berakhir_uploaded:
                chart_cols.append('Telah Berakhir')
            if beroperasi_uploaded:
                chart_cols.append('Beroperasi')
            st.bar_chart(summary.set_index('Region')[chart_cols])

            st.divider()

            # --- C: SPBU classification: Baru NFR vs Kembali Tersedia vs Aktif ---
            # Master FULL knows who has active tenants (Tersewa). Berakhir = had tenant before.
            # Available SPBU = SPBU with >=1 Tersedia lahan (effective status)
            _avail_df = _work[_tersedia_mask]
            available_spbu = set(_avail_df['No SPBU'].astype(str).str.strip())
            tersewa_spbu   = set(_work[_tersewa_mask]['No SPBU'].astype(str).str.strip())  # has active tenant

            _e_spbu_col = next((c for c in ['No. SPBU', 'No SPBU', 'SPBU'] if c in df_e.columns), None) if berakhir_uploaded else None
            e_spbu = set(df_e[_e_spbu_col].astype(str).str.strip()) if (_e_spbu_col and berakhir_uploaded) else set()

            # SPBU with available space but NO active tenant on any space
            master_only = available_spbu - tersewa_spbu
            baru_spbu    = master_only - e_spbu    # never had NFR tenant
            kembali_spbu = master_only & e_spbu    # had tenant, now free
            aktif_spbu   = available_spbu & tersewa_spbu  # has available space AND active tenant (excess)

            st.markdown('**Klasifikasi SPBU**')
            st.caption('SPBU yang memiliki lahan Tersedia, diklasifikasikan berdasarkan status tenant-nya. '
                       'Master FULL sudah mengandung info Tersewa, sehingga Beroperasi tidak wajib. '
                       'Klasifikasi Baru/Kembali Tersedia memerlukan file Telah Berakhir.')

            cc1, cc2, cc3, cc4 = st.columns(4)
            cc1.metric('SPBU Punya Lahan Tersedia', len(available_spbu))
            cc2.metric('Aktif (punya Tersedia + Tersewa)',
                       len(aktif_spbu),
                       help='SPBU yang punya lahan tersedia sekaligus lahan tersewa (excess capacity).')
            cc3.metric('Baru NFR',
                       len(baru_spbu),
                       help='Punya lahan tersedia, tidak ada lahan tersewa, dan tidak ada di Telah Berakhir '
                            '→ belum pernah punya tenant NFR.')
            cc4.metric('Kembali Tersedia',
                       len(kembali_spbu),
                       help='Punya lahan tersedia, tidak ada lahan tersewa, TAPI ada di Telah Berakhir '
                            '→ pernah punya tenant, kontrak sudah berakhir.')

            digit_map = {str(i + 1): r for i, r in enumerate(REGIONS)}
            # Build classification table from available SPBU
            sub = _avail_df[_avail_df['No SPBU'].astype(str).str.strip().isin(master_only)].copy()
            sub['_s'] = sub['No SPBU'].astype(str).str.strip()
            sub['Region'] = sub['_s'].str[0].map(digit_map).fillna('Lainnya')
            sub['Status NFR'] = sub['_s'].apply(
                lambda x: 'Kembali Tersedia' if x in kembali_spbu else 'Baru NFR'
            )
            spbu_table = (sub.groupby(['No SPBU', 'Region', 'Status NFR'])
                             .size().rename('Jumlah Lahan')
                             .reset_index())

            st.dataframe(spbu_table, width='stretch', hide_index=True)

            region_class = (spbu_table.groupby(['Region', 'Status NFR'])['Jumlah Lahan']
                                      .sum().unstack(fill_value=0))
            st.bar_chart(region_class)

            st.divider()

            # --- D: breakdown per kategori & top SPBU (only when Beroperasi uploaded) ---
            kat = pd.DataFrame(columns=['Kategori', 'Jumlah'])
            top_spbu = pd.DataFrame()
            if beroperasi_uploaded and df_b is not None:
                st.markdown('**Breakdown per Kategori (Beroperasi)**')
                _kat_col = next((c for c in ['Kategori', 'Kategori NFR', 'Category'] if c in df_b.columns), None)
                if _kat_col:
                    kat = (df_b[_kat_col].value_counts()
                           .reset_index()
                           .rename(columns={_kat_col: 'Kategori', 'count': 'Jumlah'}))
                else:
                    st.caption('Kolom Kategori tidak ditemukan di file Beroperasi.')
                kc1, kc2 = st.columns([1, 2])
                with kc1:
                    st.dataframe(kat, width='stretch', hide_index=True)
                with kc2:
                    st.bar_chart(kat.set_index('Kategori')['Jumlah'])

                st.divider()
                st.markdown('**Top 20 SPBU Paling Banyak Tenant (Beroperasi)**')
                _spbu_col  = next((c for c in ['No. SPBU', 'No SPBU', 'SPBU'] if c in df_b.columns), None)
                _brand_col = next((c for c in ['Nama Brand', 'Brand', 'Nama Tenant'] if c in df_b.columns), None)
                _kota_col  = next((c for c in ['Kab./Kota', 'Kota/Kab', 'Kota'] if c in df_b.columns), None)
                if _spbu_col and _brand_col:
                    _agg = {'Jumlah_Tenant': (_brand_col, 'count'), 'Region': ('Region', 'first')}
                    if 'Provinsi' in df_b.columns: _agg['Provinsi'] = ('Provinsi', 'first')
                    if _kota_col: _agg['Kota'] = (_kota_col, 'first')
                    top_spbu = (df_b.groupby(_spbu_col)
                                    .agg(**_agg)
                                    .sort_values('Jumlah_Tenant', ascending=False)
                                    .head(20)
                                    .reset_index()
                                    .rename(columns={_spbu_col: 'No SPBU', 'Jumlah_Tenant': 'Jumlah Tenant'}))
                    st.dataframe(top_spbu, width='stretch', hide_index=True)
                else:
                    st.caption(f'Kolom No. SPBU atau Nama Brand tidak ditemukan. Kolom tersedia: {list(df_b.columns)}')

            st.divider()

            # --- Export Excel ---
            st.markdown('**Export**')
            if st.button('Siapkan Export Occupancy', key='occ_export_btn'):
                st.session_state['_occ_export_ready'] = True

            if st.session_state.get('_occ_export_ready'):
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                from openpyxl.utils import get_column_letter

                BLUE_HDR  = '0070C0'
                WHITE     = 'FFFFFF'
                THIN_DARK = Side(border_style='thin', color='9DC3E6')
                THIN      = Side(border_style='thin', color='BDD7EE')
                hdr_border  = Border(left=THIN_DARK, right=THIN_DARK,
                                     top=THIN_DARK, bottom=THIN_DARK)
                data_border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
                hdr_fill  = PatternFill('solid', fgColor=BLUE_HDR)
                data_fill = PatternFill('solid', fgColor=WHITE)
                hdr_font  = Font(color=WHITE, bold=True, size=11, name='Calibri')
                data_font = Font(size=11, name='Calibri')
                hdr_align  = Alignment(horizontal='center', vertical='center',
                                       wrap_text=True)
                data_align = Alignment(horizontal='left', vertical='center',
                                       wrap_text=False)

                OCC_COL_WIDTHS = {
                    'Region': 18, 'No SPBU': 14, 'Provinsi': 18,
                    'Tersedia': 12, 'Tersewa': 12, 'Beroperasi': 14, 'Telah Berakhir': 16,
                    'Total': 10, '% Occupancy': 14,
                    'Kategori': 28, 'Jumlah': 12,
                    'No Tenant': 12, 'Jumlah Tenant': 14,
                    'Kota': 20, 'Status NFR': 20, 'Jumlah Lahan': 14,
                }

                def _style_occ_sheet(ws, df):
                    ws.sheet_view.showGridLines = False
                    # write header
                    for ci, col in enumerate(df.columns, 1):
                        cell = ws.cell(row=1, column=ci, value=col)
                        cell.fill = hdr_fill
                        cell.font = hdr_font
                        cell.border = hdr_border
                        cell.alignment = hdr_align
                        w = OCC_COL_WIDTHS.get(col, 14)
                        ws.column_dimensions[get_column_letter(ci)].width = w
                    ws.row_dimensions[1].height = 28
                    # write data rows — append values in bulk (1 call/row), then style
                    import pandas as _pd
                    _rows = df.where(_pd.notna(df), None).values.tolist()
                    for row_vals in _rows:
                        ws.append(row_vals)
                    for ri in range(2, len(_rows) + 2):
                        for ci in range(1, len(df.columns) + 1):
                            cell = ws.cell(row=ri, column=ci)
                            cell.fill = data_fill
                            cell.font = data_font
                            cell.border = data_border
                            cell.alignment = data_align
                        ws.row_dimensions[ri].height = 15
                    # freeze + autofilter
                    ws.freeze_panes = 'A2'
                    ws.auto_filter.ref = ws.dimensions

                from openpyxl import Workbook as _WB
                wb = _WB()
                sheets = [
                    ('Per Region',  summary),
                    ('Klasifikasi SPBU', spbu_table),
                ]
                if beroperasi_uploaded and not kat.empty:
                    sheets.append(('Per Kategori', kat))
                if beroperasi_uploaded and not top_spbu.empty:
                    sheets.append(('Top SPBU', top_spbu))

                first = True
                for sname, sdf in sheets:
                    if first:
                        ws = wb.active
                        ws.title = sname
                        first = False
                    else:
                        ws = wb.create_sheet(sname)
                    _style_occ_sheet(ws, sdf)

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                fname = f'occupancy_nfr_{datetime.date.today().strftime("%Y-%m-%d")}.xlsx'
                st.download_button('Download Excel Occupancy',
                                   data=buf.getvalue(),
                                   file_name=fname,
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as _occ_err:
        st.error(f'Error di tab Occupancy: {_occ_err}')
        import traceback
        st.code(traceback.format_exc(), language='python')


# ----- tab Cleansing (uses Master FULL from tab Occupancy if uploaded) -----
with tabs[7]:
    try:
        st.markdown("""
<div style="margin-bottom:0.5rem;">
    <div style="font-size:1.1rem;font-weight:700;color:#1A1A1A;">Cleansing — Deteksi Duplikat Kode Lahan</div>
    <div style="font-size:0.82rem;color:#666;margin-top:2px;">
        Identifikasi data double di Master Lahan FULL. Master diambil otomatis dari <b>sidebar</b>.
        <br>Dua sinyal duplikat: <b>(A)</b> Kode Lahan muncul &gt;1× dan <b>(B)</b> Status Approval "Ubah" (edit pending).
    </div>
</div>
""", unsafe_allow_html=True)

        show_metodologi('cleansing')

        _cleans_master_bytes = st.session_state.get('_dash_occ_master')
        if not _cleans_master_bytes:
            st.info('Upload **NFR Master Lahan FULL** di **sidebar** untuk memulai deteksi duplikat.')
        else:
            dup = detect_duplicates(_cleans_master_bytes)
            if dup.empty:
                st.success('✅ Tidak ditemukan duplikat Kode Lahan di master ini.')
            else:
                _kode_col = 'Kode Lahan'
                _tipe_col = 'Tipe Duplikat'

                # --- KPI cards ---
                n_identik = int((dup[_tipe_col] == 'Identik').sum())
                n_beda    = int((dup[_tipe_col] == 'Beda Nilai').sum())
                n_mixed   = int((dup[_tipe_col] == 'Mixed Status').sum())
                n_ubah    = int((dup[_tipe_col] == 'Ubah Pending').sum())
                n_dup_kode = dup[_kode_col].nunique()
                n_rows = len(dup)

                # Row 1: Overview metrics
                k1, k2, k3 = st.columns(3)
                k1.metric('Kode Duplikat', f'{n_dup_kode:,}')
                k2.metric('Total Baris', f'{n_rows:,}')
                k3.metric('Mixed Status', f'{n_mixed:,}',
                          help='1 Kode muncul sebagai Tersedia + Tersewa. Bug paling berbahaya.')

                # Row 2: Breakdown by duplicate type
                k4, k5, k6 = st.columns(3)
                k4.metric('Beda Nilai', f'{n_beda:,}',
                          help='Kode sama, Luas/Harga berbeda (reinput).')
                k5.metric('Identik', f'{n_identik:,}',
                          help='Copy persis — simpan 1, hapus sisanya.')
                k6.metric('Ubah Pending', f'{n_ubah:,}',
                          help='Edit pending — versi baru tanpa hapus lama.')

                st.divider()

                # --- Breakdown chart per Tipe Duplikat ---
                da, db = st.columns([1, 2])
                with da:
                    st.markdown('**Breakdown per Tipe**')
                    tipe_vc = dup[_tipe_col].value_counts().reset_index()
                    tipe_vc.columns = ['Tipe Duplikat', 'Jumlah Baris']
                    st.dataframe(tipe_vc, width='stretch', hide_index=True)
                with db:
                    st.markdown('**Distribusi per Tipe Duplikat**')
                    st.bar_chart(tipe_vc.set_index('Tipe Duplikat')['Jumlah Baris'])

                st.divider()

                # --- Breakdown per Region ---
                st.markdown('**Duplikat per Region**')
                if 'Region' in dup.columns:
                    reg_dup = (dup.groupby('Region')[_kode_col].nunique()
                               .reset_index().rename(columns={_kode_col: 'Jumlah Kode Duplikat'})
                               .sort_values('Jumlah Kode Duplikat', ascending=False))
                    st.dataframe(reg_dup, width='stretch', hide_index=True)
                    st.bar_chart(reg_dup.set_index('Region')['Jumlah Kode Duplikat'])

                st.divider()

                # --- Filter & detail table ---
                st.markdown('**Detail Duplikat**')
                fc1, fc2, fc3 = st.columns(3)
                _tipe_opts = ['Semua'] + sorted(dup[_tipe_col].dropna().unique().tolist())
                _sel_tipe = fc1.selectbox('Filter Tipe Duplikat', _tipe_opts, key='cln_tipe')
                _sel_reg  = fc2.selectbox('Filter Region',
                                          ['Semua'] + sorted(dup['Region'].dropna().unique().tolist()),
                                          key='cln_reg') if 'Region' in dup.columns else 'Semua'
                _search = fc3.text_input('Cari Kode Lahan / No SPBU', key='cln_search')

                _disp = dup.copy()
                if _sel_tipe != 'Semua':
                    _disp = _disp[_disp[_tipe_col] == _sel_tipe]
                if _sel_reg != 'Semua' and 'Region' in _disp.columns:
                    _disp = _disp[_disp['Region'] == _sel_reg]
                if _search:
                    _q = _search.strip().lower()
                    _disp = _disp[_disp[_kode_col].astype(str).str.lower().str.contains(_q, na=False) |
                                  _disp['No SPBU'].astype(str).str.lower().str.contains(_q, na=False)]

                # kolom tampilan
                _show_cols = ['No SPBU', _kode_col, 'Luas (M2)', 'Harga (Rp)', 'Tenant',
                              'Tanggal Mulai Sewa', 'Tanggal Berakhir Sewa', 'Harga Sewa (Rp)',
                              'Status Sewa', 'Status Approval', 'Region',
                              'Tipe Duplikat', 'Jumlah Kemunculan', 'Sinyal', 'Catatan']
                _show_cols = [c for c in _show_cols if c in _disp.columns]
                _disp = _disp[_show_cols]

                st.caption(f'Menampilkan {len(_disp):,} baris dari {len(dup):,} baris duplikat.')
                st.dataframe(_disp, width='stretch', hide_index=True,
                             column_config={
                                 'Harga (Rp)':       st.column_config.NumberColumn(format='Rp %,.0f'),
                                 'Harga Sewa (Rp)':  st.column_config.NumberColumn(format='Rp %,.0f'),
                                 'Luas (M2)':        st.column_config.NumberColumn(format='%.1f'),
                                 'Jumlah Kemunculan': st.column_config.NumberColumn(format='%d'),
                             })

                st.divider()

                # --- Export Excel ---
                st.markdown('**Export**')
                if st.button('Siapkan Export Cleansing', key='cln_export_btn'):
                    st.session_state['_cln_export_ready'] = True

                if st.session_state.get('_cln_export_ready'):
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter
                    from openpyxl import Workbook as _WB2

                    RED_HDR   = 'C00000'
                    WHITE     = 'FFFFFF'
                    THIN      = Side(border_style='thin', color='BFBFBF')
                    bdr2      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
                    hdr_fill2 = PatternFill('solid', fgColor=RED_HDR)
                    hdr_font2 = Font(color=WHITE, bold=True, size=11, name='Calibri')
                    hdr_al    = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    dat_al    = Alignment(horizontal='left', vertical='center', wrap_text=False)
                    dat_font2 = Font(size=11, name='Calibri')
                    dat_fill2 = PatternFill('solid', fgColor=WHITE)

                    CLN_WIDTHS = {'No SPBU': 14, 'Kode Lahan': 22, 'Luas (M2)': 12,
                                  'Harga (Rp)': 18, 'Tenant': 24, 'Tanggal Mulai Sewa': 16,
                                  'Tanggal Berakhir Sewa': 18, 'Harga Sewa (Rp)': 18,
                                  'Status Sewa': 14, 'Status Approval': 22, 'Region': 16,
                                  'Tipe Duplikat': 16, 'Jumlah Kemunculan': 18,
                                  'Sinyal': 18, 'Catatan': 50}

                    def _style_cln(ws, df):
                        ws.sheet_view.showGridLines = False
                        for ci, col in enumerate(df.columns, 1):
                            cell = ws.cell(row=1, column=ci, value=col)
                            cell.fill = hdr_fill2; cell.font = hdr_font2
                            cell.border = bdr2; cell.alignment = hdr_al
                            ws.column_dimensions[get_column_letter(ci)].width = CLN_WIDTHS.get(col, 14)
                        ws.row_dimensions[1].height = 28
                        for row in df.where(pd.notna(df), None).values.tolist():
                            ws.append(row)
                        for ri in range(2, len(df) + 2):
                            for ci in range(1, len(df.columns) + 1):
                                cell = ws.cell(row=ri, column=ci)
                                cell.fill = dat_fill2; cell.font = dat_font2
                                cell.border = bdr2; cell.alignment = dat_al
                            ws.row_dimensions[ri].height = 15
                        ws.freeze_panes = 'A2'
                        ws.auto_filter.ref = ws.dimensions

                    wb2 = _WB2()
                    # Sheet 1: Summary
                    ws_s = wb2.active
                    ws_s.title = 'Summary'
                    ws_s.sheet_view.showGridLines = False
                    ws_s.merge_cells('A1:F1')
                    t = ws_s['A1']
                    t.value = 'Laporan Deteksi Duplikat Kode Lahan'
                    t.font = Font(bold=True, size=14, name='Calibri')
                    t.alignment = Alignment(horizontal='left', vertical='center')
                    ws_s.row_dimensions[1].height = 28
                    _kpis = [('Kode Duplikat', n_dup_kode), ('Total Baris', n_rows),
                             ('Mixed Status', n_mixed), ('Beda Nilai', n_beda),
                             ('Identik', n_identik), ('Ubah Pending', n_ubah)]
                    for ci, (lab, val) in enumerate(_kpis, 1):
                        c1 = ws_s.cell(row=3, column=ci, value=lab)
                        c1.fill = PatternFill('solid', fgColor=RED_HDR)
                        c1.font = Font(color=WHITE, bold=True, size=9)
                        c1.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        c1.border = bdr2
                        c2 = ws_s.cell(row=4, column=ci, value=val)
                        c2.fill = PatternFill('solid', fgColor='F4F6F9')
                        c2.font = Font(bold=True, size=16)
                        c2.alignment = Alignment(horizontal='center', vertical='center')
                        c2.border = bdr2
                        ws_s.column_dimensions[get_column_letter(ci)].width = 22
                    ws_s.row_dimensions[3].height = 30
                    ws_s.row_dimensions[4].height = 30

                    # Sheet 2: Detail
                    ws_d2 = wb2.create_sheet('Detail Duplikat')
                    _style_cln(ws_d2, dup[_show_cols])

                    # Sheet 3: Mixed Status priority
                    _mixed_df = dup[dup[_tipe_col] == 'Mixed Status'][_show_cols] if n_mixed > 0 else None
                    if _mixed_df is not None and not _mixed_df.empty:
                        ws_m = wb2.create_sheet('Mixed Status (Prioritas)')
                        _style_cln(ws_m, _mixed_df)

                    buf2 = io.BytesIO()
                    wb2.save(buf2)
                    buf2.seek(0)
                    _fname2 = f'cleansing_duplikat_{datetime.date.today().strftime("%Y-%m-%d")}.xlsx'
                    st.download_button('Download Excel Cleansing',
                                       data=buf2.getvalue(),
                                       file_name=_fname2,
                                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as _cln_err:
        st.error(f'Error di tab Cleansing: {_cln_err}')
        import traceback
        st.code(traceback.format_exc(), language='python')


# ----- tab Progress Input (SPBU belum input data lahan) -----
with tabs[8]:
    try:
        st.markdown("""
<div style="margin-bottom:0.5rem;">
    <div style="font-size:1.1rem;font-weight:700;color:#1A1A1A;">Progress Input — SPBU Belum Input Data Lahan</div>
    <div style="font-size:0.82rem;color:#666;margin-top:2px;">
        Upload <b>DDMS / List Lembaga Penyalur</b> (master semua SPBU) untuk identifikasi SPBU yang belum input data lahan NFR.
        <br>Data Master Lahan diambil otomatis dari tab <b>Occupancy</b> bila sudah diupload.
    </div>
</div>
""", unsafe_allow_html=True)

        show_metodologi('progress')

        pc1, pc2 = st.columns(2)
        ddms_up = pc1.file_uploader('DDMS / List Lembaga Penyalur (xlsx)',
                                    type=['xlsx'], key='ddms_up',
                                    help='File master seluruh SPBU se-Indonesia (datatarikanmasterddms atau List Lembaga Penyalur). '
                                         'Filter SPBU Reguler otomatis.')
        master_source = pc2.radio('Sumber Master Lahan NFR',
                                  ['Dari sidebar (otomatis)', 'Upload manual'],
                                  horizontal=True, key='pi_master_src',
                                  help='Pilih sumber data Master Lahan untuk perbandingan.')

        _pi_master_bytes = None
        if master_source.startswith('Dari'):
            _pi_master_bytes = st.session_state.get('_dash_occ_master')
        else:
            _pi_manual_up = pc2.file_uploader('NFR Master Lahan FULL (xlsx)', type=['xlsx'],
                                              key='pi_master_up')
            _pi_master_bytes = to_bytes(_pi_manual_up) if _pi_manual_up else None

        if not ddms_up:
            st.info('Upload **DDMS / List Lembaga Penyalur** untuk memulai analisis progress input.')
        elif _pi_master_bytes is None:
            st.info('Data Master Lahan NFR belum tersedia. '
                    'Upload di **sidebar** atau pilih "Upload manual" di atas.')
        else:
            _ddms_bytes = to_bytes(ddms_up)
            # Cache di session state agar persist
            st.session_state['_pi_ddms_bytes'] = _ddms_bytes
            st.session_state['_pi_master_bytes'] = _pi_master_bytes

            df_ddms = load_ddms(_ddms_bytes)
            df_master = occ_load_master(_pi_master_bytes)

            if df_ddms.empty:
                st.error('Tidak bisa membaca file DDMS. Pastikan format kolom benar (AgenNo / No. SPBU).')
            else:
                ddms_spbu = set(df_ddms['No SPBU'].astype(str).str.strip())
                master_spbu = set(df_master['No SPBU'].astype(str).str.strip())
                sudah_input = ddms_spbu & master_spbu
                belum_input = ddms_spbu - master_spbu

                n_ddms = len(ddms_spbu)
                n_sudah = len(sudah_input)
                n_belum = len(belum_input)
                pct_belum = round(n_belum / n_ddms * 100, 1) if n_ddms else 0

                # --- KPI cards ---
                k1, k2, k3, k4 = st.columns(4)
                k1.metric('Total SPBU Reguler', f'{n_ddms:,}')
                k2.metric('Sudah Input Lahan', f'{n_sudah:,}')
                k3.metric('Belum Input', f'{n_belum:,} ({pct_belum}%)')
                if 'Progress Input (%)' in df_ddms.columns:
                    avg_prog = df_ddms['Progress Input (%)'].dropna().mean()
                    k4.metric('Rata-rata Progress', f'{avg_prog:.0f}%' if avg_prog == avg_prog else '-')
                else:
                    k4.metric('Coverage', f'{n_sudah}/{n_ddms}')

                st.divider()

                # --- Progress buckets ---
                if 'Progress Input (%)' in df_ddms.columns:
                    st.markdown('**Distribusi Progress Input (SPBU Reguler)**')
                    def _bucket(p):
                        if pd.isna(p) or p is None: return 'NULL'
                        if p == 0: return '0%'
                        if p < 50: return '1-49%'
                        if p < 80: return '50-79%'
                        if p < 90: return '80-89%'
                        if p < 100: return '90-99%'
                        return '100%'
                    df_ddms['_bucket'] = df_ddms['Progress Input (%)'].apply(_bucket)
                    bucket_order = ['0%', '1-49%', '50-79%', '80-89%', '90-99%', '100%', 'NULL']
                    bkt = (df_ddms['_bucket'].value_counts().reindex(bucket_order).fillna(0).astype(int)
                           .reset_index())
                    bkt.columns = ['Progress', 'Jumlah SPBU']
                    bkt['Persentase'] = (bkt['Jumlah SPBU'] / n_ddms * 100).round(1)
                    bc1, bc2 = st.columns([1, 2])
                    with bc1:
                        st.dataframe(bkt, width='stretch', hide_index=True)
                    with bc2:
                        st.bar_chart(bkt.set_index('Progress')['Jumlah SPBU'])
                    st.divider()

                # --- Region breakdown ---
                st.markdown('**SPBU Belum Input per Region**')
                belum_df = df_ddms[df_ddms['No SPBU'].astype(str).str.strip().isin(belum_input)].copy()
                if 'Region' in belum_df.columns:
                    reg_belum = (belum_df.groupby('Region')['No SPBU'].nunique()
                                 .reset_index().rename(columns={'No SPBU': 'Belum Input'})
                                 .sort_values('Belum Input', ascending=False))
                    # bandingkan dengan total per region
                    reg_total = (df_ddms.groupby('Region')['No SPBU'].nunique()
                                 .reset_index().rename(columns={'No SPBU': 'Total'}))
                    reg_merge = reg_total.merge(reg_belum, on='Region', how='left')
                    reg_merge['Belum Input'] = reg_merge['Belum Input'].fillna(0).astype(int)
                    reg_merge['% Belum Input'] = (reg_merge['Belum Input'] / reg_merge['Total'].replace(0,1) * 100).round(1)
                    reg_merge = reg_merge.sort_values('Belum Input', ascending=False)
                    st.dataframe(reg_merge, width='stretch', hide_index=True,
                                 column_config={
                                     'Total': st.column_config.NumberColumn(format='%d'),
                                     'Belum Input': st.column_config.NumberColumn(format='%d'),
                                     '% Belum Input': st.column_config.ProgressColumn(
                                         '% Belum Input', format='%.1f%%', min_value=0, max_value=100),
                                 })
                    st.bar_chart(reg_merge.set_index('Region')[['Total', 'Belum Input']])
                st.divider()

                # --- Filter & detail table ---
                st.markdown('**Detail SPBU Belum Input**')
                fc1, fc2, fc3 = st.columns(3)
                _reg_opts = ['Semua'] + sorted(belum_df['Region'].dropna().unique().tolist()) if 'Region' in belum_df.columns else ['Semua']
                _sel_reg = fc1.selectbox('Filter Region', _reg_opts, key='pi_reg')
                _bkt_opts = ['Semua', '0%', '1-49%', '50-79%', '80-89%', '90-99%', '100%']
                _sel_bkt = fc2.selectbox('Filter Progress', _bkt_opts, key='pi_bkt') if '_bucket' in belum_df.columns else 'Semua'
                _search = fc3.text_input('Cari No SPBU / Nama', key='pi_search')

                _disp = belum_df.copy()
                if _sel_reg != 'Semua' and 'Region' in _disp.columns:
                    _disp = _disp[_disp['Region'] == _sel_reg]
                if _sel_bkt != 'Semua' and '_bucket' in _disp.columns:
                    _disp = _disp[_disp['_bucket'] == _sel_bkt]
                if _search:
                    _q = _search.strip().lower()
                    _disp = _disp[_disp['No SPBU'].astype(str).str.lower().str.contains(_q, na=False) |
                                  _disp.get('Nama SPBU', pd.Series(dtype=str)).astype(str).str.lower().str.contains(_q, na=False)]

                _show_cols = ['No SPBU', 'Nama SPBU', 'Provinsi', 'Kota/Kab', 'Region',
                              'Progress Input (%)', 'Status Operasi']
                _show_cols = [c for c in _show_cols if c in _disp.columns]
                _disp2 = _disp[_show_cols].copy()

                st.caption(f'Menampilkan {len(_disp2):,} dari {n_belum:,} SPBU belum input.')
                st.dataframe(_disp2, width='stretch', hide_index=True,
                             column_config={
                                 'Progress Input (%)': st.column_config.NumberColumn(format='%.0f%%'),
                             } if 'Progress Input (%)' in _show_cols else None)

                st.divider()

                # --- Export Excel ---
                st.markdown('**Export**')
                if st.button('Siapkan Export Progress Input', key='pi_export_btn'):
                    st.session_state['_pi_export_ready'] = True

                if st.session_state.get('_pi_export_ready'):
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter
                    from openpyxl import Workbook as _WB3

                    GREEN_HDR = '5A8A00'
                    WHITE     = 'FFFFFF'
                    THIN3     = Side(border_style='thin', color='BFBFBF')
                    bdr3      = Border(left=THIN3, right=THIN3, top=THIN3, bottom=THIN3)
                    hdr_f3    = PatternFill('solid', fgColor=GREEN_HDR)
                    hdr_ft3   = Font(color=WHITE, bold=True, size=11, name='Calibri')
                    hdr_al3   = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    dat_al3   = Alignment(horizontal='left', vertical='center')
                    dat_ft3   = Font(size=11, name='Calibri')
                    dat_f3    = PatternFill('solid', fgColor=WHITE)

                    def _style_pi(ws, df):
                        ws.sheet_view.showGridLines = False
                        for ci, col in enumerate(df.columns, 1):
                            cell = ws.cell(row=1, column=ci, value=col)
                            cell.fill = hdr_f3; cell.font = hdr_ft3
                            cell.border = bdr3; cell.alignment = hdr_al3
                        ws.row_dimensions[1].height = 28
                        for row in df.where(pd.notna(df), None).values.tolist():
                            ws.append(row)
                        for ri in range(2, len(df) + 2):
                            for ci in range(1, len(df.columns) + 1):
                                cell = ws.cell(row=ri, column=ci)
                                cell.fill = dat_f3; cell.font = dat_ft3
                                cell.border = bdr3; cell.alignment = dat_al3
                            ws.row_dimensions[ri].height = 15
                        ws.freeze_panes = 'A2'
                        ws.auto_filter.ref = ws.dimensions

                    wb3 = _WB3()
                    # Sheet 1: Summary
                    ws_s3 = wb3.active
                    ws_s3.title = 'Summary'
                    ws_s3.sheet_view.showGridLines = False
                    ws_s3.merge_cells('A1:D1')
                    t3 = ws_s3['A1']
                    t3.value = 'Laporan Progress Input SPBU — Data Lahan NFR'
                    t3.font = Font(bold=True, size=14, name='Calibri')
                    t3.alignment = Alignment(horizontal='left', vertical='center')
                    ws_s3.row_dimensions[1].height = 28
                    _kpis3 = [('Total SPBU Reguler', n_ddms), ('Sudah Input', n_sudah),
                              ('Belum Input', n_belum), ('% Belum', pct_belum)]
                    for ci, (lab, val) in enumerate(_kpis3, 1):
                        c1 = ws_s3.cell(row=3, column=ci, value=lab)
                        c1.fill = PatternFill('solid', fgColor=GREEN_HDR)
                        c1.font = Font(color=WHITE, bold=True, size=9)
                        c1.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        c1.border = bdr3
                        c2 = ws_s3.cell(row=4, column=ci, value=val)
                        c2.fill = PatternFill('solid', fgColor='F4F6F9')
                        c2.font = Font(bold=True, size=16)
                        c2.alignment = Alignment(horizontal='center', vertical='center')
                        c2.border = bdr3
                        ws_s3.column_dimensions[get_column_letter(ci)].width = 22
                    ws_s3.row_dimensions[3].height = 30
                    ws_s3.row_dimensions[4].height = 30

                    # Sheet 2: Per Region
                    if 'Region' in belum_df.columns:
                        ws_r3 = wb3.create_sheet('Per Region')
                        _style_pi(ws_r3, reg_merge)

                    # Sheet 3: Progress buckets
                    if 'Progress Input (%)' in df_ddms.columns:
                        ws_b3 = wb3.create_sheet('Progress Bucket')
                        _style_pi(ws_b3, bkt)

                    # Sheet 4: Detail belum input
                    ws_d3 = wb3.create_sheet('Detail Belum Input')
                    _style_pi(ws_d3, _disp2)

                    buf3 = io.BytesIO()
                    wb3.save(buf3)
                    buf3.seek(0)
                    _fname3 = f'progress_input_spbu_{datetime.date.today().strftime("%Y-%m-%d")}.xlsx'
                    st.download_button('Download Excel Progress Input',
                                       data=buf3.getvalue(),
                                       file_name=_fname3,
                                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as _pi_err:
        st.error(f'Error di tab Progress Input: {_pi_err}')
        import traceback
        st.code(traceback.format_exc(), language='python')


# ----- tab Tenant Cleansing (uses Master FULL from tab Occupancy) -----
with tabs[9]:
    try:
        st.markdown("""
<div style="margin-bottom:0.5rem;">
    <div style="font-size:1.1rem;font-weight:700;color:#1A1A1A;">Tenant — Cleansing Data Tenant</div>
    <div style="font-size:0.82rem;color:#666;margin-top:2px;">
        Identifikasi masalah data tenant di Master Lahan FULL (baris Tersewa). Master diambil otomatis dari <b>sidebar</b>.
        <br>Lima tipe temuan: <b>Kontrak Expired</b>, <b>Expiring ≤90 hari</b>, <b>Tenant Kosong</b>, <b>Harga Sewa Kosong</b>, <b>Tgl Sewa Kosong</b>.
    </div>
</div>
""", unsafe_allow_html=True)

        show_metodologi('tenant')

        _tc_master_bytes = st.session_state.get('_dash_occ_master')
        if not _tc_master_bytes:
            st.info('Upload **NFR Master Lahan FULL** di **sidebar** untuk memulai cleansing data tenant.')
        else:
            tc = detect_tenant_issues(_tc_master_bytes)
            if tc.empty:
                st.success('✅ Tidak ada baris Tersewa di master ini.')
            else:
                _tipe_col = 'Tipe Temuan'
                _status_col = 'Status Temuan'

                # --- KPI cards (hitung non-eksklusif dari Tipe Temuan string) ---
                n_total = len(tc)
                _has = lambda label: tc[_tipe_col].str.contains(label, na=False)
                n_expired_all = int(_has('Kontrak Expired').sum())
                n_expiring_all = int(_has('Expiring ≤90 hari').sum())

                # Row 1: Overview and time-based issues
                k1, k2, k3 = st.columns(3)
                k1.metric('Total Tersewa', f'{n_total:,}')
                k2.metric('Kontrak Expired', f'{n_expired_all:,}',
                          help='Tgl Berakhir sudah lewat tapi Status masih Tersewa. Seharusnya Tersedia.')
                k3.metric('Expiring ≤90 hari', f'{n_expiring_all:,}',
                          help='Kontrak berakhir dalam 90 hari — perlu perpanjangan.')

                # Row 2: Data completeness issues
                k4, k5, k6 = st.columns(3)
                k4.metric('Tenant Kosong', f'{int(_has("Tenant Kosong").sum()):,}',
                          help='Status Tersewa tapi nama Tenant kosong — inkonsistensi.')
                k5.metric('Harga Sewa Kosong', f'{int(_has("Harga Sewa Kosong").sum()):,}',
                          help='Harga Sewa (Rp) 0/null — nilai kontrak tidak diketahui.')
                k6.metric('Tgl Sewa Kosong', f'{int(_has("Tgl Sewa Kosong").sum()):,}',
                          help='Tgl Mulai atau Tgl Berakhir null — data sewa tidak lengkap.')

                st.divider()

                # --- Breakdown per Status Temuan ---
                da, db = st.columns([1, 2])
                with da:
                    st.markdown('**Breakdown per Status Temuan**')
                    st_vc = tc[_status_col].value_counts().reset_index()
                    st_vc.columns = ['Status Temuan', 'Jumlah Baris']
                    st.dataframe(st_vc, width='stretch', hide_index=True)
                with db:
                    st.markdown('**Distribusi per Status Temuan**')
                    st.bar_chart(st_vc.set_index('Status Temuan')['Jumlah Baris'])

                st.divider()

                # --- Per Region breakdown (expired per region) ---
                st.markdown('**Kontrak Expired per Region**')
                if 'Region' in tc.columns:
                    tc_expired = tc[tc[_tipe_col].str.contains('Kontrak Expired', na=False)]
                    reg_exp = (tc_expired.groupby('Region').size()
                               .reset_index().rename(columns={0: 'Kontrak Expired'})
                               .sort_values('Kontrak Expired', ascending=False))
                    # bandingkan total tersewa per region
                    reg_total = (tc.groupby('Region').size()
                                 .reset_index().rename(columns={0: 'Total Tersewa'}))
                    reg_merge = reg_total.merge(reg_exp, on='Region', how='left')
                    reg_merge['Kontrak Expired'] = reg_merge['Kontrak Expired'].fillna(0).astype(int)
                    reg_merge['% Expired'] = (reg_merge['Kontrak Expired'] / reg_merge['Total Tersewa'].replace(0,1) * 100).round(1)
                    st.dataframe(reg_merge, width='stretch', hide_index=True,
                                 column_config={
                                     'Total Tersewa': st.column_config.NumberColumn(format='%d'),
                                     'Kontrak Expired': st.column_config.NumberColumn(format='%d'),
                                     '% Expired': st.column_config.ProgressColumn(
                                         '% Expired', format='%.1f%%', min_value=0, max_value=100),
                                 })
                    st.bar_chart(reg_merge.set_index('Region')[['Total Tersewa', 'Kontrak Expired']])

                st.divider()

                # --- Filter & detail table ---
                st.markdown('**Detail Temuan**')
                fc1, fc2, fc3 = st.columns(3)
                _tipe_opts = ['Semua'] + sorted(tc[_status_col].dropna().unique().tolist())
                _sel_tipe = fc1.selectbox('Filter Status Temuan', _tipe_opts, key='tc_tipe')
                _sel_reg = fc2.selectbox('Filter Region',
                                        ['Semua'] + sorted(tc['Region'].dropna().unique().tolist()),
                                        key='tc_reg') if 'Region' in tc.columns else 'Semua'
                _search = fc3.text_input('Cari No SPBU / Tenant / Kode Lahan', key='tc_search')

                _disp = tc.copy()
                if _sel_tipe != 'Semua':
                    _disp = _disp[_disp[_status_col] == _sel_tipe]
                if _sel_reg != 'Semua' and 'Region' in _disp.columns:
                    _disp = _disp[_disp['Region'] == _sel_reg]
                if _search:
                    _q = _search.strip().lower()
                    _disp = _disp[_disp['No SPBU'].astype(str).str.lower().str.contains(_q, na=False) |
                                  _disp.get('Tenant', pd.Series(dtype=str)).astype(str).str.lower().str.contains(_q, na=False) |
                                  _disp.get('Kode Lahan', pd.Series(dtype=str)).astype(str).str.lower().str.contains(_q, na=False)]

                _show_cols = ['No SPBU', 'Kode Lahan', 'Tenant', 'Tanggal Mulai Sewa',
                              'Tanggal Berakhir Sewa', 'Harga Sewa (Rp)', 'Status Temuan',
                              'Tipe Temuan', 'Region', 'Catatan']
                _show_cols = [c for c in _show_cols if c in _disp.columns]
                _disp2 = _disp[_show_cols].copy()

                st.caption(f'Menampilkan {len(_disp2):,} dari {len(tc):,} baris Tersewa.')
                _cc = {}
                if 'Harga Sewa (Rp)' in _show_cols:
                    _cc['Harga Sewa (Rp)'] = st.column_config.NumberColumn(format='Rp %,.0f')
                if 'Tanggal Mulai Sewa' in _show_cols:
                    _cc['Tanggal Mulai Sewa'] = st.column_config.DateColumn(format='YYYY-MM-DD')
                if 'Tanggal Berakhir Sewa' in _show_cols:
                    _cc['Tanggal Berakhir Sewa'] = st.column_config.DateColumn(format='YYYY-MM-DD')
                st.dataframe(_disp2, width='stretch', hide_index=True,
                             column_config=_cc if _cc else None)

                st.divider()

                # --- Export Excel ---
                st.markdown('**Export**')
                if st.button('Siapkan Export Tenant Cleansing', key='tc_export_btn'):
                    st.session_state['_tc_export_ready'] = True

                if st.session_state.get('_tc_export_ready'):
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter
                    from openpyxl import Workbook as _WB4

                    ORANGE_HDR = 'F4B942'
                    WHITE = 'FFFFFF'
                    THIN4 = Side(border_style='thin', color='BFBFBF')
                    bdr4 = Border(left=THIN4, right=THIN4, top=THIN4, bottom=THIN4)
                    hdr_f4 = PatternFill('solid', fgColor=ORANGE_HDR)
                    hdr_ft4 = Font(color=WHITE, bold=True, size=11, name='Calibri')
                    hdr_al4 = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    dat_al4 = Alignment(horizontal='left', vertical='center')
                    dat_ft4 = Font(size=11, name='Calibri')
                    dat_f4 = PatternFill('solid', fgColor=WHITE)

                    TC_WIDTHS = {'No SPBU': 14, 'Kode Lahan': 22, 'Tenant': 28,
                                 'Tanggal Mulai Sewa': 18, 'Tanggal Berakhir Sewa': 18,
                                 'Harga Sewa (Rp)': 18, 'Status Temuan': 20, 'Tipe Temuan': 28,
                                 'Region': 16, 'Catatan': 60}

                    def _style_tc(ws, df):
                        ws.sheet_view.showGridLines = False
                        for ci, col in enumerate(df.columns, 1):
                            cell = ws.cell(row=1, column=ci, value=col)
                            cell.fill = hdr_f4; cell.font = hdr_ft4
                            cell.border = bdr4; cell.alignment = hdr_al4
                            ws.column_dimensions[get_column_letter(ci)].width = TC_WIDTHS.get(col, 14)
                        ws.row_dimensions[1].height = 28
                        for row in df.where(pd.notna(df), None).values.tolist():
                            ws.append(row)
                        for ri in range(2, len(df) + 2):
                            for ci in range(1, len(df.columns) + 1):
                                cell = ws.cell(row=ri, column=ci)
                                cell.fill = dat_f4; cell.font = dat_ft4
                                cell.border = bdr4; cell.alignment = dat_al4
                            ws.row_dimensions[ri].height = 15
                        ws.freeze_panes = 'A2'
                        ws.auto_filter.ref = ws.dimensions

                    wb4 = _WB4()
                    # Sheet 1: Summary
                    ws_s4 = wb4.active
                    ws_s4.title = 'Summary'
                    ws_s4.sheet_view.showGridLines = False
                    ws_s4.merge_cells('A1:F1')
                    t4 = ws_s4['A1']
                    t4.value = 'Laporan Cleansing Data Tenant'
                    t4.font = Font(bold=True, size=14, name='Calibri')
                    t4.alignment = Alignment(horizontal='left', vertical='center')
                    ws_s4.row_dimensions[1].height = 28
                    _kpis4 = [('Total Tersewa', n_total), ('Kontrak Expired', n_expired_all),
                              ('Expiring ≤90 hari', n_expiring_all),
                              ('Tenant Kosong', int(_has("Tenant Kosong").sum())),
                              ('Harga Sewa Kosong', int(_has("Harga Sewa Kosong").sum())),
                              ('Tgl Sewa Kosong', int(_has("Tgl Sewa Kosong").sum()))]
                    for ci, (lab, val) in enumerate(_kpis4, 1):
                        c1 = ws_s4.cell(row=3, column=ci, value=lab)
                        c1.fill = PatternFill('solid', fgColor=ORANGE_HDR)
                        c1.font = Font(color=WHITE, bold=True, size=9)
                        c1.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        c1.border = bdr4
                        c2 = ws_s4.cell(row=4, column=ci, value=val)
                        c2.fill = PatternFill('solid', fgColor='F4F6F9')
                        c2.font = Font(bold=True, size=16)
                        c2.alignment = Alignment(horizontal='center', vertical='center')
                        c2.border = bdr4
                        ws_s4.column_dimensions[get_column_letter(ci)].width = 22
                    ws_s4.row_dimensions[3].height = 30
                    ws_s4.row_dimensions[4].height = 30

                    # Sheet 2: Detail
                    ws_d4 = wb4.create_sheet('Detail Temuan')
                    _style_tc(ws_d4, _disp2)

                    # Sheet 3: Kontrak Expired (Prioritas)
                    _exp_df = tc[tc[_tipe_col].str.contains('Kontrak Expired', na=False)][_show_cols] if n_expired_all > 0 else None
                    if _exp_df is not None and not _exp_df.empty:
                        ws_e4 = wb4.create_sheet('Kontrak Expired (Prioritas)')
                        _style_tc(ws_e4, _exp_df)

                    buf4 = io.BytesIO()
                    wb4.save(buf4)
                    buf4.seek(0)
                    _fname4 = f'tenant_cleansing_{datetime.date.today().strftime("%Y-%m-%d")}.xlsx'
                    st.download_button('Download Excel Tenant Cleansing',
                                       data=buf4.getvalue(),
                                       file_name=_fname4,
                                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as _tc_err:
        st.error(f'Error di tab Tenant: {_tc_err}')
        import traceback
        st.code(traceback.format_exc(), language='python')

