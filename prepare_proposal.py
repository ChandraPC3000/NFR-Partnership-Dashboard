#!/usr/bin/env python3
"""
prepare_proposal.py — Dari export NFR Master Lahan → proposal siap kirim ke tenant.

Output: SATU file Excel berisi:
  - Sheet "Data"    : lokasi hasil filter (Kopken-style + enrichment)
  - Sheet "Charts"  : chart native Excel (bar per kota, histogram luas, scatter harga-luas)
  - Sheet "Peta"    : render peta lokasi (gambar)

Input:
  --master   : file export NFR Master Lahan (wajib)
  --template : file Template (enrichment: alamat, koordinat, SES, Red/Blue Ocean, harga)
  --pic      : file berisi kontak PIC (opsional; sheet "PIC" kolom: No SPBU, Nama, No HP)

Filter (semua opsional):
  --region   : Sumbagut|Sumbagsel|JBB|JBT|Jatimbalinus|Kalimantan|Sulawesi|Maluku Papua
  --kota     : substring Kota/Kab (mis. "Jakarta", "Bandung")
  --min-luas / --max-luas  : m2
  --min-harga / --max-harga: Rp
  --usaha    : kata kunci jenis usaha (dicocokkan ke Red Ocean/Blue Ocean, mis. "coffee", "grocery")
  --status   : default "Tersedia"
  --output   : prefiks nama file (default "Proposal")

Contoh:
  python3 prepare_proposal.py \
      --master "NFR Master Lahan 19-August-2026.xlsx" \
      --template "Available Space NFR at SPBU - August 2026 V1 Template.xlsx" \
      --pic "Available Space NFR at SPBU - August 2026 V1 Kopken.xlsx" \
      --region JBB --kota "Jakarta Barat" --min-luas 20 --usaha coffee \
      --output "Proposal_ClientX"
"""
import argparse, datetime, os, sys
from collections import Counter
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

REGIONS = ['Sumbagut', 'Sumbagsel', 'JBB', 'JBT', 'Jatimbalinus',
           'Kalimantan', 'Sulawesi', 'Maluku Papua']
REGION_DIGIT = {name: str(i + 1) for i, name in enumerate(REGIONS)}
NAVY = '0D2F4F'

# Palet warna Excel (accent) utk bar chart — tiap bar beda warna (gaya default Excel).
CHART_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
                '264478', '9E480E', '636363', '997300', '2F5597', 'C55A11']


def read_sheet(path, sheet, header_row, colmap):
    """Baca sheet -> list of dict (kunci = nama kolom dari colmap {nama: idx 1-based}).

    Cepat: pakai pd.read_excel (C parser) + usecols hanya kolom yang dipakai.
    Sekali pass, bukan cell-by-cell openpyxl (50-100x lebih lambat untuk 31k baris).
    """
    import pandas as pd
    usecols = [c - 1 for c in colmap.values()]              # 0-based untuk pandas
    df = pd.read_excel(path, sheet_name=sheet, header=header_row - 1,
                       usecols=usecols, engine='openpyxl')
    df.columns = list(colmap.keys())                        # rename ke nama logis
    # NaN -> None agar logika downstream (or / _num / is not None) sama seperti openpyxl
    df = df.where(pd.notna(df), None)
    return df.to_dict('records')


def _header_positions(path, sheet, header_row):
    """Baca baris header -> {nama_header: posisi 1-based}. Nama di-strip; duplikat: ambil pertama."""
    import pandas as pd
    hdr = pd.read_excel(path, sheet_name=sheet, header=None, nrows=header_row,
                        engine='openpyxl')
    row = hdr.iloc[header_row - 1].tolist()
    out = {}
    for i, v in enumerate(row):
        if pd.notna(v):
            nm = str(v).strip()
            if nm and nm not in out:
                out[nm] = i + 1
    return out


def _resolve_colmap(path, sheet, header_row, fields):
    """Bangun {nama_logis: posisi 1-based} dari NAMA HEADER (bukan posisi tetap).

    fields = {nama_logis: [kandidat nama header]}. Tahan terhadap kolom
    ditambah/dipindah di template baru (mis. 'Management Fee' bikin semua
    kolom setelahnya bergeser — mapping by-name tetap benar).
    """
    pos = _header_positions(path, sheet, header_row)
    cm = {}
    for logical, candidates in fields.items():
        for cand in candidates:
            if cand in pos:
                cm[logical] = pos[cand]
                break
    return cm


# Field master -> kandidat nama header
_MASTER_FIELDS = {
    'No SPBU':        ['No SPBU', 'No. SPBU', 'SPBU', 'AgenNo'],
    'Kode Lahan':     ['Kode Lahan'],
    'Luas':           ['Luas (M2)', 'Luas'],
    'Harga':          ['Harga (Rp)'],
    'Tenant':         ['Tenant'],
    'Tgl Mulai':      ['Tanggal Mulai Sewa', 'Tgl Mulai Sewa'],
    'Tgl Akhir':      ['Tanggal Berakhir Sewa', 'Tgl Akhir Sewa'],
    'Harga Sewa':     ['Harga Sewa (Rp)'],
    'Status Sewa':    ['Status Sewa'],
    'Status Approval':['Status Approval'],
}

# Field template -> kandidat nama header (September 2026: 'Harga (Rp)' -> 'Harga Sewa (Rp)',
# tambah 'Management Fee' → semua kolom setelahnya bergeser)
_TEMPLATE_FIELDS = {
    'No SPBU':         ['No SPBU', 'No. SPBU'],
    'Kode Lahan':      ['Kode Lahan'],
    'Nama Perusahaan': ['Nama Perusahaan'],
    'Provinsi':        ['Provinsi'],
    'Kota/Kab':        ['Kota/Kab'],
    'Kecamatan':       ['Kecamatan'],
    'Desa':            ['Desa/Kelurahan', 'Desa'],
    'Alamat':          ['Alamat'],
    'Tipe SPBU':       ['Tipe SPBU'],
    'Kelas':           ['Kelas SPBU', 'Kelas'],
    'Luas':            ['Luas (M2)', 'Luas'],
    'Harga':           ['Harga (Rp)', 'Harga Sewa (Rp)'],
    'Latitude':        ['Latitude'],
    'Longitude':       ['Longitude'],
    'Road Type':       ['Road Type'],
    'SES':             ['SES'],
    'Red Ocean':       ['Red Ocean'],
    'Blue Ocean':      ['Blue Ocean'],
    'Status Sewa':     ['Status Sewa'],
    'GES':             ['Green Energy Station (GES)', 'GES'],
    'EV':              ['Fasilitas EV', 'EV'],
    'PSO':             ['PSO'],
    'Non PSO':         ['Non PSO'],
    'Total':           ['Total'],
    'PSO2':            ['PSO2'],
    'Non PSO2':        ['Non PSO2'],
    'Total2':          ['Total2'],
}


def load_master(path):
    cm = _resolve_colmap(path, 'NFR Master Lahan', 1, _MASTER_FIELDS)
    if not cm:
        return {}
    rows = read_sheet(path, 'NFR Master Lahan', 1, cm)
    out = {}
    for r in rows:
        k = str(r.get('Kode Lahan') or '').strip()
        if k:
            out.setdefault(k, r)
    return out


def load_template(path):
    cm = _resolve_colmap(path, 'Data', 4, _TEMPLATE_FIELDS)
    if not cm:
        return {}
    rows = read_sheet(path, 'Data', 4, cm)
    out = {}
    for r in rows:
        k = str(r.get('Kode Lahan') or '').strip()
        if k:
            out.setdefault(k, r)
    return out


def _clean(v):
    """Bersihkan nilai kontak; string kosong / 'NULL' -> None."""
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.upper() == 'NULL':
        return None
    return s


def load_pic(path):
    """Baca kontak PIC per No SPBU. Kunci join = No SPBU; Nama boleh kosong ('NULL').

    Mengembalikan {No SPBU: {'nama','hp','email','office'}}.
    Mendukung header: 'No SPBU'/'AgentNo', 'Nama/PIC/Manager', 'No HP', 'Email', 'Phone Office'.
    """
    import pandas as pd
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return {}

    def _find_hdr(ws):
        """Scan baris 1-7 cari header 'spbu'/'agent'. Murah (<=7 baris)."""
        max_c = ws.max_column or 1
        for r in range(1, min(8, (ws.max_row or 1) + 1)):
            vals = [ws.cell(row=r, column=c).value for c in range(1, max_c + 1)]
            if any(v and ('spbu' in str(v).lower() or 'agent' in str(v).lower()) for v in vals):
                return r
        return 1

    def by_header(ws_name):
        ws = wb[ws_name]
        hdr = _find_hdr(ws)
        max_c = ws.max_column or 1
        # deteksi keyword->index dari header (selalu murah, 1 baris x max_c)
        def col_idx(*kws):
            for c in range(1, max_c + 1):
                v = str(ws.cell(row=hdr, column=c).value or '').lower()
                if any(k in v for k in kws):
                    return c
            return None
        c_spbu = col_idx('spbu', 'agent')
        if not c_spbu:
            return {}
        c_nama = col_idx('nama', 'pic', 'manager')
        c_hp = col_idx('hp', 'telp', 'phone', 'kontak')
        c_email = col_idx('email')
        c_office = col_idx('office', 'kantor')
        # Baca seluruh data sekaligus via pd.read_excel (C parser, bukan per-cell)
        df = pd.read_excel(path, sheet_name=ws_name, header=hdr - 1, engine='openpyxl')
        out = {}
        # Map 1-based idx -> pandas column position (kolom header yang terbaca pd)
        cols = df.columns.tolist()
        def _val(row_pos, one_based_idx):
            if one_based_idx is None or one_based_idx - 1 >= len(cols):
                return None
            v = df.iloc[row_pos, one_based_idx - 1]
            return None if pd.isna(v) else v
        for i in range(len(df)):
            spbu = _val(i, c_spbu)
            if spbu is None:
                continue
            out[str(spbu).strip()] = {
                'nama':   _clean(_val(i, c_nama)),
                'hp':     _clean(_val(i, c_hp)),
                'email':  _clean(_val(i, c_email)),
                'office': _clean(_val(i, c_office)),
            }
        return out

    if 'PIC' in wb.sheetnames:
        return by_header('PIC')

    # fallback gaya file Kopken: kolom 3 header-nya 'No SPBU'
    if 'Sheet1' in wb.sheetnames:
        ws = wb['Sheet1']
        hdr = _find_hdr(ws)
        if hdr and 'spbu' in str(ws.cell(row=hdr, column=3).value or '').lower():
            df = pd.read_excel(path, sheet_name='Sheet1', header=hdr - 1, engine='openpyxl')
            out = {}
            for i in range(len(df)):
                spbu = df.iloc[i, 2] if len(df.columns) > 2 else None   # kolom 3 = idx 2
                if spbu is None or pd.isna(spbu):
                    continue
                def _g(pos):
                    v = df.iloc[i, pos - 1] if pos and pos - 1 < len(df.columns) else None
                    return _clean(None if (v is not None and pd.isna(v)) else v)
                out[str(spbu).strip()] = {
                    'nama': _g(13), 'hp': _g(14), 'email': None, 'office': None,
                }
            return out

    # default: deteksi kolom dari header di sheet pertama
    return by_header(wb.sheetnames[0])


def build_proposal(rows, out_path):
    """rows: list dict hasil filter (gabungan master+template+PIC)."""
    from openpyxl.chart import Reference
    wb = Workbook()

    # ---------------- Sheet Data ----------------
    ws = wb.active
    ws.title = 'Data'
    cols = ['No SPBU', 'Kode Lahan', 'Provinsi', 'Kota/Kab', 'Kecamatan', 'Alamat',
            'Tipe SPBU', 'Luas (M2)', 'Harga (Rp)', 'Status Sewa', 'Latitude', 'Longitude',
            'Road Type', 'SES', 'Red Ocean', 'Blue Ocean', 'PIC', 'No HP']
    hf = PatternFill('solid', fgColor=NAVY)
    hfont = Font(color='FFFFFF', bold=True)
    for ci, c in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=c)
        cell.fill = hf; cell.font = hfont
    for ri, r in enumerate(rows, 2):
        for ci, c in enumerate(cols, 1):
            ws.cell(row=ri, column=ci, value=r.get(c))
    for col, w in zip('ABCDEFGHIJKLMNOPQR', [9, 20, 13, 18, 16, 45, 11, 10, 13, 12, 11, 11, 20, 22, 28, 28, 20, 14]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:R{len(rows)+1}"

    # ---------------- Sheet Charts ----------------
    ws2 = wb.create_sheet('Charts')
    # tabel agregasi: jumlah per kota
    ws2['A1'] = 'Kota/Kab'; ws2['B1'] = 'Jumlah'
    per_kota = Counter(r.get('Kota/Kab') or '(tanpa kota)' for r in rows)
    rr = 2
    for k, n in per_kota.most_common(15):
        ws2.cell(row=rr, column=1, value=k); ws2.cell(row=rr, column=2, value=n); rr += 1
    bar = BarChart(); bar.type = 'col'; bar.title = 'Jumlah Lokasi per Kota'
    bar.add_data(Reference(ws2, min_col=2, min_row=1, max_row=rr-1), titles_from_data=True)
    bar.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=rr-1))
    ws2.add_chart(bar, 'D2')

    # histogram luas (bin)
    ws2['F1'] = 'Luas (M2)'; ws2['G1'] = 'Jumlah'
    bins = [(0, 20), (20, 40), (40, 60), (60, 100), (100, 200), (200, 1000)]
    br = 2
    for lo, hi in bins:
        n = sum(1 for r in rows if (r.get('Luas (M2)') or 0) and lo <= (r.get('Luas (M2)') or 0) < hi)
        ws2.cell(row=br, column=6, value=f"{lo}-{hi}")
        ws2.cell(row=br, column=7, value=n); br += 1
    bar2 = BarChart(); bar2.type = 'col'; bar2.title = 'Sebaran Luas (M2)'
    bar2.add_data(Reference(ws2, min_col=7, min_row=1, max_row=br-1), titles_from_data=True)
    bar2.set_categories(Reference(ws2, min_col=6, min_row=2, max_row=br-1))
    ws2.add_chart(bar2, 'F16')

    # scatter harga vs luas — ref ke sheet Data
    n_data = len(rows) + 1
    sc = ScatterChart(); sc.title = 'Harga vs Luas'; sc.style = 13
    sc.x_axis.title = 'Luas (M2)'; sc.y_axis.title = 'Harga (Rp)'
    xref = Reference(ws, min_col=8, min_row=2, max_row=n_data)
    yref = Reference(ws, min_col=9, min_row=2, max_row=n_data)
    sc.series.append(Series(yref, xref, title='Lokasi'))
    ws2.add_chart(sc, 'J2')

    # ---------------- Sheet Peta ----------------
    ws3 = wb.create_sheet('Peta')
    map_png = render_map(rows)
    img = XLImage(map_png)
    ws3.add_image(img, 'B2')
    ws3.column_dimensions['B'].width = 60

    wb.save(out_path)
    return out_path


def render_map(rows):
    """Render peta lokasi ke PNG menggunakan tile map (staticmap + OpenStreetMap).

    Fallback ke matplotlib scatter bila staticmap tidak terinstall atau tidak ada koneksi.
    """
    import io

    pts = [(r.get('Latitude'), r.get('Longitude'), r.get('Status Sewa') or '')
           for r in rows
           if r.get('Latitude') and r.get('Longitude')
           and -11 <= (r.get('Latitude') or 0) <= 6
           and 95 <= (r.get('Longitude') or 0) <= 141]

    if not pts:
        return _render_map_empty()

    try:
        import socket
        # Batasi durasi download tile OpenStreetMap — tanpa timeout, host yang
        # lambat/unreachable bisa membuat export menggantung tak bisa di-stop.
        _prev_timeout = socket.getdefaulttimeout()
        socket.setdefaulttimeout(6)
        try:
            from staticmap import StaticMap, CircleMarker
            # Tentukan zoom & ukuran: lebih banyak titik → zoom lebih jauh
            n = len(pts)
            size = 1200 if n > 50 else 900

            # Downsampel bila titik terlalu banyak: 2 marker/circle per titik; untuk
            # 31k lokasi artinya 62k circle digambar satu-satu di Python — lambat.
            # Di PNG ukuran 1200px, ribuan titik saling tindih → sampling tak terlihat.
            _MAP_CAP = 3000
            if n > _MAP_CAP:
                step = max(1, n // _MAP_CAP)
                pts = pts[::step]
                n = len(pts)

            m = StaticMap(size, size,
                          url_template='https://tile.openstreetmap.org/{z}/{x}/{y}.png')

            for lat, lon, status in pts:
                color = '#5A8A00' if status == 'Tersedia' else ('#E31B23' if status == 'Tersewa' else '#888888')
                m.add_marker(CircleMarker((lon, lat), color, 10))
                m.add_marker(CircleMarker((lon, lat), 'white', 4))  # titik putih di tengah

            img = m.render()
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            buf.seek(0)
            return buf
        finally:
            socket.setdefaulttimeout(_prev_timeout)

    except Exception:
        # Fallback ke matplotlib bila staticmap error / offline
        return _render_map_matplotlib(pts)


def _render_map_empty():
    import io, matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    plt.figure(figsize=(8, 6))
    plt.text(0.5, 0.5, 'Tidak ada data koordinat', ha='center', va='center', fontsize=14)
    plt.axis('off')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=100)
    plt.close()
    buf.seek(0)
    return buf


def _render_map_matplotlib(pts):
    import io, matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    # Cap juga di fallback: 31k titik digambar matplotlib juga lambat.
    # Di PNG 1200px, sampling tak terlihat bedanya.
    _MAP_CAP = 3000
    if len(pts) > _MAP_CAP:
        step = max(1, len(pts) // _MAP_CAP)
        pts = pts[::step]
    lats = [p[0] for p in pts]
    lons = [p[1] for p in pts]
    colors = ['#5A8A00' if p[2] == 'Tersedia' else '#E31B23' for p in pts]
    plt.figure(figsize=(10, 8))
    plt.scatter(lons, lats, c=colors, s=60, alpha=0.8, edgecolor='white')
    plt.xlabel('Longitude'); plt.ylabel('Latitude')
    plt.title('Peta Lokasi'); plt.grid(alpha=0.3)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf


def _num(v):
    """Konversi ke float bila memungkinkan, else None (untuk data campuran string/angka)."""
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def build_dataframe(master, templ, pic=None):
    """Gabung SEMUA kolom master + template + PIC -> pandas DataFrame.

    Kolom yang ada di dua-duanya: nilai diambil master dulu, kecuali
    'Harga (Rp)' -> pakai Template (harga sebenarnya; master sering 0 untuk Tersedia).
    """
    import pandas as pd
    pic = pic or {}
    rows = []
    for k, m in master.items():
        t = templ.get(k, {})
        no_spbu = m.get('No SPBU') or t.get('No SPBU') or ''
        p = pic.get(str(no_spbu).strip(), {})
        rows.append({
            # --- master (lease/status) ---
            'No SPBU': no_spbu, 'Kode Lahan': k,
            'Tenant': m.get('Tenant'),
            'Tgl Mulai Sewa': m.get('Tgl Mulai'),
            'Tgl Akhir Sewa': m.get('Tgl Akhir'),
            'Harga Sewa (Rp)': _num(m.get('Harga Sewa')),
            'Status Sewa': m.get('Status Sewa'),
            'Status Approval': m.get('Status Approval'),
            # --- template (market/geo/address) ---
            'Nama Perusahaan': t.get('Nama Perusahaan'),
            'Provinsi': t.get('Provinsi'), 'Kota/Kab': t.get('Kota/Kab'),
            'Kecamatan': t.get('Kecamatan'), 'Desa/Kelurahan': t.get('Desa'),
            'Alamat': t.get('Alamat'), 'Tipe SPBU': t.get('Tipe SPBU'),
            'Kelas SPBU': t.get('Kelas'),
            'Luas (M2)': _num(m.get('Luas') if m.get('Luas') is not None else t.get('Luas')),
            'Harga (Rp)': _num(t.get('Harga') if t.get('Harga') is not None else m.get('Harga')),
            'Latitude': _num(t.get('Latitude')), 'Longitude': _num(t.get('Longitude')),
            'Road Type': t.get('Road Type'), 'SES': t.get('SES'),
            'Red Ocean': t.get('Red Ocean'), 'Blue Ocean': t.get('Blue Ocean'),
            'GES': t.get('GES'), 'EV': t.get('EV'),
            'PSO': _num(t.get('PSO')), 'Non PSO': _num(t.get('Non PSO')),
            'Total': _num(t.get('Total')), 'PSO2': _num(t.get('PSO2')),
            'Non PSO2': _num(t.get('Non PSO2')), 'Total2': _num(t.get('Total2')),
            # --- PIC (opsional) ---
            'PIC': p.get('nama'), 'No HP': p.get('hp'),
            'Email': p.get('email'), 'Phone Office': p.get('office'),
        })
    out = pd.DataFrame(rows)
    # Normalisasi kolom mixed-type (campur int & str) jadi string — kalau dibiarkan,
    # pyarrow/st.dataframe error & lambat (fallback Arrow tiap render 31k baris).
    for _c in out.columns:
        if out[_c].dtype == object:
            _nonnull = out[_c].dropna()
            if not _nonnull.empty:
                _has_str = any(isinstance(v, str) for v in _nonnull)
                _has_non = any(not isinstance(v, str) for v in _nonnull)
                if _has_str and _has_non:
                    out[_c] = out[_c].astype(str).str.strip()
    # No SPBU selalu string (kunci join & filter; file campur int & str).
    if 'No SPBU' in out.columns:
        out['No SPBU'] = out['No SPBU'].astype(str).str.strip()
    return out


def filter_df(df, region=None, statuses=None, kota=None,
              min_luas=None, max_luas=None, min_harga=None, max_harga=None,
              usaha=None, provinsi=None, tipe_spbu=None,
              throughput_col='Total', min_throughput_hari=None,
              transaksi_col='Total2', min_transaksi_bulan=None):
    """Terapkan filter pada DataFrame.

    Throughput: kolom PSO/Non PSO/Total dalam kL/bulan -> konversi /30 untuk kL/hari.
    Transaksi: kolom PSO2/Non PSO2/Total2 dalam transaksi/bulan.
    """
    out = df
    if region:
        if isinstance(region, (list, tuple, set)):
            digits = {REGION_DIGIT[r] for r in region if r in REGION_DIGIT}
            out = out[out['No SPBU'].astype(str).str[0].isin(digits)]
        elif region != 'Semua':
            out = out[out['No SPBU'].astype(str).str[0] == REGION_DIGIT[region]]
    if statuses:
        out = out[out['Status Sewa'].isin(statuses)]
    if provinsi:
        out = out[out['Provinsi'].isin(provinsi)]
    if kota:
        if isinstance(kota, (list, tuple, set)):
            out = out[out['Kota/Kab'].isin(kota)]
        else:
            out = out[out['Kota/Kab'].fillna('').str.contains(kota, case=False, na=False)]
    if tipe_spbu:
        out = out[out['Tipe SPBU'].isin(tipe_spbu)]
    if min_luas is not None:
        out = out[out['Luas (M2)'].fillna(-1) >= min_luas]
    if max_luas is not None:
        out = out[out['Luas (M2)'].fillna(1e9) <= max_luas]
    if min_harga is not None:
        out = out[out['Harga (Rp)'].fillna(-1) >= min_harga]
    if max_harga is not None:
        out = out[out['Harga (Rp)'].fillna(1e18) <= max_harga]
    if usaha:
        hay = (out['Red Ocean'].fillna('') + ' ' + out['Blue Ocean'].fillna('')).str.lower()
        out = out[hay.str.contains(usaha.lower(), na=False)]
    if min_throughput_hari is not None and throughput_col in out.columns:
        # Data dalam kL/bulan -> konversi ke kL/hari (/30)
        out = out[out[throughput_col].fillna(0) / 30 >= min_throughput_hari]
    if min_transaksi_bulan is not None and transaksi_col in out.columns:
        out = out[out[transaksi_col].fillna(0) >= min_transaksi_bulan]
    return out


def build_proposal_workbook(df, data_cols=None):
    """Workbook (Data + Charts + Peta) — styling mengikuti template Pertamina."""
    from openpyxl.utils import get_column_letter

    if data_cols is None:
        data_cols = [c for c in df.columns if c not in ('Latitude', 'Longitude')]

    BLUE_HDR   = '0070C0'   # biru header template
    WHITE      = 'FFFFFF'
    THIN       = Side(border_style='thin', color='BDD7EE')
    THIN_DARK  = Side(border_style='thin', color='9DC3E6')
    hdr_border = Border(left=THIN_DARK, right=THIN_DARK, top=THIN_DARK, bottom=THIN_DARK)
    data_border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    white_fill  = PatternFill('solid', fgColor=WHITE)

    # kolom lebar (template-like): default 14, overrides per nama kolom
    COL_WIDTHS = {
        'No SPBU': 12, 'Kode Lahan': 18, 'Nama Perusahaan': 36,
        'Provinsi': 18, 'Kota/Kab': 20, 'Kecamatan': 28,
        'Desa/Kelurahan': 24, 'Alamat': 60, 'Tipe SPBU': 12,
        'Kelas SPBU': 13, 'Luas (M2)': 12, 'Harga (Rp)': 18,
        'Latitude': 13, 'Longitude': 13, 'Road Type': 22,
        'SES': 52, 'Red Ocean': 52, 'Blue Ocean': 52,
        'Status Sewa': 14, 'GES': 14, 'EV': 12,
        'PSO': 12, 'Non PSO': 12, 'Total': 10,
        'PSO2': 10, 'Non PSO2': 11, 'Total2': 10,
    }

    wb = Workbook()
    ws = wb.active
    ws.title = 'Data'

    # --- matikan gridlines ---
    ws.sheet_view.showGridLines = False

    # --- header row 1 ---
    hf    = PatternFill('solid', fgColor=BLUE_HDR)
    hfont = Font(color=WHITE, bold=True, size=11, name='Calibri')
    for ci, c in enumerate(data_cols, 1):
        cell = ws.cell(row=1, column=ci, value=c)
        cell.fill   = hf
        cell.font   = hfont
        cell.border = hdr_border
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=True)
    ws.row_dimensions[1].height = 28

    # --- data rows ---
    dfont = Font(size=11, name='Calibri')
    dalign = Alignment(horizontal='left', vertical='center', wrap_text=False)
    # Tulis nilai baris sekaligus via ws.append (1 call/baris, bukan N cell)
    import pandas as pd
    _data = df[data_cols].where(pd.notna(df[data_cols]), None).values.tolist()
    for row_vals in _data:
        ws.append(row_vals)
    # Styling per-cell data HANYA bila baris sedikit — styling 31k×27 sel ~40 dtk
    # (fill+font+border+alignment per sel), bikin export "menggantung" tak bisa di-stop.
    # Untuk baris banyak: header & struktur tetap di-style, data polos (Excel tetap rapi:
    # nilai + kolom lebar + freeze + autofilter tetap ada). Batas = 6000 baris.
    n_rows = len(_data)
    _STYLE_CAP = 6000
    if n_rows <= _STYLE_CAP:
        for ri in range(2, n_rows + 2):
            for ci in range(1, len(data_cols) + 1):
                cell = ws.cell(row=ri, column=ci)
                cell.fill   = white_fill
                cell.font   = dfont
                cell.border = data_border
                cell.alignment = dalign
            ws.row_dimensions[ri].height = 15

    # --- column widths ---
    for i, c in enumerate(data_cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = COL_WIDTHS.get(c, 14)

    # --- freeze: baris header + kolom No SPBU tetap terlihat saat scroll ---
    # No SPBU selalu kolom pertama; freeze di B2
    ws.freeze_panes = 'B2'

    # --- auto filter ---
    ws.auto_filter.ref = f"A1:{get_column_letter(len(data_cols))}1"

    # ---- Charts ----
    ws2 = wb.create_sheet('Charts')
    ws2.sheet_view.showGridLines = False

    # data tabel (kolom A-B: kota; D-E: luas)
    hf2 = PatternFill('solid', fgColor=BLUE_HDR)
    _hdr = lambda ws, r, c, v: _set_hdr(ws, r, c, v, hf2,
                                          Font(color=WHITE, bold=True, size=10),
                                          hdr_border)

    per_kota = df['Kota/Kab'].fillna('(tanpa kota)').value_counts().head(15)
    _hdr(ws2, 1, 1, 'Kota/Kab'); _hdr(ws2, 1, 2, 'Jumlah')
    for i, (k, n) in enumerate(per_kota.items(), 2):
        ws2.cell(row=i, column=1, value=k).font = Font(size=10)
        ws2.cell(row=i, column=2, value=int(n)).font = Font(size=10)

    luas_vals = df['Luas (M2)']
    _hdr(ws2, 1, 4, 'Luas (M2)'); _hdr(ws2, 1, 5, 'Jumlah')
    br = 2
    for lo, hi in [(0, 20), (20, 40), (40, 60), (60, 100), (100, 200), (200, 1000)]:
        n = int(luas_vals.between(lo, hi, inclusive='left').sum())
        ws2.cell(row=br, column=4, value=f"{lo}-{hi}").font = Font(size=10)
        ws2.cell(row=br, column=5, value=n).font = Font(size=10)
        br += 1

    bar = BarChart(); bar.type = 'col'; bar.title = 'Jumlah Lokasi per Kota'
    bar.width = 14; bar.height = 10; bar.style = 10
    last = len(per_kota) + 1
    bar.add_data(Reference(ws2, min_col=2, min_row=1, max_row=last), titles_from_data=True)
    bar.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=last))
    ws2.add_chart(bar, 'I2')

    bar2 = BarChart(); bar2.type = 'col'; bar2.title = 'Sebaran Luas (M2)'
    bar2.width = 14; bar2.height = 10; bar2.style = 10
    bar2.add_data(Reference(ws2, min_col=5, min_row=1, max_row=br - 1), titles_from_data=True)
    bar2.set_categories(Reference(ws2, min_col=4, min_row=2, max_row=br - 1))
    ws2.add_chart(bar2, 'I22')

    if 'Luas (M2)' in data_cols and 'Harga (Rp)' in data_cols:
        li = data_cols.index('Luas (M2)') + 1
        hi_idx = data_cols.index('Harga (Rp)') + 1
        sc = ScatterChart(); sc.title = 'Harga vs Luas'; sc.style = 10
        sc.x_axis.title = 'Luas (M2)'; sc.y_axis.title = 'Harga (Rp)'
        sc.width = 14; sc.height = 10
        sc.series.append(Series(Reference(ws, min_col=hi_idx, min_row=2, max_row=len(df)+1),
                                Reference(ws, min_col=li, min_row=2, max_row=len(df)+1),
                                title='Lokasi'))
        ws2.add_chart(sc, 'I42')

    # ---- Peta (hanya untuk export kecil; skip bila baris banyak) ----
    # render_map = download tile OSM + gambar ribuan titik → lambat & bergantung
    # jaringan. Untuk export besar (mis. >2000 baris) peta image tidak informatif
    # (ribuan titik bertumpuk) & bikin "Menyiapkan Excel" lama. Data/Charts tetap.
    _MAP_SHEET_MAX = 2000
    if len(df) <= _MAP_SHEET_MAX:
        ws3 = wb.create_sheet('Peta')
        ws3.sheet_view.showGridLines = False
        img = XLImage(render_map(df.to_dict('records')))
        ws3.add_image(img, 'B2')
        ws3.column_dimensions['B'].width = 80
        ws3.row_dimensions[2].height = 400

    return wb


def _xw_point_colors(n_points):
    """Daftar format per-titik (dPt) utk xlsxwriter add_series — tiap bar warna beda."""
    return [{'fill': {'color': CHART_COLORS[i % len(CHART_COLORS)]}} for i in range(n_points)]


def build_proposal_workbook_fast(df, data_cols=None, include_charts=True, summary_df=None):
    """Workbook Excel CEPAT utk data besar — pakai xlsxwriter (~2x lbh cepat dari openpyxl).

    Sheet Data diisi dari `df` (hasil filter). Sheet Ringkasan Region/Kota/SPBU
    dihitung dari `summary_df` (FULL master) bila diberikan — supaya angka
    Tersedia/Tersewa/Total tetap utuh walau Data hanya berisi baris terfilter.
    """
    import io
    import xlsxwriter

    if data_cols is None:
        data_cols = [c for c in df.columns if c not in ('Latitude', 'Longitude')]
    data_cols = list(data_cols)

    # summary_df = FULL master utk sheet Ringkasan (default: pakai df = sama dgn Data)
    _df_sum = summary_df if summary_df is not None else df
    # Kolom lebar (ringkas; Alamat lebih lebar)
    _W = {'Alamat': 60, 'Nama Perusahaan': 36, 'Kecamatan': 28, 'Desa/Kelurahan': 24,
          'Red Ocean': 52, 'Blue Ocean': 52, 'SES': 52, 'Kode Lahan': 18,
          'No SPBU': 12, 'Luas (M2)': 12, 'Harga (Rp)': 18, 'Latitude': 13,
          'Longitude': 13, 'Provinsi': 18, 'Kota/Kab': 20, 'Status Sewa': 14}
    buf = io.BytesIO()
    # constant_memory TIDAK dipakai: mode itu hanya utk tulis-baris-berurutan & MENJATUHKAN
    # sel di kolom lain (mis. kolom G/H utk tabel chart) / sheet kedua. Tanpa mode itu,
    # 31k baris tetap cepat (~5-6 dtk) & chart/data tersimpan benar.
    wb = xlsxwriter.Workbook(buf)
    ws = wb.add_worksheet('Data')
    ws.hide_gridlines(2)

    hdr_fmt = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#0070C0',
                             'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                             'border': 1})
    # header
    for ci, c in enumerate(data_cols):
        ws.write_string(0, ci, str(c), hdr_fmt)
        ws.set_column(ci, ci, _W.get(c, 14))
    ws.set_row(0, 28)
    ws.freeze_panes(1, 1)

    # data — sanitize NaN/NaT -> None SEKALI (bukan per-sel), lalu write_row dengan SATU
    # format border (di-reuse semua cell — xlsxwriter efisien, ~+2s utk 31k baris).
    import pandas as pd
    def _clean(v):
        if v is None:
            return None
        try:
            if pd.isna(v):
                return None
        except (TypeError, ValueError):
            pass
        return v
    dat_fmt = wb.add_format({'border': 1, 'valign': 'vcenter'})
    _rows = [[_clean(v) for v in row] for row in df[data_cols].values.tolist()]
    for ri, row in enumerate(_rows, start=1):
        ws.write_row(ri, 0, row, dat_fmt)

    ws.autofilter(0, 0, len(_rows), len(data_cols) - 1)

    # ---- Charts (murah: referensi range, bukan salin 31k baris) ----
    if include_charts and 'Kota/Kab' in df.columns:
        ws2 = wb.add_worksheet('Charts')
        ws2.hide_gridlines(2)   # sama spt Data: tanpa gridline -> area kosong tampak putih
        # Format sama seperti sheet Data: header biru (0070C0) + border, sel data putih+border
        ch_hdr = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#0070C0',
                                'align': 'center', 'valign': 'vcenter', 'border': 1})
        ch_dat = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter'})
        ch_dat_num = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                    'num_format': '#,##0'})

        # tabel ringkas per kota (top 15) — kolom A/B
        kota = df['Kota/Kab'].fillna('(tanpa kota)').value_counts().head(15)
        ws2.write_string(0, 0, 'Kota/Kab', ch_hdr)
        ws2.write_string(0, 1, 'Jumlah', ch_hdr)
        for i, (k, n) in enumerate(kota.items(), start=1):
            ws2.write_string(i, 0, str(k), ch_dat)
            ws2.write_number(i, 1, int(n), ch_dat_num)
        # histogram luas — tabel di kolom D/E
        luas_vals = df['Luas (M2)']
        ws2.write_string(0, 3, 'Luas (M2)', ch_hdr)
        ws2.write_string(0, 4, 'Jumlah', ch_hdr)
        _br = 1
        for lo, hi in [(0, 20), (20, 40), (40, 60), (60, 100), (100, 200), (200, 1000)]:
            n = int(luas_vals.between(lo, hi, inclusive='left').sum())
            ws2.write_string(_br, 3, f"{lo}-{hi}", ch_dat)
            ws2.write_number(_br, 4, n, ch_dat_num)
            _br += 1
        n_luas = _br - 1

        # ---- chart 1: Jumlah Lokasi per Kota (log-scale agar bar kecil terlihat) ----
        bar = wb.add_chart({'type': 'column'})
        bar.add_series({
            'name': 'Jumlah',
            'categories': ['Charts', 1, 0, len(kota), 0],
            'values':     ['Charts', 1, 1, len(kota), 1],
            'points': _xw_point_colors(len(kota)),
        })
        bar.set_title({'name': 'Jumlah Lokasi per Kota'})
        bar.set_style(10)
        # skala log: tanpa kota (25k) vs kota lain (139…) — bar kecil baru terlihat
        bar.set_y_axis({'log_base': 10})
        bar.set_size({'width': 500, 'height': 320})
        bar.set_plotarea({'fill': {'color': 'white'}, 'border': {'none': True}})
        ws2.insert_chart('I2', bar)

        # ---- chart 2: Sebaran Luas (M2), ditaruh di bawah chart 1 (vertical) ----
        bar2 = wb.add_chart({'type': 'column'})
        bar2.add_series({
            'name': 'Jumlah',
            'categories': ['Charts', 1, 3, n_luas, 3],
            'values':     ['Charts', 1, 4, n_luas, 4],
            'points': _xw_point_colors(n_luas),
        })
        bar2.set_title({'name': 'Sebaran Luas (M2)'})
        bar2.set_style(10)
        bar2.set_size({'width': 500, 'height': 320})
        bar2.set_plotarea({'fill': {'color': 'white'}, 'border': {'none': True}})
        ws2.insert_chart('I28', bar2)

        # ---- scatter harga vs luas — di bawah chart 2 (kolom I, vertikal) ----
        if 'Luas (M2)' in data_cols and 'Harga (Rp)' in data_cols:
            li = data_cols.index('Luas (M2)')
            hi_idx = data_cols.index('Harga (Rp)')
            sc = wb.add_chart({'type': 'scatter'})
            sc.add_series({
                'name': 'Lokasi',
                'categories': ['Data', 1, li, len(_rows), li],
                'values':     ['Data', 1, hi_idx, len(_rows), hi_idx],
            })
            sc.set_title({'name': 'Harga vs Luas'})
            sc.set_x_axis({'name': 'Luas (M2)'})
            sc.set_y_axis({'name': 'Harga (Rp)'})
            sc.set_style(10)
            sc.set_size({'width': 500, 'height': 320})
            sc.set_plotarea({'fill': {'color': 'white'}, 'border': {'none': True}})
            ws2.insert_chart('I54', sc)

    # ---- Ringkasan Region sheet (spt contoh Excel) ----
    # Derive Region dari digit-1 No SPBU bila belum ada (spt logika app_nfr).
    _df_reg = _df_sum
    if 'Region' not in _df_reg.columns and 'No SPBU' in _df_reg.columns:
        _d2r = {str(i + 1): r for i, r in enumerate(REGIONS)}
        _df_reg = _df_reg.copy()
        _df_reg['Region'] = (_df_reg['No SPBU'].astype(str).str.strip().str[0]
                             .map(_d2r).fillna('(lainnya)'))
    if 'Region' in _df_reg.columns:
        ws3 = wb.add_worksheet('Ringkasan Region')
        ws3.hide_gridlines(2)
        reg_hdr = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#0070C0',
                                 'align': 'center', 'valign': 'vcenter', 'border': 1})
        reg_dat = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter'})
        reg_num = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                'num_format': '#,##0'})
        reg_avg = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                'num_format': '#,##0.0'})

        reg_headers = ['Region', 'Total Lokasi', 'Tersedia', 'Tersewa',
                       'Rata-rata Luas (m²)', 'Rata-rata Harga (Rp)']
        _widths = [16, 14, 12, 12, 20, 22]
        for ci, h in enumerate(reg_headers):
            ws3.write_string(0, ci, h, reg_hdr)
            ws3.set_column(ci, ci, _widths[ci])
        ws3.set_row(0, 28)
        ws3.freeze_panes(1, 1)

        # Build agg dict conditionally (avoid ternary — lambdas reference columns)
        _reg_agg_d = {}
        _count_col = 'Kode Lahan' if 'Kode Lahan' in _df_reg.columns else 'No SPBU'
        _reg_agg_d['total'] = (_count_col, 'count')
        if 'Status Sewa' in _df_reg.columns:
            _reg_agg_d['tersedia'] = ('Status Sewa', lambda x: (x.astype(str).str.strip() == 'Tersedia').sum())
            _reg_agg_d['tersewa']  = ('Status Sewa', lambda x: (x.astype(str).str.strip() == 'Tersewa').sum())
        else:
            _reg_agg_d['tersedia'] = (_count_col, lambda x: 0)
            _reg_agg_d['tersewa']  = (_count_col, lambda x: 0)
        if 'Luas (M2)' in _df_reg.columns:
            _reg_agg_d['rata_luas'] = ('Luas (M2)', 'mean')
        else:
            _reg_agg_d['rata_luas'] = (_count_col, lambda x: 0)
        if 'Harga (Rp)' in _df_reg.columns:
            _reg_agg_d['rata_harga'] = ('Harga (Rp)', 'mean')
        else:
            _reg_agg_d['rata_harga'] = (_count_col, lambda x: 0)

        _reg_agg = _df_reg.groupby('Region', dropna=False).agg(**_reg_agg_d).reset_index()

        for ri, row in _reg_agg.iterrows():
            ws3.write_string(ri + 1, 0, str(row['Region']), reg_dat)
            ws3.write_number(ri + 1, 1, int(row['total']), reg_num)
            ws3.write_number(ri + 1, 2, int(row['tersedia']), reg_num)
            ws3.write_number(ri + 1, 3, int(row['tersewa']), reg_num)
            ws3.write_number(ri + 1, 4, float(row['rata_luas']), reg_avg)
            ws3.write_number(ri + 1, 5, float(row['rata_harga']), reg_num)
        ws3.autofilter(0, 0, len(_reg_agg), len(reg_headers) - 1)

    # ---- Ringkasan Kota/Kab sheet (spt contoh Excel) ----
    if 'Kota/Kab' in _df_reg.columns:
        ws4 = wb.add_worksheet('Ringkasan Kota')
        ws4.hide_gridlines(2)
        kota_hdr = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#0070C0',
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})
        kota_dat = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter'})
        kota_num = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                  'num_format': '#,##0'})
        kota_avg = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                 'num_format': '#,##0.0'})

        kota_headers = ['Kota/Kab', 'Total Lokasi', 'Tersedia', 'Tersewa',
                        'Rata-rata Luas (m²)', 'Rata-rata Harga (Rp)']
        _kwidths = [24, 14, 12, 12, 20, 22]
        for ci, h in enumerate(kota_headers):
            ws4.write_string(0, ci, h, kota_hdr)
            ws4.set_column(ci, ci, _kwidths[ci])
        ws4.set_row(0, 28)
        ws4.freeze_panes(1, 1)

        _kota_agg_d = {}
        _kota_count = 'Kode Lahan' if 'Kode Lahan' in _df_reg.columns else 'No SPBU'
        _kota_agg_d['total'] = (_kota_count, 'count')
        if 'Status Sewa' in _df_reg.columns:
            _kota_agg_d['tersedia'] = ('Status Sewa', lambda x: (x.astype(str).str.strip() == 'Tersedia').sum())
            _kota_agg_d['tersewa']  = ('Status Sewa', lambda x: (x.astype(str).str.strip() == 'Tersewa').sum())
        else:
            _kota_agg_d['tersedia'] = (_kota_count, lambda x: 0)
            _kota_agg_d['tersewa']  = (_kota_count, lambda x: 0)
        if 'Luas (M2)' in _df_reg.columns:
            _kota_agg_d['rata_luas'] = ('Luas (M2)', 'mean')
        else:
            _kota_agg_d['rata_luas'] = (_kota_count, lambda x: 0)
        if 'Harga (Rp)' in _df_reg.columns:
            _kota_agg_d['rata_harga'] = ('Harga (Rp)', 'mean')
        else:
            _kota_agg_d['rata_harga'] = (_kota_count, lambda x: 0)

        _kota_agg = (_df_reg.groupby('Kota/Kab', dropna=False)
                      .agg(**_kota_agg_d)
                      .reset_index()
                      .sort_values('total', ascending=False))
        _kota_agg['Kota/Kab'] = _kota_agg['Kota/Kab'].fillna('(tanpa kota)')

        for r_idx, (_, row) in enumerate(_kota_agg.iterrows(), start=1):
            ws4.write_string(r_idx, 0, str(row['Kota/Kab']), kota_dat)
            ws4.write_number(r_idx, 1, int(row['total']), kota_num)
            ws4.write_number(r_idx, 2, int(row['tersedia']), kota_num)
            ws4.write_number(r_idx, 3, int(row['tersewa']), kota_num)
            ws4.write_number(r_idx, 4, float(row['rata_luas']), kota_avg)
            ws4.write_number(r_idx, 5, float(row['rata_harga']), kota_num)
        ws4.autofilter(0, 0, len(_kota_agg), len(kota_headers) - 1)

    # ---- Ringkasan SPBU sheet (spt contoh Excel) ----
    if 'No SPBU' in _df_reg.columns:
        ws5 = wb.add_worksheet('Ringkasan SPBU')
        ws5.hide_gridlines(2)
        spbu_hdr = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#0070C0',
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})
        spbu_dat = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter'})
        spbu_num = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                 'num_format': '#,##0'})
        spbu_avg = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'vcenter',
                                 'num_format': '#,##0.0'})

        spbu_headers = ['No SPBU', 'Region', 'Provinsi', 'Kota/Kab', 'Total Lahan',
                        'Rata-rata Luas (m²)', 'Rata-rata Harga (Rp)', 'Tersewa', 'Occupancy %']
        _swidths = [14, 16, 18, 20, 12, 18, 22, 12, 14]
        for ci, h in enumerate(spbu_headers):
            ws5.write_string(0, ci, h, spbu_hdr)
            ws5.set_column(ci, ci, _swidths[ci])
        ws5.set_row(0, 28)
        ws5.freeze_panes(1, 1)

        _spbu_agg_d = {}
        _spbu_count = 'Kode Lahan' if 'Kode Lahan' in _df_reg.columns else 'No SPBU'
        _spbu_agg_d['total'] = (_spbu_count, 'count')
        if 'Region' in _df_reg.columns:
            _spbu_agg_d['region'] = ('Region', 'first')
        if 'Provinsi' in _df_reg.columns:
            _spbu_agg_d['provinsi'] = ('Provinsi', 'first')
        if 'Kota/Kab' in _df_reg.columns:
            _spbu_agg_d['kota'] = ('Kota/Kab', 'first')
        if 'Luas (M2)' in _df_reg.columns:
            _spbu_agg_d['rata_luas'] = ('Luas (M2)', 'mean')
        if 'Harga (Rp)' in _df_reg.columns:
            _spbu_agg_d['rata_harga'] = ('Harga (Rp)', 'mean')
        if 'Status Sewa' in _df_reg.columns:
            _spbu_agg_d['tersewa'] = ('Status Sewa', lambda x: (x.astype(str).str.strip() == 'Tersewa').sum())
        else:
            _spbu_agg_d['tersewa'] = (_spbu_count, lambda x: 0)

        _spbu_agg = (_df_reg.groupby('No SPBU', dropna=False)
                      .agg(**_spbu_agg_d)
                      .reset_index()
                      .sort_values('total', ascending=False))
        # Occupancy % = Tersewa / Total × 100
        _spbu_agg['occupancy'] = (_spbu_agg['tersewa'] / _spbu_agg['total'].replace(0, 1) * 100).round(1)

        for r_idx, (_, row) in enumerate(_spbu_agg.iterrows(), start=1):
            ws5.write_string(r_idx, 0, str(row['No SPBU']), spbu_dat)
            ws5.write_string(r_idx, 1, str(row.get('region', '')), spbu_dat)
            ws5.write_string(r_idx, 2, str(row.get('provinsi', '')), spbu_dat)
            ws5.write_string(r_idx, 3, str(row.get('kota', '')), spbu_dat)
            ws5.write_number(r_idx, 4, int(row['total']), spbu_num)
            ws5.write_number(r_idx, 5, float(row.get('rata_luas', 0)), spbu_avg)
            ws5.write_number(r_idx, 6, float(row.get('rata_harga', 0)), spbu_num)
            ws5.write_number(r_idx, 7, int(row['tersewa']), spbu_num)
            ws5.write_number(r_idx, 8, float(row['occupancy']), spbu_avg)
        ws5.autofilter(0, 0, len(_spbu_agg), len(spbu_headers) - 1)

    wb.close()
    return buf.getvalue()


def _set_hdr(ws, r, c, v, fill, font, border):
    cell = ws.cell(row=r, column=c, value=v)
    cell.fill = fill; cell.font = font; cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')
    return cell


def main():
    ap = argparse.ArgumentParser(description='Buat proposal NFR dari master lahan')
    ap.add_argument('--master', required=True, help='File export NFR Master Lahan')
    ap.add_argument('--template', help='File Template (enrichment)')
    ap.add_argument('--pic', help='File PIC (opsional)')
    ap.add_argument('--region', choices=REGIONS, help='Filter region')
    ap.add_argument('--kota', help='Substring Kota/Kab')
    ap.add_argument('--min-luas', type=float); ap.add_argument('--max-luas', type=float)
    ap.add_argument('--min-harga', type=float); ap.add_argument('--max-harga', type=float)
    ap.add_argument('--usaha', help='Keyword jenis usaha (Red/Blue Ocean)')
    ap.add_argument('--status', default='Tersedia')
    ap.add_argument('--output', default='Proposal')
    args = ap.parse_args()

    print("Membaca master lahan ...")
    master = load_master(args.master)
    print(f"  master: {len(master)} kode unik")

    templ = {}
    if args.template:
        print("Membaca template (enrichment) ...")
        templ = load_template(args.template)
        print(f"  template: {len(templ)} kode")

    pic = {}
    if args.pic:
        print("Membaca PIC ...")
        pic = load_pic(args.pic)
        print(f"  PIC: {len(pic)} No SPBU")

    # pool = status sesuai filter
    pool = {k: v for k, v in master.items() if (v.get('Status Sewa') or '') == args.status}
    print(f"  status '{args.status}': {len(pool)}")

    rows = []
    for k, m in pool.items():
        t = templ.get(k, {})
        no_spbu = m.get('No SPBU') or t.get('No SPBU') or ''
        if args.region:
            if str(no_spbu)[0] != REGION_DIGIT[args.region]:
                continue
        kota = t.get('Kota/Kab') or ''
        if args.kota and args.kota.lower() not in str(kota).lower():
            continue
        luas = m.get('Luas') if m.get('Luas') is not None else t.get('Luas')
        if args.min_luas is not None and (luas is None or luas < args.min_luas): continue
        if args.max_luas is not None and (luas is None or luas > args.max_luas): continue
        harga = t.get('Harga') if t.get('Harga') is not None else m.get('Harga')
        if args.min_harga is not None and (harga is None or harga < args.min_harga): continue
        if args.max_harga is not None and (harga is None or harga > args.max_harga): continue
        if args.usaha:
            hay = f"{t.get('Red Ocean') or ''} {t.get('Blue Ocean') or ''}".lower()
            if args.usaha.lower() not in hay:
                continue
        pic_row = pic.get(str(no_spbu).strip(), {})
        rows.append({
            'No SPBU': no_spbu, 'Kode Lahan': k,
            'Provinsi': t.get('Provinsi'), 'Kota/Kab': kota, 'Kecamatan': t.get('Kecamatan'),
            'Alamat': t.get('Alamat'), 'Tipe SPBU': t.get('Tipe SPBU'),
            'Luas (M2)': luas, 'Harga (Rp)': harga,
            'Status Sewa': m.get('Status Sewa'),
            'Latitude': t.get('Latitude'), 'Longitude': t.get('Longitude'),
            'Road Type': t.get('Road Type'), 'SES': t.get('SES'),
            'Red Ocean': t.get('Red Ocean'), 'Blue Ocean': t.get('Blue Ocean'),
            'PIC': pic_row.get('nama'), 'No HP': pic_row.get('hp'),
        })

    print(f"Hasil filter: {len(rows)} lokasi")
    if not rows:
        print("Tidak ada hasil. Coba longgarkan filter.")
        sys.exit(1)

    today = datetime.date.today().isoformat()
    out_dir = os.path.dirname(os.path.abspath(args.master))
    out = os.path.join(out_dir, f"{args.output}_{today}.xlsx")
    build_proposal(rows, out)
    print(f"OK -> {out}")


if __name__ == '__main__':
    main()
