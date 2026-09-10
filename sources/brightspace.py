#!/usr/bin/env python3
"""
sources/brightspace.py — Konfigurasi & logika khusus brightspace.pertamina.com.

Adapter ini berisi SEMUA yang spesifik ke situs Brightspace:
endpoint, parameter filter, klasifikasi wilayah (Jakarta Barat, region),
normalisasi nama wilayah, dan jalur file data. Skrip di tools/ hanya
memakai fungsi dari sini + scrapekit.

Catatan penting (ditemukan saat analisis):
- Filter kota (IdKota) untuk lahan TIDAK berfungsi di server (mengembalikan
  semua lahan). Karena itu dipakai trik rentang halaman + filter alamat/koordinat.
- Data wilayah diambil dari field `regencyName` per SPBU (akurat 100%).
"""
import json, os, re

HERE = os.path.dirname(os.path.abspath(__file__))          # .../tools/sources
TOOLS = os.path.dirname(HERE)                              # .../tools
DELIVERABLES = os.path.dirname(TOOLS)                      # .../deliverables
DATA = os.path.join(DELIVERABLES, 'data')

API = 'https://brightspace.pertamina.com'
LAHAN_ITEMS_URL = API + '/list/lahan/items'
SPBU_ITEMS_URL = API + '/list/spbu/items'
REGENCIES_URL = API + '/list/regencies'

CACHE = '/tmp/price0_lahan.jsonl'                          # cache hasil scrape
PHOTO_MAP = os.path.join(DATA, 'photo_map.json')
WILAYAH_PATH = os.path.join(DATA, 'wilayah.json')

# ---- region berdasar digit pertama No SPBU ----
REGIONS = ['Sumbagut', 'Sumbagsel', 'JBB', 'JBT', 'Jatimbalinus',
           'Kalimantan', 'Sulawesi', 'Maluku Papua']
DIGIT = {str(i + 1): name for i, name in enumerate(REGIONS)}


def region_for(spbu_code):
    """Region sheet dari digit pertama No SPBU."""
    return DIGIT.get(spbu_code[0] if spbu_code else '', REGIONS[0])


# ---- rentang halaman yang memuat kode lahan harga 0 ----
def all_pages():
    """Semua halaman annual + monthly untuk harga 0 (10.629 lahan)."""
    return [(1, p) for p in range(1, 446)] + [(2, p) for p in range(1, 18)]


def jb_pages():
    """Halaman yang memuat kode Jakarta Barat (annual hlm 111-131, bulanan 5-7)."""
    return [(1, p) for p in range(111, 132)] + [(2, p) for p in range(5, 8)]


def price0_params(period, page):
    """Parameter query untuk list lahan dengan Harga Sewa = 0."""
    return {'TypePeriod': period, 'MinHarga': 0, 'MaxHarga': 0, 'Page': page}


# ---------------------------------------------------------------- item
def load_items():
    """Semua lahan harga 0 (dedup by landCode) dari cache JSONL."""
    import scrapekit as sk
    return sk.load_items_from_jsonl(CACHE)


# ---------------------------------------------------------------- wilayah
def load_wilayah():
    """Map landCode (kode bersih) -> Wilayah. Key dinormalisasi tanpa spasi."""
    w = {}
    if os.path.exists(WILAYAH_PATH):
        raw = json.load(open(WILAYAH_PATH))
        for k, v in raw.items():
            w[re.sub(r'\s+', '', k)] = v
    return w


# ---------------------------------------------------------------- klasifikasi Jakarta Barat
LAT = (-6.27, -6.06)
LON = (106.68, 106.85)
WJ_KEC = ['cengkareng', 'grogol', 'kalideres', 'kebon jeruk', 'kembangan',
          'palmerah', 'taman sari', 'tambora']
WJ_STR = ['tomang', 'daan mogot', 'roxy', 'glodok', 'pinangsia', 'jembatan besi',
          'meruya', 'kedoya', 'duri kosambi', 'rawa buaya', 'tegal alur', 'semanan',
          'kota bambu', 'slipi', 's. parman', 'kamal raya', 'peta barat', 'ulujami',
          'srengseng', 'joglo', 'kemanggisan', 'cendrawasih']
NOT_WB = ['jakarta pusat', 'jakarta selatan', 'jakarta timur', 'jakarta utara',
          'kodya tangerang', 'tangerang', 'soekarno hatta', 'tendean',
          'pasar minggu', 'pondok pinang', 'petir,', 'ciledug', 'jombang']
POS_LABEL = {'IN': 'Ruangan/Bangunan', 'OU': 'Tanah Kosong'}


def in_box(it):
    try:
        la = float(it.get('latitude') or 0); lo = float(it.get('longitude') or 0)
    except Exception:
        return False
    return la != 0 and LAT[0] <= la <= LAT[1] and LON[0] <= lo <= LON[1]


def is_jakarta_barat(it):
    a = (it.get('address') or '').lower()
    if any(n in a for n in NOT_WB):
        return False
    if 'jakarta barat' in a or 'jakbar' in a:
        return True
    if in_box(it) and (any(k in a for k in WJ_KEC) or any(k in a for k in WJ_STR)):
        return True
    return False


def jb_rows(items):
    """Baris detail untuk Excel/CSV Jakarta Barat (sama seperti extract_lahan)."""
    rows = []
    for it in sorted(items, key=lambda x: x['landCode']):
        code = it['landCode']
        rows.append({
            'No SPBU': code.split('-')[0],
            'Kode Lahan': code,
            'Status Sewa': it.get('rentStatName'),
            'Harga Sewa': it.get('rentalPrice'),
            'Periode': it.get('rentalPeriodName'),
            'Posisi': f"{it.get('landPosition')} ({POS_LABEL.get(it.get('landPosition'), '')})",
            'Luas (m2)': it.get('landArea'),
            'Panjang (m)': it.get('landLength'),
            'Lebar (m)': it.get('landWidth'),
            'Daya Listrik (VA/W)': it.get('powerCapacity'),
            'Alamat': it.get('address'),
            'Lintang': it.get('latitude'),
            'Bujur': it.get('longitude'),
        })
    return rows
