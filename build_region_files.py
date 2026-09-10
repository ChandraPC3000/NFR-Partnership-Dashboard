#!/usr/bin/env python3
"""
build_region_files.py — Pisah lahan Harga Sewa=0 menjadi 8 FILE Excel per region.

Mengatasi limit rendering gambar Excel (~7.500 gambar/workbook):
setiap file hanya berisi ≤2.900 gambar, sehingga SEMUA foto tampil.

Kolom tiap file: No | No SPBU | Kode Lahan | Wilayah | Status Sewa | Foto

Logika situs (endpoint, filter, wilayah, region) ada di sources/brightspace.py,
helper umum (retry, thumbnail, Excel) di scrapekit.py.

Contoh: python3 build_region_files.py
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import scrapekit as sk
from sources import brightspace as bs

THUMB_H = 84
COLS = ['No', 'No SPBU', 'Kode Lahan', 'Wilayah', 'Status Sewa', 'Foto']
OUT_DIR = os.path.join(bs.DATA, 'by-region')


def build_one(name, rows, url_map):
    wb = __import__('openpyxl').Workbook()
    ws = wb.active
    ws.title = name
    sk.init_ws(ws, COLS)
    no_photo = 0
    for i, (code, spbu, wil, status, url) in enumerate(rows, 1):
        r = i + 1
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=spbu)
        ws.cell(row=r, column=3, value=code)
        ws.cell(row=r, column=4, value=wil)
        ws.cell(row=r, column=5, value=status)
        rel = url_map.get(url)
        if rel and os.path.exists(os.path.join(bs.DATA, rel)):
            if not sk.embed_photo(ws, f'F{r}', os.path.join(bs.DATA, rel), THUMB_H):
                no_photo += 1
        else:
            no_photo += 1
        ws.row_dimensions[r].height = int(THUMB_H * 72 / 96) + 4
    sk.style_ws(ws, [('A', 7), ('B', 12), ('C', 22), ('D', 24), ('E', 13), ('F', 12)],
                filter_ref=f"A1:F{len(rows)+1}")
    out = os.path.join(OUT_DIR, f'{name}.xlsx')
    wb.save(out)
    mb = os.path.getsize(out) / 1024 / 1024
    print(f"  {name}: {len(rows)} baris, {len(rows)-no_photo} foto, {mb:.1f} MB, tanpa foto {no_photo}",
          flush=True)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    items = bs.load_items()
    wilayah = bs.load_wilayah()
    url_map = json.load(open(bs.PHOTO_MAP)) if os.path.exists(bs.PHOTO_MAP) else {}

    groups = {name: [] for name in bs.REGIONS}
    for c, it in items.items():
        clean = re.sub(r'\s+', '', c)
        spbu = clean.split('-')[0]
        region = bs.region_for(spbu)
        groups[region].append((clean, spbu, wilayah.get(clean, ''),
                               it.get('rentStatName') or '',
                               it.get('imageURL') or ''))
    for name in bs.REGIONS:
        groups[name].sort()
        build_one(name, groups[name], url_map)

    print(f"SELESAI -> {OUT_DIR}")


if __name__ == '__main__':
    main()
