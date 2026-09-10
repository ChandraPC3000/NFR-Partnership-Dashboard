#!/usr/bin/env python3
"""
extract_lahan.py — Ambil data lahan "Harga Sewa = 0" di Jakarta Barat
secara otomatis dari brightspace.pertamina.com (tanpa screenshot manual).

Cara kerja:
  1. Ambil hanya halaman yang mengandung kode Jakarta Barat
     (Tahunan: hlm 111-131, Bulanan: hlm 5-7) dengan filter harga 0-0.
  2. Dedup per kode lahan.
  3. Saring lokasi Jakarta Barat (alamat + koordinat).
  4. Tulis CSV + Excel ke deliverables/data/.

Note: filter kota (IdKota) di sisi server TIDAK berfungsi untuk lahan,
jadi kami mengandalkan rentang halaman + filter alamat/koordinat.

Contoh: python3 extract_lahan.py
"""
import csv, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import scrapekit as sk
from sources import brightspace as bs

COLS = ['No SPBU', 'Kode Lahan', 'Status Sewa', 'Harga Sewa', 'Periode', 'Posisi',
        'Luas (m2)', 'Panjang (m)', 'Lebar (m)', 'Daya Listrik (VA/W)',
        'Alamat', 'Lintang', 'Bujur']
WIDTHS = [10, 20, 11, 11, 9, 22, 10, 11, 10, 16, 55, 10, 10]


def collect_jb():
    """Lahan Jakarta Barat harga 0 (dedup), diambil dari halaman JB saja."""
    raw = sk.fetch_items_pages(bs.LAHAN_ITEMS_URL,
                               [bs.price0_params(*j) for j in bs.jb_pages()])
    seen = {}
    for it in raw:
        seen[it['landCode']] = it
    return [it for it in seen.values() if bs.is_jakarta_barat(it)]


def write_outputs(rows, out_dir):
    csv_path = os.path.join(out_dir, 'jakarta-barat-hargasewa-0-full.csv')
    with open(csv_path, 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=COLS)
        w.writeheader()
        for r in rows:
            w.writerow(r)

    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Lahan Harga 0 Jakarta Barat'
    sk.init_ws(ws, COLS)
    for ri, r in enumerate(rows, 2):
        for ci, c in enumerate(COLS, 1):
            ws.cell(row=ri, column=ci, value=r[c])
    sk.style_ws(ws, list(zip('ABCDEFGHIJKLM', WIDTHS)))
    xlsx_path = os.path.join(out_dir, 'jakarta-barat-hargasewa-0-full.xlsx')
    wb.save(xlsx_path)
    return csv_path, xlsx_path


def main():
    out_dir = bs.DATA
    os.makedirs(out_dir, exist_ok=True)
    print("Mengambil data halaman Jakarta Barat ...")
    jb = collect_jb()
    rows = bs.jb_rows(jb)
    csv_p, xlsx_p = write_outputs(rows, out_dir)
    from collections import Counter
    print(f"Total lahan Harga Sewa=0 di Jakarta Barat: {len(rows)}")
    print("Status:", dict(Counter(r['Status Sewa'] for r in rows)))
    print("Periode:", dict(Counter(r['Periode'] for r in rows)))
    print("No SPBU unik:", len({r['No SPBU'] for r in rows}))
    print(f"CSV : {csv_p}")
    print(f"XLSX: {xlsx_p}")


if __name__ == '__main__':
    main()
