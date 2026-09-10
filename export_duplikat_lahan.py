#!/usr/bin/env python3
"""
export_duplikat_lahan.py
Deteksi dan export duplikat Kode Lahan dari NFR Master Lahan ke Excel.
"""
import sys
import datetime
from collections import Counter
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INPUT  = Path('/Users/fwzmwrdy/Documents/Work/Pertamina Patra Niaga/NFR Management/NFR Master Lahan 28-August-2026.xlsx')
OUTPUT = Path('/Users/fwzmwrdy/Documents/Work/Pertamina Patra Niaga/NFR Management') / \
         f'Duplikat_Kode_Lahan_{datetime.date.today().strftime("%Y-%m-%d")}.xlsx'

# ── styling ──────────────────────────────────────────────────────────────────
RED_HDR   = 'C00000'
BLUE_HDR  = '0070C0'
ORANGE    = 'F4B942'
YELLOW    = 'FFF2CC'
WHITE     = 'FFFFFF'
LIGHT_RED = 'FCE4D6'
LIGHT_BLU = 'DEEAF1'

thin = Side(border_style='thin', color='BFBFBF')
bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

def hdr_style(cell, hex_color):
    cell.fill   = PatternFill('solid', fgColor=hex_color)
    cell.font   = Font(color=WHITE, bold=True, size=11, name='Calibri')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = bdr

def dat_style(cell, hex_bg=WHITE, bold=False, center=False):
    cell.fill   = PatternFill('solid', fgColor=hex_bg)
    cell.font   = Font(size=11, name='Calibri', bold=bold)
    cell.alignment = Alignment(horizontal='center' if center else 'left',
                               vertical='center', wrap_text=False)
    cell.border = bdr

def set_col_width(ws, col_idx, width):
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# ── load data ─────────────────────────────────────────────────────────────────
print('Membaca Master Lahan...')
wb_in = openpyxl.load_workbook(INPUT, read_only=True, data_only=True)
ws_in = wb_in['NFR Master Lahan']
rows  = list(ws_in.iter_rows(values_only=True))
header, data = rows[0], rows[1:]

ki = {h: i for i, h in enumerate(header)}
KODE   = ki['Kode Lahan']
SPBU   = ki['No SPBU']
LUAS   = ki['Luas (M2)']
HARGA  = ki['Harga (Rp)']
TENANT = ki['Tenant']
TGL_M  = ki['Tanggal Mulai Sewa']
TGL_B  = ki['Tanggal Berakhir Sewa']
SEWA   = ki['Harga Sewa (Rp)']
STATUS = ki['Status Sewa']
APPROV = ki['Status Approval']

# ── detect duplicates ─────────────────────────────────────────────────────────
kode_counts = Counter(r[KODE] for r in data)
dup_kodes   = {k for k, v in kode_counts.items() if v > 1}

dup_rows = [r for r in data if r[KODE] in dup_kodes]

# classify each group
def classify_group(group):
    """Identik = semua baris persis sama (luas+harga+status). Beda = ada perbedaan nilai."""
    sig = set((r[LUAS], r[HARGA], r[STATUS]) for r in group)
    if len(sig) == 1:
        return 'Identik'
    return 'Beda Nilai'

groups = {}
for r in dup_rows:
    groups.setdefault(r[KODE], []).append(r)

print(f'Kode Lahan duplikat: {len(groups)}')
print(f'Total baris terdampak: {len(dup_rows)}')

identik_count = sum(1 for g in groups.values() if classify_group(g) == 'Identik')
beda_count    = len(groups) - identik_count
print(f'  Tipe Identik (copy persis): {identik_count}')
print(f'  Tipe Beda Nilai (reinput):  {beda_count}')

# ── build Excel ───────────────────────────────────────────────────────────────
wb_out = openpyxl.Workbook()

# ── Sheet 1: Summary ──────────────────────────────────────────────────────────
ws1 = wb_out.active
ws1.title = 'Summary'
ws1.sheet_view.showGridLines = False

# title
ws1.merge_cells('A1:F1')
t = ws1['A1']
t.value = 'Laporan Deteksi Duplikat Kode Lahan — NFR Master Lahan'
t.font  = Font(bold=True, size=14, name='Calibri', color='1A1A1A')
t.alignment = Alignment(horizontal='left', vertical='center')
ws1.row_dimensions[1].height = 28

ws1.merge_cells('A2:F2')
sub = ws1['A2']
sub.value = f'Tanggal analisis: {datetime.date.today().strftime("%d %B %Y")}  |  Sumber: {INPUT.name}'
sub.font  = Font(size=10, name='Calibri', color='666666')
sub.alignment = Alignment(horizontal='left', vertical='center')
ws1.row_dimensions[2].height = 18

ws1.row_dimensions[3].height = 10

# KPI cards
kpi = [
    ('Total Lahan di Master',        len(data),       BLUE_HDR),
    ('Kode Lahan Duplikat',          len(groups),     RED_HDR),
    ('Total Baris Terdampak',        len(dup_rows),   RED_HDR),
    ('Tipe Identik (copy persis)',   identik_count,   ORANGE),
    ('Tipe Beda Nilai (reinput)',     beda_count,      ORANGE),
    ('Kode Lahan Unik (bersih)',     len(kode_counts) - len(groups), '5A8A00'),
]
ws1.row_dimensions[4].height = 20
ws1.row_dimensions[5].height = 40
ws1.row_dimensions[6].height = 28
for ci, (label, val, color) in enumerate(kpi, 1):
    lc = ws1.cell(row=5, column=ci, value=label)
    lc.fill      = PatternFill('solid', fgColor=color)
    lc.font      = Font(color=WHITE, bold=True, size=9, name='Calibri')
    lc.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)
    lc.border    = bdr
    vc = ws1.cell(row=6, column=ci, value=val)
    vc.fill      = PatternFill('solid', fgColor='F4F6F9')
    vc.font      = Font(bold=True, size=16, name='Calibri')
    vc.alignment = Alignment(horizontal='center', vertical='center')
    vc.border    = bdr
    ws1.column_dimensions[get_column_letter(ci)].width = 22

ws1.row_dimensions[7].height = 10

# breakdown per Tipe Duplikat
ws1.cell(row=8, column=1, value='Breakdown per Tipe').font = Font(bold=True, size=11, name='Calibri')
ws1.row_dimensions[8].height = 20
for ci, h in enumerate(['Tipe Duplikat', 'Jumlah Kode', 'Total Baris', 'Keterangan'], 1):
    c = ws1.cell(row=9, column=ci, value=h)
    hdr_style(c, BLUE_HDR)
ws1.row_dimensions[9].height = 22

breakdown = [
    ('Identik', identik_count,
     sum(len(g) for k, g in groups.items() if classify_group(g) == 'Identik'),
     'Semua baris persis sama — kemungkinan copy dari sistem lama'),
    ('Beda Nilai', beda_count,
     sum(len(g) for k, g in groups.items() if classify_group(g) == 'Beda Nilai'),
     'Kode sama tapi Luas/Harga berbeda — kemungkinan reinput atau update tidak hapus entri lama'),
]
for ri, (tipe, jml, brs, ket) in enumerate(breakdown, 10):
    ws1.cell(row=ri, column=1, value=tipe).font   = Font(size=11, name='Calibri', bold=True)
    ws1.cell(row=ri, column=2, value=jml).font    = Font(size=11, name='Calibri')
    ws1.cell(row=ri, column=3, value=brs).font    = Font(size=11, name='Calibri')
    ws1.cell(row=ri, column=4, value=ket).font    = Font(size=10, name='Calibri', color='555555')
    for ci in range(1, 5):
        c = ws1.cell(row=ri, column=ci)
        c.fill      = PatternFill('solid', fgColor=LIGHT_RED if ri % 2 == 0 else WHITE)
        c.alignment = Alignment(vertical='center')
        c.border    = bdr
    ws1.row_dimensions[ri].height = 18

ws1.column_dimensions['D'].width = 60

# ── Sheet 2: Detail Duplikat ──────────────────────────────────────────────────
ws2 = wb_out.create_sheet('Detail Duplikat')
ws2.sheet_view.showGridLines = False

out_header = ['No', 'Kode Lahan', 'No SPBU', 'Luas (M2)', 'Harga (Rp)',
              'Tenant', 'Tgl Mulai Sewa', 'Tgl Berakhir Sewa',
              'Harga Sewa (Rp)', 'Status Sewa', 'Status Approval',
              'Tipe Duplikat', 'Jumlah Kemunculan', 'Catatan']

col_widths = [5, 22, 14, 12, 18, 24, 16, 18, 18, 14, 16, 16, 18, 40]
for ci, (h, w) in enumerate(zip(out_header, col_widths), 1):
    c = ws2.cell(row=1, column=ci, value=h)
    hdr_style(c, RED_HDR)
    ws2.column_dimensions[get_column_letter(ci)].width = w
ws2.row_dimensions[1].height = 28
ws2.freeze_panes = 'A2'
ws2.auto_filter.ref = f'A1:{get_column_letter(len(out_header))}1'

row_num = 2
no = 1
for kode in sorted(groups.keys()):
    group   = groups[kode]
    tipe    = classify_group(group)
    count   = len(group)
    bg      = LIGHT_RED if tipe == 'Identik' else YELLOW

    for gi, r in enumerate(group):
        # catatan per baris
        if tipe == 'Identik':
            catatan = 'KEMUNGKINAN DUPLIKAT MURNI — perlu verifikasi & hapus salah satu'
        else:
            luas_vals = list(set(x[LUAS] for x in group))
            harga_vals = list(set(x[HARGA] for x in group))
            catatan = f'Luas bervariasi: {sorted(luas_vals)} | Harga bervariasi: {sorted(harga_vals)}'

        vals = [
            no if gi == 0 else '',
            r[KODE], r[SPBU], r[LUAS], r[HARGA],
            r[TENANT], r[TGL_M], r[TGL_B], r[SEWA],
            r[STATUS], r[APPROV],
            tipe if gi == 0 else '',
            count if gi == 0 else '',
            catatan if gi == 0 else '',
        ]
        for ci, val in enumerate(vals, 1):
            c = ws2.cell(row=row_num, column=ci, value=val)
            dat_style(c, hex_bg=bg,
                      bold=(ci in (2, 12) and gi == 0),
                      center=(ci in (1, 3, 4, 5, 9, 12, 13)))
            if ci in (5, 9) and isinstance(val, (int, float)) and val:
                c.number_format = '#,##0'
        ws2.row_dimensions[row_num].height = 15
        row_num += 1

    # separator row between groups
    for ci in range(1, len(out_header) + 1):
        c = ws2.cell(row=row_num, column=ci, value='')
        c.fill = PatternFill('solid', fgColor='E8E8E8')
        c.border = bdr
    ws2.row_dimensions[row_num].height = 4
    row_num += 1
    no += 1

# ── Sheet 3: Identik saja (paling jelas untuk dihapus) ───────────────────────
ws3 = wb_out.create_sheet('Identik — Prioritas Hapus')
ws3.sheet_view.showGridLines = False

for ci, (h, w) in enumerate(zip(out_header, col_widths), 1):
    c = ws3.cell(row=1, column=ci, value=h)
    hdr_style(c, '5A8A00')
    ws3.column_dimensions[get_column_letter(ci)].width = w
ws3.row_dimensions[1].height = 28
ws3.freeze_panes = 'A2'
ws3.auto_filter.ref = f'A1:{get_column_letter(len(out_header))}1'

row_num3 = 2
no3 = 1
for kode in sorted(groups.keys()):
    group = groups[kode]
    if classify_group(group) != 'Identik':
        continue
    count = len(group)
    for gi, r in enumerate(group):
        vals = [
            no3 if gi == 0 else '',
            r[KODE], r[SPBU], r[LUAS], r[HARGA],
            r[TENANT], r[TGL_M], r[TGL_B], r[SEWA],
            r[STATUS], r[APPROV],
            'Identik' if gi == 0 else '',
            count if gi == 0 else '',
            'Simpan 1 baris, hapus sisanya' if gi == 0 else '',
        ]
        for ci, val in enumerate(vals, 1):
            c = ws3.cell(row=row_num3, column=ci, value=val)
            dat_style(c, hex_bg=LIGHT_RED if gi > 0 else WHITE,
                      bold=(gi == 0 and ci == 2),
                      center=(ci in (1, 3, 4, 5, 9, 12, 13)))
            if ci in (5, 9) and isinstance(val, (int, float)) and val:
                c.number_format = '#,##0'
        ws3.row_dimensions[row_num3].height = 15
        row_num3 += 1
    row_num3 += 1
    no3 += 1

# ── save ──────────────────────────────────────────────────────────────────────
wb_out.save(OUTPUT)
print(f'\nFile tersimpan: {OUTPUT}')
