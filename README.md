# NFR-Apps — Tools Analisis & Proposal NFR Lahan (Pertamina Patra Niaga)

Alat bantu untuk mengolah data lahan **NFR (Non-Fuel Retail)** di SPBU Pertamina:

- Menarik data lahan dari platform **Brightspace** (`brightspace.pertamina.com`)
- Menggabungkan dengan **Template** (enrichment: SES, Red Ocean, Blue Ocean, Road Type, harga, dll.)
- Memfilter sesuai kebutuhan team & menyiapkan **proposal** untuk prospective tenant (baru / expand)
- **App lokal (Streamlit)**: upload file → filter → lihat tabel/peta/chart → export Excel

---

## Struktur

```
tools/
  scrapekit.py                 # helper umum: HTTP retry, cache JSONL, thumbnail, tulis Excel
  sources/brightspace.py       # adapter khusus brightspace.pertamina.com (endpoint, filter, wilayah)
  prepare_proposal.py          # CLI: export master → enrich → filter → Excel (Data + Charts + Peta)
  app_nfr.py                   # app Streamlit lokal (upload, filter, sort, peta, chart, export)
  build_all_excel.py           # Excel seluruh lahan Harga Sewa=0 (No SPBU, Kode, Status)
  build_region_files.py        # pisah per region (8 file, hindari limit gambar Excel)
  build_region_sheets.py       # 1 workbook, 8 sheet per region
  download_all_photos.py       # unduh semua foto lahan → photo_map.json
  download_photos.py           # unduh foto lahan Jakarta Barat
  embed_all_photos.py          # Excel + thumbnail foto (semua lahan)
  embed_photos.py              # Excel + thumbnail foto (Jakarta Barat)
  extract_lahan.py             # ekstrak lahan Jakarta Barat (harga 0)
```

## Install

```bash
pip install -r requirements.txt
```

## Jalankan app (Streamlit)

```bash
cd tools
streamlit run app_nfr.py
```

Atau double-click **`Buka NFR App.command`**.

App ini untuk **internal**: upload file master lahan + template + PIC → filter → lihat peta/chart → **Export Excel (Data + Charts + Peta)**.

## CLI proposal (per request tenant)

```bash
python3 tools/prepare_proposal.py \
  --master "NFR Master Lahan <tanggal>.xlsx" \
  --template "Available Space NFR at SPBU - V1 Template.xlsx" \
  --pic "PICSPBU.xlsx" \
  --region JBB --kota "Jakarta Barat" --min-luas 20 --max-harga 100000000 \
  --usaha coffee \
  --output "Proposal_ClientA"
```

Filter opsional: `--region` (8 region), `--kota`, `--min/max-luas`, `--min/max-harga`, `--usaha` (cocokkan Red/Blue Ocean), `--status` (default Tersedia).

## Catatan

- **Data hasil (folder `data/`) TIDAK di-commit** (lihat `.gitignore`). File input/data berada di lokal, bukan di repo.
- Kode ini untuk **penggunaan internal Pertamina Patra Niaga** — repo PRIVATE.
- Bagian web scraping: logika spesifik platform ada di `tools/sources/brightspace.py`; bagian umum (retry, pagination, cache, Excel) di `tools/scrapekit.py`, sehingga mudah menambah sumber data lain.
