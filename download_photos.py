#!/usr/bin/env python3
"""
download_photos.py — Unduh semua foto lahan (Harga Sewa=0) di Jakarta Barat.

Hasil: deliverables/data/photos/<kode-lahan>/img-<n>.jpg

Contoh: python3 download_photos.py
"""
import os, sys
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import scrapekit as sk
from sources import brightspace as bs

PHOTO_DIR = os.path.join(bs.DATA, 'photos')


def one(job):
    url, path = job
    return sk.download_image(url, path)


def main():
    items = {c: it for c, it in bs.load_items().items() if bs.is_jakarta_barat(it)}
    print(f"Lahan: {len(items)}")

    jobs = []
    seen = set()
    for code, it in sorted(items.items()):
        folder = os.path.join(PHOTO_DIR, code)
        os.makedirs(folder, exist_ok=True)
        for i, url in enumerate(sk.image_urls(it), 1):
            jobs.append((url, os.path.join(folder, f'img-{i}')))
            seen.add(url)
    print(f"Total request foto: {len(jobs)}, URL unik: {len(seen)}")

    ok = fail = 0
    bytes_ = 0
    with ThreadPoolExecutor(max_workers=8) as ex:
        for fut in as_completed([ex.submit(one, j) for j in jobs]):
            good, final, n = fut.result()
            if good:
                ok += 1; bytes_ += n
            else:
                fail += 1
                sys.stderr.write(f"  GAGAL: {final}\n")
    print(f"Selesai: {ok} foto terunduh, {fail} gagal, total {bytes_/1024/1024:.1f} MB")
    print(f"Folder foto: {os.path.abspath(PHOTO_DIR)}")


if __name__ == '__main__':
    main()
