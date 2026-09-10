#!/usr/bin/env python3
"""
download_all_photos.py — Unduh SEMUA foto lahan (Harga Sewa=0, seluruh Indonesia).

- Kumpulkan semua URL unik dari lahan.
- Simpan ke data/photos_all/<hash>.<ext> (dedup per URL).
- Tulis data/photo_map.json secara inkremental (url -> path relatif).
- Bisa dihentikan & dilanjutkan (skip file yang sudah ada).

Contoh: python3 download_all_photos.py
"""
import hashlib, json, os, sys, time
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import scrapekit as sk
from sources import brightspace as bs

PHOTO_DIR = os.path.join(bs.DATA, 'photos_all')
MAP = bs.PHOTO_MAP


def one(job):
    url, fname = job
    return sk.download_image(url, os.path.join(PHOTO_DIR, fname))


def main():
    os.makedirs(PHOTO_DIR, exist_ok=True)
    items = bs.load_items()
    urls = set()
    for it in items.values():
        urls.update(sk.image_urls(it))
    urls = sorted(urls)
    print(f"Lahan: {len(items)} | URL unik: {len(urls)}", flush=True)

    url_map = {}
    if os.path.exists(MAP):
        url_map = json.load(open(MAP))

    todo = [u for u in urls
            if u not in url_map or not os.path.exists(os.path.join(bs.DATA, url_map[u]))]
    print(f"Perlu diunduh: {len(todo)} | sudah: {len(urls)-len(todo)}", flush=True)

    lock = __import__('threading').Lock()
    done = 0

    def work(url):
        nonlocal done
        fname = hashlib.md5(url.encode()).hexdigest()[:20] + '.img'
        res = one((url, fname))
        with lock:
            if res[0]:
                base = os.path.basename(res[1])
                url_map[url] = 'photos_all/' + base
            done += 1
            if done % 50 == 0 or done == len(todo):
                json.dump(url_map, open(MAP, 'w'))
                print(f"  {done}/{len(todo)}", flush=True)
        return res[0]

    stats = {'ok': 0, 'fail': 0}
    with ThreadPoolExecutor(max_workers=10) as ex:
        for fut in as_completed([ex.submit(work, u) for u in todo]):
            stats['ok' if fut.result() else 'fail'] += 1
            time.sleep(0.05)

    json.dump(url_map, open(MAP, 'w'))
    n = len([1 for u in urls if u in url_map and os.path.exists(os.path.join(bs.DATA, url_map[u]))])
    print(f"SELESAI. ok={stats['ok']} fail={stats['fail']} | terunduh: {n}/{len(urls)}", flush=True)


if __name__ == '__main__':
    main()
