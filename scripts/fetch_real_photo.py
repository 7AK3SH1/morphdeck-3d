#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Real Open-License Photography Fetcher (Wikimedia Commons API + Cache) for morphdeck-3d.

Fetches high-resolution CC0 / CC-BY / Public-Domain architectural, museum, historical,
alpine, and gallery photographs — never HDRI preview calibration spheres or checkerboard floors.
"""

import argparse
import hashlib
import io
import json
import os
import tempfile
import urllib.parse
import urllib.request
from PIL import Image

USER_AGENT = "MorphDeck3D/3.0 (https://github.com/7AK3SH1/morphdeck-3d)"
COMMONS_API = "https://commons.wikimedia.org/w/api.php"

CURATED_FALLBACKS = {
    "villa": "modern glass luxury villa architecture",
    "alps": "Matterhorn Zermatt glacier panorama",
    "museum": "Louvre gallery hall painting room",
    "gallery": "Uffizi gallery Florence hall",
    "aviation": "historic propeller aircraft monochrome",
    "world_map": "vintage world map parchment cartography",
    "studio": "camera lens studio moody dark",
}


def _cache_dir() -> str:
    d = os.path.join(tempfile.gettempdir(), "morphdeck3d_photo_cache")
    os.makedirs(d, exist_ok=True)
    return d


def _search_commons(query: str, min_width: int, max_width: int) -> list[tuple[float, int, str, str]]:
    params = urllib.parse.urlencode({
        "action": "query",
        "generator": "search",
        "gsrsearch": f"{query} filetype:bitmap",
        "gsrnamespace": "6",
        "gsrlimit": "20",
        "prop": "imageinfo",
        "iiprop": "url|size|mime",
        "iiurlwidth": str(max_width),
        "format": "json",
    })
    url = f"{COMMONS_API}?{params}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=25) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    pages = list(data.get("query", {}).get("pages", {}).values())
    candidates = []
    for p in pages:
        infos = p.get("imageinfo") or []
        if not infos:
            continue
        info = infos[0]
        w = int(info.get("width", 0))
        h = int(info.get("height", 0))
        mime = (info.get("mime") or "").lower()
        if w < min_width or h < 650 or w <= h:
            continue
        if "tiff" in mime or "svg" in mime or "gif" in mime:
            continue
        dl_url = info.get("thumburl") or info.get("url")
        if dl_url:
            candidates.append((abs((w / max(1, h)) - 1.5), -w, dl_url, p.get("title", "")))
    return candidates


def fetch_wikimedia_photo(query: str, out_path: str, min_width: int = 1200, max_width: int = 2400) -> str:
    """Search Wikimedia Commons for a landscape open-license photo, download & save as PNG."""
    key = hashlib.sha256(f"{query}|{min_width}|{max_width}".encode("utf-8")).hexdigest()[:16]
    cached = os.path.join(_cache_dir(), f"{key}.png")
    if os.path.exists(cached) and os.path.getsize(cached) > 25000:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        Image.open(cached).convert("RGB").save(out_path, "PNG")
        return out_path

    words = [w for w in query.split() if w.strip()]
    sub_queries = [query]
    if len(words) > 3:
        sub_queries.append(" ".join(words[:3]))
    if len(words) > 2:
        sub_queries.append(" ".join(words[:2]))

    candidates = []
    for q_try in sub_queries:
        candidates = _search_commons(q_try, min_width=min_width, max_width=max_width)
        if candidates:
            break

    if not candidates:
        raise RuntimeError(f"No suitable landscape bitmap found on Wikimedia Commons for '{query}'")

    candidates.sort()
    last_err = None
    for _, _, dl_url, _title in candidates[:4]:
        try:
            r2 = urllib.request.Request(dl_url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(r2, timeout=45) as rimg:
                raw = rimg.read()
            im = Image.open(io.BytesIO(raw)).convert("RGB")
            im.save(cached, "PNG")
            os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
            im.save(out_path, "PNG")
            return out_path
        except Exception as exc:
            last_err = exc
            continue

    raise RuntimeError(f"Failed downloading Wikimedia candidates for '{query}': {last_err}")


def fetch_url_photo(url: str, out_path: str) -> str:
    key = hashlib.sha256(url.encode("utf-8")).hexdigest()[:16]
    cached = os.path.join(_cache_dir(), f"url_{key}.png")
    if os.path.exists(cached) and os.path.getsize(cached) > 15000:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        Image.open(cached).convert("RGB").save(out_path, "PNG")
        return out_path

    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=45) as resp:
        raw = resp.read()
    im = Image.open(io.BytesIO(raw)).convert("RGB")
    im.save(cached, "PNG")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    im.save(out_path, "PNG")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch real open-license landscape photos via Wikimedia Commons.")
    parser.add_argument("--query", default="", help="Search query (e.g. 'modern luxury villa glass')")
    parser.add_argument("--preset", choices=sorted(CURATED_FALLBACKS.keys()), help="Curated architectural/editorial preset")
    parser.add_argument("--url", default="", help="Direct image URL")
    parser.add_argument("--out", required=True, help="Output PNG path")
    args = parser.parse_args()

    if args.url:
        fetch_url_photo(args.url, args.out)
    else:
        q = CURATED_FALLBACKS[args.preset] if args.preset else args.query
        if not q:
            parser.error("Provide --query, --preset, or --url")
        fetch_wikimedia_photo(q, args.out)
    im = Image.open(args.out)
    print(f"Fetched real photograph ({im.width}x{im.height}) -> {args.out}")


if __name__ == "__main__":
    main()
