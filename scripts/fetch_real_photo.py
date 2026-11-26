#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Museum-Grade Open-License Photography Fetcher (Wikimedia Featured & Quality Images)
for morphdeck-3d v4.0.0.

Prioritizes juried 'Featured pictures on Wikimedia Commons' and 'Quality images'
while filtering out noisy tourist crowd snapshots, street clutter, and calibration spheres.
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

USER_AGENT = "MorphDeck3D/4.0 (https://github.com/7AK3SH1/morphdeck-3d)"
COMMONS_API = "https://commons.wikimedia.org/w/api.php"

CURATED_FALLBACKS = {
    "villa": "minimalist modern architecture pool reflection",
    "alps": "Matterhorn Riffelsee reflection sunrise",
    "museum": "Kunsthistorisches Museum Vienna gallery hall",
    "gallery": "Louvre Apollo Gallery ceiling architecture",
    "aviation": "Lockheed Electra vintage aircraft monochrome",
    "world_map": "Ortelius Typus Orbis Terrarum map",
    "studio": "Leica camera lens macro studio",
}

NOISY_TITLE_TOKENS = (
    "tourist",
    "tourists",
    "crowd",
    "visitors",
    "selfie",
    "protest",
    "demonstration",
    "rally",
    "parking_lot",
    "bus_stop",
    "construction_site",
    "hdri",
    "equirectangular",
    "calibration",
    "montage",
    "collage",
    "composite",
    "manipulation",
    "photoshop",
    "surreal",
    "fantasy",
    "fictional",
    "hoax",
    "cartoon",
    "illustration",
    "logo",
    "icon",
    "diagram",
    "coat_of_arms",
    "flag_of",
)


def _cache_dir() -> str:
    d = os.path.join(tempfile.gettempdir(), "morphdeck4_photo_cache")
    os.makedirs(d, exist_ok=True)
    return d


def _search_commons_tier(gsrsearch: str, min_width: int, max_width: int) -> list[tuple[float, int, str, str]]:
    params = urllib.parse.urlencode({
        "action": "query",
        "generator": "search",
        "gsrsearch": gsrsearch,
        "gsrnamespace": "6",
        "gsrlimit": "25",
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
    q_low = gsrsearch.lower()
    for p in pages:
        title = p.get("title", "")
        t_low = title.lower().replace(" ", "_")
        if any(bad in t_low and bad not in q_low for bad in NOISY_TITLE_TOKENS):
            continue
        infos = p.get("imageinfo") or []
        if not infos:
            continue
        info = infos[0]
        w = int(info.get("width", 0))
        h = int(info.get("height", 0))
        mime = (info.get("mime") or "").lower()
        if w < min_width or h < 720 or w <= int(h * 1.18):
            continue
        if any(bad_m in mime for bad_m in ("tiff", "svg", "gif", "pdf", "xcf")):
            continue
        aspect = w / max(1.0, float(h))
        if aspect > 2.65:
            continue
        dl_url = info.get("thumburl") or info.get("url")
        if dl_url:
            # Ideal 16:9 widescreen aspect = 1.778
            aspect_penalty = abs(aspect - 1.72)
            res_bonus = -min(w, 3840) / 4000.0
            candidates.append((aspect_penalty + res_bonus * 0.35, -w, dl_url, title))
    candidates.sort()
    return candidates


def fetch_wikimedia_photo(query: str, out_path: str, min_width: int = 1400, max_width: int = 2560) -> str:
    """Search Wikimedia Commons Featured & Quality collections first, download & save crisp RGB PNG."""
    key = hashlib.sha256(f"v4|{query}|{min_width}|{max_width}".encode("utf-8")).hexdigest()[:16]
    cached = os.path.join(_cache_dir(), f"{key}.png")
    if os.path.exists(cached) and os.path.getsize(cached) > 30000:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        Image.open(cached).convert("RGB").save(out_path, "PNG")
        return out_path

    words = [w for w in query.split() if w.strip()]
    base_queries = [query]
    if len(words) > 3:
        base_queries.append(" ".join(words[:3]))
    if len(words) > 2:
        base_queries.append(" ".join(words[:2]))

    search_plan = []
    neg = '-incategory:"Photo_montages" -incategory:"Digital_art" -montage -collage'
    for bq in base_queries:
        search_plan.append(f'incategory:"Featured_pictures_on_Wikimedia_Commons" {bq} {neg} filetype:bitmap')
        search_plan.append(f'incategory:"Quality_images" {bq} {neg} filetype:bitmap')
    for bq in base_queries:
        search_plan.append(f"{bq} {neg} filetype:bitmap")

    candidates = []
    for gsr in search_plan:
        try:
            hits = _search_commons_tier(gsr, min_width=min_width, max_width=max_width)
        except Exception:
            hits = []
        if hits:
            candidates = hits
            break

    if not candidates and min_width > 1100:
        for bq in base_queries:
            try:
                hits = _search_commons_tier(f"{bq} filetype:bitmap", min_width=1024, max_width=max_width)
            except Exception:
                hits = []
            if hits:
                candidates = hits
                break

    if not candidates:
        raise RuntimeError(f"No suitable landscape photograph found on Wikimedia Commons for '{query}'")

    last_err = None
    for _, _, dl_url, _title in candidates[:5]:
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
    key = hashlib.sha256(f"v4|{url}".encode("utf-8")).hexdigest()[:16]
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
    parser = argparse.ArgumentParser(description="Fetch juried Featured/Quality landscape photos via Wikimedia Commons.")
    parser.add_argument("--query", default="", help="Search query (e.g. 'Matterhorn Riffelsee sunrise')")
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
    print(f"Fetched juried photograph ({im.width}x{im.height}) -> {args.out}")


if __name__ == "__main__":
    main()
