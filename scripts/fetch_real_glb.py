#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Real High-Poly CC0 / Open-License 3D (.glb) Resolver & glTF-to-GLB Packer
for morphdeck-3d.

Resolves studio-grade 3D assets (never toy procedural primitives) from:
  1. Khronos glTF-Sample-Assets (high-poly PBR studio .glb library)
  2. Poly Haven CC0 3D Models API (500+ scanned photogrammetry assets -> packed .glb)
  3. NASA 3D Resources (220+ authentic spacecraft/rover/satellite .glb assets)
"""

import argparse
import json
import os
import struct
import tempfile
import time
import urllib.parse
import urllib.request

USER_AGENT = "MorphDeck3D/3.0 (https://github.com/7AK3SH1/morphdeck-3d)"

KHRONOS_INDEX_URL = "https://raw.githubusercontent.com/KhronosGroup/glTF-Sample-Assets/main/Models/model-index.json"
KHRONOS_GLB_URL = "https://raw.githubusercontent.com/KhronosGroup/glTF-Sample-Assets/main/Models/{name}/glTF-Binary/{name}.glb"
POLYHAVEN_ASSETS_URL = "https://api.polyhaven.com/assets?t=models"
POLYHAVEN_FILES_URL = "https://api.polyhaven.com/files/{asset_id}"
NASA_LIST_URL = "https://api.github.com/repos/nasa/NASA-3D-Resources/contents/3D%20Models"
NASA_DIR_URL = "https://api.github.com/repos/nasa/NASA-3D-Resources/contents/3D%20Models/{folder}"

DOMAIN_KHRONOS_MAP = [
    (("car", "auto", "vehicle", "mobility", "race", "ev", "transport", "porsche", "ferrari"), "CarConcept"),
    (("camera", "photo", "lens", "cinema", "film", "media", "optics", "leica", "luxury"), "AntiqueCamera"),
    (("sneaker", "shoe", "footwear", "fashion", "retail", "apparel", "drop"), "MaterialsVariantsShoe"),
    (("helmet", "cyber", "defense", "gaming", "armor", "vr", "scifi"), "DamagedHelmet"),
    (("sofa", "couch", "interior", "furniture", "living", "hospitality"), "GlamVelvetSofa"),
    (("chair", "seat", "office", "studio"), "SheenChair"),
    (("chess", "strategy", "executive", "consulting", "board", "game"), "ABeautifulGame"),
    (("sunglass", "eyewear", "glasses"), "SunglassesKhronos"),
    (("plant", "botanical", "eco", "green", "climate", "nature", "bio", "farm"), "DiffuseTransmissionPlant"),
    (("bottle", "beverage", "consumer", "drink", "packaging"), "WaterBottle"),
    (("biotech", "dna", "genomic", "amber", "pharma", "lab", "research"), "MosquitoInAmber"),
    (("medical", "anatomy", "health", "skull", "neuro", "clinical"), "ScatteringSkull"),
    (("city", "urban", "architecture", "building", "smart city", "real estate"), "VirtualCity"),
    (("dragon", "sculpture", "museum", "art", "crystal", "gallery"), "DragonAttenuation"),
    (("lantern", "lamp", "light", "energy"), "Lantern"),
    (("audio", "music", "sound", "speaker", "podcast"), "BoomBox"),
]


def _http_get(url: str, timeout: int = 35) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _http_json(url: str, timeout: int = 25):
    return json.loads(_http_get(url, timeout=timeout).decode("utf-8"))


def list_khronos_models() -> list[str]:
    idx = _http_json(KHRONOS_INDEX_URL)
    skip = (
        "Test", "Compare", "Box", "Triangle", "Cube", "Normal", "Texture",
        "Interpolation", "Unlit", "Vertex", "Sparse", "Recursive", "Simple",
        "Environment", "Coordinates", "Orientation",
    )
    return [
        m["name"]
        for m in idx
        if "glTF-Binary" in m.get("variants", {}) and not any(t in m["name"] for t in skip)
    ]


def fetch_khronos_glb(model_name: str, out_glb: str) -> str:
    url = KHRONOS_GLB_URL.format(name=urllib.parse.quote(model_name))
    data = _http_get(url, timeout=90)
    if data[:4] != b"glTF":
        raise ValueError(f"Invalid GLB downloaded from {url}")
    os.makedirs(os.path.dirname(os.path.abspath(out_glb)), exist_ok=True)
    with open(out_glb, "wb") as f:
        f.write(data)
    return out_glb


def _polyhaven_catalog(cache_ttl_sec: int = 86400) -> dict:
    cache_path = os.path.join(tempfile.gettempdir(), "morphdeck_polyhaven_models.json")
    if os.path.exists(cache_path) and (time.time() - os.path.getmtime(cache_path)) < cache_ttl_sec:
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    cat = _http_json(POLYHAVEN_ASSETS_URL)
    try:
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(cat, f)
    except OSError:
        pass
    return cat


def search_polyhaven(query: str, limit: int = 12) -> list[tuple[str, list[str]]]:
    q_terms = [t.strip().lower() for t in query.split() if t.strip()]
    cat = _polyhaven_catalog()
    scored = []
    for asset_id, meta in cat.items():
        hay = " ".join([asset_id, meta.get("name", "")] + meta.get("categories", []) + meta.get("tags", [])).lower()
        hits = sum(1 for t in q_terms if t in hay)
        if hits > 0 or not q_terms:
            scored.append((hits, asset_id, meta.get("categories", [])))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [(aid, cats) for _, aid, cats in scored[:limit]]


def _pack_gltf_to_glb(gltf: dict, base_url: str, includes: dict, out_glb: str) -> str:
    bin_chunks = bytearray()
    buffers = gltf.get("buffers", [])
    if buffers:
        buf0_uri = buffers[0].get("uri", "")
        if buf0_uri:
            raw_url = includes.get(buf0_uri, {}).get("url") or urllib.parse.urljoin(base_url, buf0_uri)
            bin_chunks.extend(_http_get(raw_url, timeout=60))
            buffers[0].pop("uri", None)

    def _pad4(ba: bytearray, fill: int = 0) -> None:
        rem = len(ba) % 4
        if rem:
            ba.extend(bytes([fill]) * (4 - rem))

    bvs = gltf.setdefault("bufferViews", [])
    for img in gltf.get("images", []):
        uri = img.pop("uri", None)
        if not uri:
            continue
        raw_url = includes.get(uri, {}).get("url") or urllib.parse.urljoin(base_url, uri)
        img_bytes = _http_get(raw_url, timeout=60)
        _pad4(bin_chunks, 0)
        off = len(bin_chunks)
        bin_chunks.extend(img_bytes)
        bv_idx = len(bvs)
        bvs.append({"buffer": 0, "byteOffset": off, "byteLength": len(img_bytes)})
        img["bufferView"] = bv_idx
        ext = os.path.splitext(uri)[1].lower()
        img["mimeType"] = "image/png" if ext == ".png" else ("image/webp" if ext == ".webp" else "image/jpeg")

    _pad4(bin_chunks, 0)
    if not buffers:
        buffers.append({"byteLength": len(bin_chunks)})
    else:
        buffers[0]["byteLength"] = len(bin_chunks)

    json_bytes = bytearray(json.dumps(gltf, separators=(",", ":")).encode("utf-8"))
    _pad4(json_bytes, 0x20)

    total_len = 12 + 8 + len(json_bytes) + 8 + len(bin_chunks)
    os.makedirs(os.path.dirname(os.path.abspath(out_glb)), exist_ok=True)
    with open(out_glb, "wb") as f:
        f.write(struct.pack("<4sII", b"glTF", 2, total_len))
        f.write(struct.pack("<I4s", len(json_bytes), b"JSON"))
        f.write(json_bytes)
        f.write(struct.pack("<I4s", len(bin_chunks), b"BIN\x00"))
        f.write(bin_chunks)
    return out_glb


def fetch_polyhaven_glb(asset_id: str, out_glb: str, res: str = "1k") -> str:
    files = _http_json(POLYHAVEN_FILES_URL.format(asset_id=urllib.parse.quote(asset_id)))
    gltf_map = files.get("gltf", {})
    if not gltf_map:
        raise ValueError(f"Poly Haven asset '{asset_id}' has no glTF export")
    chosen_res = None
    for cand in (res, "1k", "2k", "4k", "8k"):
        if cand in gltf_map and "gltf" in gltf_map[cand]:
            chosen_res = cand
            break
    if not chosen_res:
        chosen_res = sorted(gltf_map.keys())[0]
    entry = gltf_map[chosen_res]["gltf"]
    gltf_url = entry["url"]
    gltf_json = _http_json(gltf_url, timeout=40)
    base_url = gltf_url.rsplit("/", 1)[0] + "/"
    return _pack_gltf_to_glb(gltf_json, base_url, entry.get("include", {}), out_glb)


def search_nasa(query: str, limit: int = 12) -> list[str]:
    q_terms = [t.strip().lower() for t in query.split() if t.strip()]
    dirs = [d["name"] for d in _http_json(NASA_LIST_URL) if isinstance(d, dict) and d.get("type") == "dir"]
    if not q_terms:
        return dirs[:limit]
    scored = []
    for name in dirs:
        low = name.lower()
        hits = sum(1 for t in q_terms if t in low)
        if hits:
            scored.append((hits, name))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [n for _, n in scored[:limit]]


def fetch_nasa_glb(folder_name: str, out_glb: str) -> str:
    url = NASA_DIR_URL.format(folder=urllib.parse.quote(folder_name))
    items = _http_json(url)
    glbs = [it for it in items if isinstance(it, dict) and it.get("name", "").lower().endswith(".glb") and it.get("download_url")]
    if not glbs:
        raise ValueError(f"No .glb file inside NASA 3D folder '{folder_name}'")
    target = sorted(glbs, key=lambda x: x.get("size", 0))[0]
    data = _http_get(target["download_url"], timeout=90)
    if data[:4] != b"glTF":
        raise ValueError(f"Invalid GLB in NASA folder '{folder_name}'")
    os.makedirs(os.path.dirname(os.path.abspath(out_glb)), exist_ok=True)
    with open(out_glb, "wb") as f:
        f.write(data)
    return out_glb


def resolve_auto_glb(topic: str, out_glb: str) -> tuple[str, str, str]:
    low = (topic or "").lower()
    if any(k in low for k in ("space", "apollo", "mars", "rover", "moon", "nasa", "satellite", "orbit", "aerospace", "voyager", "james webb")):
        hits = search_nasa(low, limit=5) or ["Apollo Lunar Module", "Voyager"]
        for h in hits:
            try:
                fetch_nasa_glb(h, out_glb)
                return ("nasa", h, out_glb)
            except Exception:
                continue

    for keywords, model_name in DOMAIN_KHRONOS_MAP:
        if any(kw in low for kw in keywords):
            fetch_khronos_glb(model_name, out_glb)
            return ("khronos", model_name, out_glb)

    try:
        ph_hits = search_polyhaven(low, limit=3)
        for aid, _ in ph_hits:
            try:
                fetch_polyhaven_glb(aid, out_glb, res="1k")
                return ("polyhaven", aid, out_glb)
            except Exception:
                continue
    except Exception:
        pass

    # Reliable high-poly studio camera fallback (clean opaque PBR in PowerPoint COM)
    fetch_khronos_glb("AntiqueCamera", out_glb)
    return ("khronos", "AntiqueCamera", out_glb)


def inspect_glb_stats(glb_path: str) -> dict:
    with open(glb_path, "rb") as f:
        magic, version, total_len = struct.unpack("<4sII", f.read(12))
        if magic != b"glTF" or version != 2:
            raise ValueError(f"Not a glTF 2.0 binary: {glb_path}")
        jlen, _ = struct.unpack("<I4s", f.read(8))
        gltf = json.loads(f.read(jlen).decode("utf-8"))

    accessors = gltf.get("accessors", [])
    vertices = 0
    for mesh in gltf.get("meshes", []):
        for prim in mesh.get("primitives", []):
            pos_idx = prim.get("attributes", {}).get("POSITION")
            if pos_idx is not None and 0 <= pos_idx < len(accessors):
                vertices += int(accessors[pos_idx].get("count", 0))
    return {
        "bytes": total_len,
        "meshes": len(gltf.get("meshes", [])),
        "materials": len(gltf.get("materials", [])),
        "textures": len(gltf.get("textures", [])),
        "vertices": vertices,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch real high-poly CC0/open-license .glb 3D assets.")
    parser.add_argument("--list", action="store_true", help="List curated realistic Khronos GLBs")
    parser.add_argument("--search", default="", help="Search Poly Haven & NASA 3D libraries by keyword")
    parser.add_argument("--source", choices=["khronos", "polyhaven", "nasa"], help="Explicit asset registry")
    parser.add_argument("--name", default="", help="Model name or asset ID in the chosen source")
    parser.add_argument("--res", default="1k", help="Poly Haven texture resolution (default: 1k)")
    parser.add_argument("--auto", default="", help="Auto-resolve best real GLB from topic keywords")
    parser.add_argument("--out", default="", help="Output .glb path")
    args = parser.parse_args()

    if args.list:
        models = list_khronos_models()
        print(f"Khronos realistic GLB catalog ({len(models)}):")
        for m in models:
            print(f"  - {m}")
        return

    if args.search:
        print(f"Poly Haven CC0 matches for '{args.search}':")
        for aid, cats in search_polyhaven(args.search, limit=10):
            print(f"  - {aid:30s} {cats[:3]}")
        print(f"NASA 3D matches for '{args.search}':")
        for folder in search_nasa(args.search, limit=8):
            print(f"  - {folder}")
        return

    if not args.out:
        parser.error("--out is required when fetching a model")

    if args.auto:
        src, picked, path = resolve_auto_glb(args.auto, args.out)
        stats = inspect_glb_stats(path)
        print(f"Resolved real 3D GLB ({src}: {picked}) -> {path} | stats={stats}")
        return

    if not args.source or not args.name:
        parser.error("Provide either --auto <topic> or --source <khronos|polyhaven|nasa> --name <id>")

    if args.source == "khronos":
        fetch_khronos_glb(args.name, args.out)
    elif args.source == "polyhaven":
        fetch_polyhaven_glb(args.name, args.out, res=args.res)
    elif args.source == "nasa":
        fetch_nasa_glb(args.name, args.out)

    stats = inspect_glb_stats(args.out)
    print(f"Fetched real 3D GLB ({args.source}: {args.name}) -> {args.out} | stats={stats}")


if __name__ == "__main__":
    main()
