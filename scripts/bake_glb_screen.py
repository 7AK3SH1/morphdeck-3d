#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Non-destructive glTF Binary (.glb) screen texture baker & chassis recolorer
for custom user-supplied 3D models in morphdeck-3d.

Remaps planar [0..1] UVs onto any target display mesh inside an existing .glb,
embeds a self-lit UI PNG texture (baseColorTexture + emissiveTexture), and
optionally recolors specified metallic chassis materials from an sRGB hex color.
"""

import argparse
import json
import os
import struct
from typing import Optional
import numpy as np


def hex_to_linear_rgb(hex_color: str) -> list:
    """Convert #RRGGBB sRGB hex string to glTF linear [r, g, b]."""
    h = hex_color.strip().lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Expected #RRGGBB hex color, got '{hex_color}'")
    return [((int(h[i:i + 2], 16) / 255.0) ** 2.2) for i in (0, 2, 4)]


def list_glb_info(glb_path: str) -> None:
    """Print meshes and materials inside a GLB file to aid target selection."""
    if not os.path.exists(glb_path):
        raise FileNotFoundError(f"GLB file not found: {glb_path}")
    with open(glb_path, "rb") as f:
        magic, version, _ = struct.unpack("<4sII", f.read(12))
        if magic != b"glTF":
            raise ValueError(f"Not a valid GLB file: {glb_path}")
        jlen, _ = struct.unpack("<I4s", f.read(8))
        gltf = json.loads(f.read(jlen).decode("utf-8"))

    print(f"GLB v{version}: {glb_path}")
    print("Materials:")
    for idx, mat in enumerate(gltf.get("materials", [])):
        print(f"  [{idx}] {mat.get('name', '<unnamed>')}")
    print("Meshes:")
    for m in gltf.get("meshes", []):
        mats = [p.get("material") for p in m.get("primitives", [])]
        print(f"  {m.get('name', '<unnamed>')} -> materials {mats}")


def bake_glb_screen(
    src_glb: str,
    mesh_name: str,
    mat_idx: int,
    screen_png: str,
    dst_glb: str,
    emissive: float = 0.85,
    chassis_color: Optional[str] = None,
    chassis_materials: Optional[dict] = None,
) -> None:
    """Embed a custom UI PNG onto a GLB mesh material with planar [0..1] UVs."""
    if not os.path.exists(src_glb):
        raise FileNotFoundError(f"Input GLB not found: {src_glb}")
    if not os.path.exists(screen_png):
        raise FileNotFoundError(f"Screen PNG not found: {screen_png}")

    with open(src_glb, "rb") as f:
        magic, version, _ = struct.unpack("<4sII", f.read(12))
        if magic != b"glTF":
            raise ValueError(f"Invalid GLB header: {src_glb}")
        jlen, _ = struct.unpack("<I4s", f.read(8))
        gltf = json.loads(f.read(jlen).decode("utf-8"))
        blen, _ = struct.unpack("<I4s", f.read(8))
        bin_data = bytearray(f.read(blen))

    target_mesh = next(
        (m for m in gltf.get("meshes", []) if m.get("name") == mesh_name),
        None,
    )
    if target_mesh is None:
        raise ValueError(f"Mesh '{mesh_name}' not found in {src_glb}")
    if mat_idx < 0 or mat_idx >= len(gltf.get("materials", [])):
        raise ValueError(f"Material index {mat_idx} out of range")

    if chassis_color and chassis_materials:
        lin_rgb = hex_to_linear_rgb(chassis_color)
        for c_idx, mult in chassis_materials.items():
            if 0 <= c_idx < len(gltf.get("materials", [])):
                c_mat = gltf["materials"][c_idx]
                pbr = c_mat.setdefault("pbrMetallicRoughness", {})
                pbr["baseColorFactor"] = [
                    min(1.0, max(0.0, ch * mult)) for ch in lin_rgb
                ] + [1.0]

    orig_bin = bytes(bin_data)

    for prim in target_mesh["primitives"]:
        pos_acc = gltf["accessors"][prim["attributes"]["POSITION"]]
        pos_bv = gltf["bufferViews"][pos_acc["bufferView"]]
        pos_off = pos_bv.get("byteOffset", 0) + pos_acc.get("byteOffset", 0)
        stride = pos_bv.get("byteStride", 12)
        cnt = pos_acc["count"]

        if stride == 12:
            pos = np.frombuffer(orig_bin, dtype=np.float32, count=cnt * 3, offset=pos_off).reshape(cnt, 3)
        else:
            rows = [struct.unpack_from("<3f", orig_bin, pos_off + i * stride) for i in range(cnt)]
            pos = np.array(rows, dtype=np.float32)

        xmin, xmax = float(pos[:, 0].min()), float(pos[:, 0].max())
        ymin, ymax = float(pos[:, 1].min()), float(pos[:, 1].max())
        dx = max(1e-6, xmax - xmin)
        dy = max(1e-6, ymax - ymin)

        u = np.clip((pos[:, 0] - xmin) / dx, 0.0, 1.0)
        v = np.clip((ymax - pos[:, 1]) / dy, 0.0, 1.0)
        uv_bytes = np.stack([u, v], axis=1).astype("<f4").tobytes()

        while len(bin_data) % 4 != 0:
            bin_data.append(0)
        uv_off = len(bin_data)
        bin_data.extend(uv_bytes)

        uv_bv_idx = len(gltf["bufferViews"])
        gltf["bufferViews"].append({
            "buffer": 0,
            "byteOffset": uv_off,
            "byteLength": len(uv_bytes),
            "target": 34962,
        })
        uv_acc_idx = len(gltf["accessors"])
        gltf["accessors"].append({
            "bufferView": uv_bv_idx,
            "byteOffset": 0,
            "componentType": 5126,
            "count": cnt,
            "type": "VEC2",
            "min": [0.0, 0.0],
            "max": [1.0, 1.0],
        })
        prim["attributes"]["TEXCOORD_0"] = uv_acc_idx

    with open(screen_png, "rb") as pf:
        png_bytes = pf.read()

    while len(bin_data) % 4 != 0:
        bin_data.append(0)
    img_off = len(bin_data)
    bin_data.extend(png_bytes)

    img_bv_idx = len(gltf["bufferViews"])
    gltf["bufferViews"].append({
        "buffer": 0,
        "byteOffset": img_off,
        "byteLength": len(png_bytes),
    })

    gltf.setdefault("images", [])
    img_idx = len(gltf["images"])
    gltf["images"].append({
        "bufferView": img_bv_idx,
        "mimeType": "image/png",
        "name": f"screen_{mesh_name}",
    })

    gltf.setdefault("samplers", [])
    samp_idx = len(gltf["samplers"])
    gltf["samplers"].append({
        "magFilter": 9729,
        "minFilter": 9729,
        "wrapS": 33071,
        "wrapT": 33071,
    })

    gltf.setdefault("textures", [])
    tex_idx = len(gltf["textures"])
    gltf["textures"].append({
        "sampler": samp_idx,
        "source": img_idx,
    })

    mat = gltf["materials"][mat_idx]
    mat["doubleSided"] = True
    mat["pbrMetallicRoughness"] = {
        "baseColorFactor": [1.0, 1.0, 1.0, 1.0],
        "baseColorTexture": {"index": tex_idx},
        "metallicFactor": 0.0,
        "roughnessFactor": 0.14,
    }
    mat["emissiveTexture"] = {"index": tex_idx}
    mat["emissiveFactor"] = [emissive, emissive, emissive]
    mat.pop("normalTexture", None)
    mat.pop("extensions", None)

    while len(bin_data) % 4 != 0:
        bin_data.append(0)
    gltf["buffers"][0]["byteLength"] = len(bin_data)

    json_bytes = bytearray(json.dumps(gltf, separators=(",", ":")).encode("utf-8"))
    while len(json_bytes) % 4 != 0:
        json_bytes.append(0x20)

    total_len = 12 + 8 + len(json_bytes) + 8 + len(bin_data)
    os.makedirs(os.path.dirname(os.path.abspath(dst_glb)), exist_ok=True)
    with open(dst_glb, "wb") as out:
        out.write(struct.pack("<4sII", b"glTF", version, total_len))
        out.write(struct.pack("<I4s", len(json_bytes), b"JSON"))
        out.write(json_bytes)
        out.write(struct.pack("<I4s", len(bin_data), b"BIN\x00"))
        out.write(bin_data)


def main():
    parser = argparse.ArgumentParser(description="Bake a custom UI PNG onto a user-supplied 3D GLB screen mesh.")
    parser.add_argument("--input", required=True, help="Input custom .glb file path")
    parser.add_argument("--inspect", action="store_true", help="List meshes and materials and exit")
    parser.add_argument("--mesh", help="Target screen mesh name")
    parser.add_argument("--material", type=int, help="Target screen material index")
    parser.add_argument("--image", help="UI screenshot PNG path")
    parser.add_argument("--output", help="Output .glb file path")
    parser.add_argument("--chassis-color", help="Optional #RRGGBB hex tint")
    parser.add_argument("--chassis-mat", type=int, action="append", default=[], help="Material index to tint (repeatable)")
    parser.add_argument("--emissive", type=float, default=0.85, help="Self-lit emissive multiplier (default: 0.85)")
    args = parser.parse_args()

    if args.inspect:
        list_glb_info(args.input)
        return

    if not all([args.mesh, args.material is not None, args.image, args.output]):
        parser.error("--mesh, --material, --image, and --output are required when not using --inspect")

    tint_map = {idx: 1.0 for idx in args.chassis_mat}
    bake_glb_screen(
        src_glb=args.input,
        mesh_name=args.mesh,
        mat_idx=args.material,
        screen_png=args.image,
        dst_glb=args.output,
        emissive=args.emissive,
        chassis_color=args.chassis_color,
        chassis_materials=tint_map,
    )
    print(f"Baked custom GLB cleanly -> {args.output} ({os.path.getsize(args.output)} bytes)")


if __name__ == "__main__":
    main()
