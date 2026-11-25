#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Pure-NumPy + Pillow Studio 3D GLB Software Renderer for morphdeck-3d.

Used for quick visual 3D inspection (`morphdeck.py preview-3d`) and cross-platform
OpenXML builds on macOS/Linux/CI using real high-poly .glb assets.
"""

import argparse
import io
import json
import math
import os
import struct
import sys
import numpy as np
from PIL import Image, ImageDraw

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

try:
    from inject_3d_studio_lighting import LIGHT_RIGS
except ImportError:
    LIGHT_RIGS = {}


def load_glb_scene(glb_path: str) -> dict:
    with open(glb_path, "rb") as f:
        magic, version, _ = struct.unpack("<4sII", f.read(12))
        if magic != b"glTF" or version != 2:
            raise ValueError(f"Unsupported GLB header in {glb_path}")
        jlen, _ = struct.unpack("<I4s", f.read(8))
        gltf = json.loads(f.read(jlen).decode("utf-8"))
        blen, _ = struct.unpack("<I4s", f.read(8))
        bin_data = f.read(blen)

    accessors = gltf.get("accessors", [])
    bvs = gltf.get("bufferViews", [])
    materials = gltf.get("materials", [])
    textures = gltf.get("textures", [])
    images = gltf.get("images", [])

    tex_arrs = {}
    for ti, tx in enumerate(textures):
        src_i = tx.get("source")
        if src_i is not None and 0 <= src_i < len(images):
            bv_i = images[src_i].get("bufferView")
            if bv_i is not None and 0 <= bv_i < len(bvs):
                bv = bvs[bv_i]
                off = bv.get("byteOffset", 0)
                lng = bv["byteLength"]
                im = Image.open(io.BytesIO(bin_data[off : off + lng])).convert("RGB")
                tex_arrs[ti] = np.asarray(im, dtype=np.float32) / 255.0

    def read_acc(acc_i: int) -> np.ndarray:
        acc = accessors[acc_i]
        bv = bvs[acc["bufferView"]]
        stride = bv.get("byteStride", 0)
        off = bv.get("byteOffset", 0) + acc.get("byteOffset", 0)
        cnt = acc["count"]
        ctype = acc["componentType"]
        atype = acc["type"]
        dt_map = {5126: "<f4", 5125: "<u4", 5123: "<u2", 5121: "<u1"}
        cols_map = {"SCALAR": 1, "VEC2": 2, "VEC3": 3, "VEC4": 4}
        dt = np.dtype(dt_map[ctype])
        cols = cols_map[atype]
        if stride and stride > dt.itemsize * cols:
            out = np.empty((cnt, cols), dtype=dt)
            for k in range(cnt):
                out[k] = np.frombuffer(bin_data, dtype=dt, count=cols, offset=off + k * stride)
            return out
        arr = np.frombuffer(bin_data, dtype=dt, count=cnt * cols, offset=off)
        return arr.reshape((cnt, cols)) if cols > 1 else arr.copy()

    primitives = []
    for mesh in gltf.get("meshes", []):
        for prim in mesh.get("primitives", []):
            attrs = prim["attributes"]
            if "POSITION" not in attrs:
                continue
            pos = read_acc(attrs["POSITION"]).astype(np.float32)
            norm = (
                read_acc(attrs["NORMAL"]).astype(np.float32)
                if "NORMAL" in attrs
                else np.tile([0, 0, 1], (len(pos), 1)).astype(np.float32)
            )
            uv = (
                read_acc(attrs["TEXCOORD_0"]).astype(np.float32)
                if "TEXCOORD_0" in attrs
                else np.zeros((len(pos), 2), dtype=np.float32)
            )
            if "indices" in prim:
                idx = read_acc(prim["indices"]).astype(np.int32)
            else:
                idx = np.arange(len(pos), dtype=np.int32)
            idx = idx[: (len(idx) // 3) * 3].reshape(-1, 3)

            mat_i = prim.get("material", 0)
            mat = materials[mat_i] if 0 <= mat_i < len(materials) else {}
            pbr = mat.get("pbrMetallicRoughness", {})
            base_col = np.array(pbr.get("baseColorFactor", [0.8, 0.8, 0.8, 1.0])[:3], dtype=np.float32)
            em_col = np.array(mat.get("emissiveFactor", [0.0, 0.0, 0.0])[:3], dtype=np.float32)
            metal = float(pbr.get("metallicFactor", 0.4))
            rough = float(pbr.get("roughnessFactor", 0.3))

            tex = None
            tex_info = pbr.get("baseColorTexture") or mat.get("emissiveTexture")
            if tex_info is not None:
                tex = tex_arrs.get(tex_info.get("index"))

            primitives.append({
                "pos": pos, "norm": norm, "uv": uv, "idx": idx,
                "base_col": base_col, "em_col": em_col, "metal": metal, "rough": rough, "tex": tex,
            })

    return {"primitives": primitives}


def _rot_mat(rx: float, ry: float, rz: float) -> np.ndarray:
    ax, ay, az = math.radians(rx), math.radians(ry), math.radians(rz)
    Rx = np.array([[1, 0, 0], [0, math.cos(ax), -math.sin(ax)], [0, math.sin(ax), math.cos(ax)]], dtype=np.float32)
    Ry = np.array([[math.cos(ay), 0, math.sin(ay)], [0, 1, 0], [-math.sin(ay), 0, math.cos(ay)]], dtype=np.float32)
    Rz = np.array([[math.cos(az), -math.sin(az), 0], [math.sin(az), math.cos(az), 0], [0, 0, 1]], dtype=np.float32)
    return (Rz @ Ry @ Rx).T


def render_glb_to_image(
    glb_path: str,
    width: int = 720,
    height: int = 720,
    rx: float = 12.0,
    ry: float = -24.0,
    rz: float = 2.0,
    theme: str = "commercial_studio",
    max_tris_per_prim: int = 120000,
) -> Image.Image:
    """Render any real glTF 2.0 .glb binary into a crisp transparent-background studio RGBA image."""
    scene = load_glb_scene(glb_path)
    R = _rot_mat(rx, ry, rz)

    all_pts = np.vstack([p["pos"] @ R for p in scene["primitives"]])
    center = 0.5 * (all_pts.min(axis=0) + all_pts.max(axis=0))
    span = max(0.05, float(np.max(all_pts.max(axis=0) - all_pts.min(axis=0))))
    cam_dist = span * 2.45
    scale = min(width, height) * 0.80 / (span / cam_dist)

    key_l = np.array([0.52, 0.72, 0.48], dtype=np.float32); key_l /= np.linalg.norm(key_l)
    fill_l = np.array([-0.58, 0.44, 0.68], dtype=np.float32); fill_l /= np.linalg.norm(fill_l)
    view_v = np.array([0.0, 0.0, 1.0], dtype=np.float32)
    half_v = key_l + view_v; half_v /= np.linalg.norm(half_v)

    rig = LIGHT_RIGS.get(theme) or LIGHT_RIGS.get("commercial_studio")
    if rig:
        amb_col = np.array(rig["ambient"][:3], dtype=np.float32) / 100000.0 * 0.48
        k_rgb = np.array(rig["lights"][0][:3], dtype=np.float32) / 100000.0
        f_rgb = np.array(rig["lights"][1][:3], dtype=np.float32) / 100000.0
        r_rgb = np.array(rig["lights"][2][:3], dtype=np.float32) / 100000.0
    else:
        amb_col = np.array([0.24, 0.24, 0.26], dtype=np.float32)
        k_rgb = np.array([1.0, 0.84, 0.68], dtype=np.float32)
        f_rgb = np.array([0.45, 0.68, 0.96], dtype=np.float32)
        r_rgb = np.array([0.88, 0.80, 1.0], dtype=np.float32)

    tri_xy_all, tri_z_all, tri_col_all = [], [], []
    for prim in scene["primitives"]:
        pos = (prim["pos"] @ R) - center
        norm = prim["norm"] @ R
        idx = prim["idx"]
        if len(idx) > max_tris_per_prim:
            step = max(1, len(idx) // max_tris_per_prim)
            idx = idx[::step]
        uv = prim["uv"]

        p_tri = pos[idx]
        n_tri = norm[idx].mean(axis=1)
        n_len = np.linalg.norm(n_tri, axis=1, keepdims=True)
        n_tri = n_tri / np.where(n_len < 1e-6, 1.0, n_len)
        n_face = np.where(n_tri[:, 2:3] >= 0, n_tri, -n_tri)

        z_cam = np.maximum(0.02, cam_dist - p_tri[:, :, 2])
        px = width * 0.5 + (p_tri[:, :, 0] / z_cam) * scale
        py = height * 0.5 - (p_tri[:, :, 1] / z_cam) * scale
        tri_xy = np.stack([px, py], axis=-1)
        z_depth = p_tri[:, :, 2].mean(axis=1)

        if prim["tex"] is not None:
            tex = prim["tex"]
            th, tw, _ = tex.shape
            uv_c = uv[idx].mean(axis=1)
            tx = np.clip((np.mod(uv_c[:, 0], 1.0) * (tw - 1)).astype(np.int32), 0, tw - 1)
            ty = np.clip((np.mod(uv_c[:, 1], 1.0) * (th - 1)).astype(np.int32), 0, th - 1)
            albedo_lin = np.power(tex[ty, tx], 2.0)
        else:
            albedo_lin = np.tile(prim["base_col"], (len(idx), 1))

        ndotl_key = np.clip(np.sum(n_face * key_l[None, :], axis=1), 0.0, 1.0)[:, None]
        ndotl_fill = np.clip(np.sum(n_face * fill_l[None, :], axis=1), 0.0, 1.0)[:, None]
        rim = np.power(1.0 - np.abs(n_face[:, 2:3]), 2.5) * 0.35
        ndoth = np.clip(np.sum(n_face * half_v[None, :], axis=1), 0.0, 1.0)[:, None]
        shininess = max(8.0, 72.0 * (1.0 - prim["rough"]))
        spec = np.power(ndoth, shininess) * (0.35 + 0.55 * prim["metal"])

        lit = albedo_lin * (amb_col[None, :] + (ndotl_key * 0.78) * k_rgb[None, :] + (ndotl_fill * 0.34) * f_rgb[None, :])
        lit = lit + k_rgb[None, :] * (spec * 0.90) + r_rgb[None, :] * (rim * 0.40) + prim["em_col"][None, :] * 0.85

        srgb = np.power(np.clip(lit, 0.0, 1.0), 1.0 / 2.2)
        rgb_u8 = np.clip(srgb * 255.0, 0, 255).astype(np.uint8)

        v01 = tri_xy[:, 1] - tri_xy[:, 0]
        v02 = tri_xy[:, 2] - tri_xy[:, 0]
        keep = np.abs(v01[:, 0] * v02[:, 1] - v01[:, 1] * v02[:, 0]) > 0.12
        if np.any(keep):
            tri_xy_all.append(tri_xy[keep])
            tri_z_all.append(z_depth[keep])
            tri_col_all.append(rgb_u8[keep])

    out = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    if not tri_xy_all:
        return out

    tri_xy = np.vstack(tri_xy_all)
    tri_z = np.concatenate(tri_z_all)
    tri_col = np.vstack(tri_col_all)
    order = np.argsort(tri_z)
    d = ImageDraw.Draw(out, "RGBA")
    for idx in order:
        pts = [tuple(tri_xy[idx, 0]), tuple(tri_xy[idx, 1]), tuple(tri_xy[idx, 2])]
        c = tri_col[idx]
        d.polygon(pts, fill=(int(c[0]), int(c[1]), int(c[2]), 255))
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a studio RGBA preview PNG from any real glTF 2.0 .glb file.")
    parser.add_argument("--glb", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--size", type=int, default=720)
    parser.add_argument("--rx", type=float, default=12.0)
    parser.add_argument("--ry", type=float, default=-24.0)
    parser.add_argument("--rz", type=float, default=2.0)
    parser.add_argument("--theme", default="commercial_studio")
    args = parser.parse_args()

    img = render_glb_to_image(args.glb, args.size, args.size, args.rx, args.ry, args.rz, theme=args.theme)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    img.save(args.out, "PNG")
    print(f"Rendered 3D GLB studio preview -> {args.out}")


if __name__ == "__main__":
    main()
