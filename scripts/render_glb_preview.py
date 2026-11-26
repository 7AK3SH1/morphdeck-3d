#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
High-Precision Z-Buffered Smooth-Phong & Bilinear-PBR Studio 3D (.glb) Renderer
for morphdeck-3d v4.0.0.

Upgrades over v3:
  - Full glTF 2.0 scene-graph node hierarchy traversal (TRS quaternions + 4x4 matrices)
  - True per-pixel floating-point Z-buffer (zero painter's-algorithm bleed-through)
  - Smooth interpolated vertex normals + Blinn-Phong studio 3-point specular & rim shading
  - Bilinear PBR baseColor & emissive texture sampling
  - 2x Supersampled Anti-Aliasing (SSAA) downsampled with LANCZOS
"""

import argparse
import io
import json
import math
import os
import struct
import sys
import numpy as np
from PIL import Image

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

try:
    from inject_3d_studio_lighting import LIGHT_RIGS
except ImportError:
    LIGHT_RIGS = {}


def _quat_to_mat4(q: list[float]) -> np.ndarray:
    x, y, z, w = float(q[0]), float(q[1]), float(q[2]), float(q[3])
    n = math.sqrt(x * x + y * y + z * z + w * w)
    if n > 1e-8:
        x, y, z, w = x / n, y / n, z / n, w / n
    m = np.eye(4, dtype=np.float32)
    m[0, 0] = 1.0 - 2.0 * (y * y + z * z)
    m[0, 1] = 2.0 * (x * y - z * w)
    m[0, 2] = 2.0 * (x * z + y * w)
    m[1, 0] = 2.0 * (x * y + z * w)
    m[1, 1] = 1.0 - 2.0 * (x * x + z * z)
    m[1, 2] = 2.0 * (y * z - x * w)
    m[2, 0] = 2.0 * (x * z - y * w)
    m[2, 1] = 2.0 * (y * z + x * w)
    m[2, 2] = 1.0 - 2.0 * (x * x + y * y)
    return m


def _node_local_mat4(node: dict) -> np.ndarray:
    if "matrix" in node:
        return np.array(node["matrix"], dtype=np.float32).reshape((4, 4)).T
    T = np.eye(4, dtype=np.float32)
    if "translation" in node:
        T[:3, 3] = np.array(node["translation"], dtype=np.float32)
    R = _quat_to_mat4(node["rotation"]) if "rotation" in node else np.eye(4, dtype=np.float32)
    S = np.eye(4, dtype=np.float32)
    if "scale" in node:
        s = node["scale"]
        S[0, 0], S[1, 1], S[2, 2] = float(s[0]), float(s[1]), float(s[2])
    return T @ R @ S


def _compute_smooth_normals(pos: np.ndarray, idx: np.ndarray) -> np.ndarray:
    v0 = pos[idx[:, 0]]
    v1 = pos[idx[:, 1]]
    v2 = pos[idx[:, 2]]
    fn = np.cross(v1 - v0, v2 - v0)
    norms = np.zeros_like(pos, dtype=np.float32)
    np.add.at(norms, idx[:, 0], fn)
    np.add.at(norms, idx[:, 1], fn)
    np.add.at(norms, idx[:, 2], fn)
    lens = np.linalg.norm(norms, axis=1, keepdims=True)
    return np.where(lens > 1e-8, norms / np.maximum(lens, 1e-8), np.array([0.0, 0.0, 1.0], dtype=np.float32))


def load_glb_scene(glb_path: str) -> dict:
    with open(glb_path, "rb") as f:
        magic, version, _ = struct.unpack("<4sII", f.read(12))
        if magic != b"glTF" or version != 2:
            raise ValueError(f"Unsupported GLB header in {glb_path}")
        jlen, _ = struct.unpack("<I4s", f.read(8))
        gltf = json.loads(f.read(jlen).decode("utf-8"))
        bheader = f.read(8)
        bin_data = f.read(struct.unpack("<I4s", bheader)[0]) if len(bheader) == 8 else b""

    accessors = gltf.get("accessors", [])
    bvs = gltf.get("bufferViews", [])
    materials = gltf.get("materials", [])
    textures = gltf.get("textures", [])
    images = gltf.get("images", [])
    meshes = gltf.get("meshes", [])
    nodes = gltf.get("nodes", [])

    tex_arrs = {}
    for ti, tx in enumerate(textures):
        src_i = tx.get("source")
        if src_i is not None and 0 <= src_i < len(images):
            bv_i = images[src_i].get("bufferView")
            if bv_i is not None and 0 <= bv_i < len(bvs):
                bv = bvs[bv_i]
                off = bv.get("byteOffset", 0)
                lng = bv["byteLength"]
                im = Image.open(io.BytesIO(bin_data[off : off + lng])).convert("RGBA")
                arr = np.asarray(im, dtype=np.float32) / 255.0
                # Convert sRGB albedo to linear space for physically accurate lighting
                arr_lin = np.empty_like(arr)
                arr_lin[..., :3] = np.power(arr[..., :3], 2.2)
                arr_lin[..., 3] = arr[..., 3]
                tex_arrs[ti] = arr_lin

    def read_acc(acc_i: int) -> np.ndarray:
        acc = accessors[acc_i]
        bv = bvs[acc["bufferView"]]
        stride = bv.get("byteStride", 0)
        off = bv.get("byteOffset", 0) + acc.get("byteOffset", 0)
        cnt = acc["count"]
        ctype = acc["componentType"]
        atype = acc["type"]
        dt_map = {5126: "<f4", 5125: "<u4", 5123: "<u2", 5122: "<i2", 5121: "<u1", 5120: "<i1"}
        cols_map = {"SCALAR": 1, "VEC2": 2, "VEC3": 3, "VEC4": 4}
        dt = np.dtype(dt_map[ctype])
        cols = cols_map[atype]
        if stride and stride > dt.itemsize * cols:
            raw = np.frombuffer(bin_data, dtype=np.uint8, count=cnt * stride, offset=off)
            raw = raw.reshape((cnt, stride))[:, : dt.itemsize * cols]
            arr = np.frombuffer(raw.tobytes(), dtype=dt).reshape((cnt, cols))
        else:
            arr = np.frombuffer(bin_data, dtype=dt, count=cnt * cols, offset=off)
            arr = arr.reshape((cnt, cols)) if cols > 1 else arr.copy()
        if acc.get("normalized"):
            if ctype == 5121:
                arr = arr.astype(np.float32) / 255.0
            elif ctype == 5123:
                arr = arr.astype(np.float32) / 65535.0
        return arr

    # Traverse scene graph nodes so multi-mesh assemblies align accurately
    mesh_instances: list[tuple[int, np.ndarray]] = []
    if nodes:
        scenes = gltf.get("scenes", [])
        scene_idx = gltf.get("scene", 0)
        root_nodes = scenes[scene_idx].get("nodes", []) if scenes and 0 <= scene_idx < len(scenes) else list(range(len(nodes)))

        visited = set()

        def _walk(n_idx: int, parent_m: np.ndarray) -> None:
            if n_idx in visited or not (0 <= n_idx < len(nodes)):
                return
            visited.add(n_idx)
            nd = nodes[n_idx]
            world_m = parent_m @ _node_local_mat4(nd)
            if "mesh" in nd and 0 <= nd["mesh"] < len(meshes):
                mesh_instances.append((nd["mesh"], world_m))
            for ch in nd.get("children", []):
                _walk(ch, world_m)

        for rn in root_nodes:
            _walk(rn, np.eye(4, dtype=np.float32))

    if not mesh_instances:
        mesh_instances = [(mi, np.eye(4, dtype=np.float32)) for mi in range(len(meshes))]

    primitives = []
    for mesh_i, world_m in mesh_instances:
        lin_m = world_m[:3, :3]
        trans_v = world_m[:3, 3]
        try:
            norm_m = np.linalg.inv(lin_m).T
        except np.linalg.LinAlgError:
            norm_m = lin_m

        for prim in meshes[mesh_i].get("primitives", []):
            mode = prim.get("mode", 4)
            if mode != 4:
                continue
            attrs = prim.get("attributes", {})
            if "POSITION" not in attrs:
                continue
            pos_local = read_acc(attrs["POSITION"]).astype(np.float32)
            if "indices" in prim:
                idx = read_acc(prim["indices"]).astype(np.int32)
            else:
                idx = np.arange(len(pos_local), dtype=np.int32)
            idx = idx[: (len(idx) // 3) * 3].reshape(-1, 3)
            if len(idx) == 0:
                continue

            pos = (pos_local @ lin_m.T) + trans_v[None, :]
            if "NORMAL" in attrs:
                norm_local = read_acc(attrs["NORMAL"]).astype(np.float32)
                norm = norm_local @ norm_m.T
                nl = np.linalg.norm(norm, axis=1, keepdims=True)
                norm = np.where(nl > 1e-7, norm / np.maximum(nl, 1e-7), np.array([0.0, 1.0, 0.0], dtype=np.float32))
            else:
                norm = _compute_smooth_normals(pos, idx)

            uv = (
                read_acc(attrs["TEXCOORD_0"]).astype(np.float32)[:, :2]
                if "TEXCOORD_0" in attrs
                else np.zeros((len(pos), 2), dtype=np.float32)
            )

            mat_i = prim.get("material", -1)
            mat = materials[mat_i] if 0 <= mat_i < len(materials) else {}
            if mat.get("alphaMode") == "BLEND":
                base_cf = mat.get("pbrMetallicRoughness", {}).get("baseColorFactor", [1, 1, 1, 1])
                if len(base_cf) >= 4 and float(base_cf[3]) < 0.18:
                    # Skip near-invisible glass shells so inner dials/optics remain visible
                    continue

            pbr = mat.get("pbrMetallicRoughness", {})
            base_rgba = pbr.get("baseColorFactor", [0.82, 0.82, 0.84, 1.0])
            base_col = np.power(np.array(base_rgba[:3], dtype=np.float32), 2.2)
            em_col = np.array(mat.get("emissiveFactor", [0.0, 0.0, 0.0])[:3], dtype=np.float32)
            metal = float(pbr.get("metallicFactor", 0.35))
            rough = float(pbr.get("roughnessFactor", 0.38))

            tex = None
            tex_info = pbr.get("baseColorTexture")
            if tex_info is not None:
                tex = tex_arrs.get(tex_info.get("index"))
            em_tex = None
            em_info = mat.get("emissiveTexture")
            if em_info is not None:
                em_tex = tex_arrs.get(em_info.get("index"))

            primitives.append({
                "pos": pos,
                "norm": norm,
                "uv": uv,
                "idx": idx,
                "base_col": base_col,
                "em_col": em_col,
                "metal": metal,
                "rough": rough,
                "tex": tex,
                "em_tex": em_tex,
            })

    return {"primitives": primitives}


def _rot_mat(rx: float, ry: float, rz: float) -> np.ndarray:
    ax, ay, az = math.radians(rx), math.radians(ry), math.radians(rz)
    Rx = np.array([[1, 0, 0], [0, math.cos(ax), -math.sin(ax)], [0, math.sin(ax), math.cos(ax)]], dtype=np.float32)
    Ry = np.array([[math.cos(ay), 0, math.sin(ay)], [0, 1, 0], [-math.sin(ay), 0, math.cos(ay)]], dtype=np.float32)
    Rz = np.array([[math.cos(az), -math.sin(az), 0], [math.sin(az), math.cos(az), 0], [0, 0, 1]], dtype=np.float32)
    return (Rz @ Ry @ Rx).T


def _sample_bilinear(tex: np.ndarray, u: np.ndarray, v: np.ndarray) -> np.ndarray:
    th, tw, _ = tex.shape
    uf = np.mod(u, 1.0) * (tw - 1)
    vf = np.mod(v, 1.0) * (th - 1)
    x0 = np.floor(uf).astype(np.int32)
    y0 = np.floor(vf).astype(np.int32)
    x1 = np.minimum(x0 + 1, tw - 1)
    y1 = np.minimum(y0 + 1, th - 1)
    fx = (uf - x0)[..., None]
    fy = (vf - y0)[..., None]
    c00 = tex[y0, x0, :3]
    c10 = tex[y0, x1, :3]
    c01 = tex[y1, x0, :3]
    c11 = tex[y1, x1, :3]
    return (c00 * (1.0 - fx) + c10 * fx) * (1.0 - fy) + (c01 * (1.0 - fx) + c11 * fx) * fy


def render_glb_to_image(
    glb_path: str,
    width: int = 720,
    height: int = 720,
    rx: float = 12.0,
    ry: float = -24.0,
    rz: float = 2.0,
    theme: str = "commercial_studio",
    max_tris_per_prim: int = 90000,
    ssaa: int = 2,
) -> Image.Image:
    """Render any glTF 2.0 .glb binary with Z-buffered smooth shading, bilinear PBR textures, and 2x SSAA."""
    scene = load_glb_scene(glb_path)
    if not scene["primitives"]:
        return Image.new("RGBA", (width, height), (0, 0, 0, 0))

    ss = max(1, min(2, int(ssaa))) if max(width, height) <= 960 else 1
    W = width * ss
    H = height * ss
    R = _rot_mat(rx, ry, rz)

    all_pts = np.vstack([p["pos"] @ R for p in scene["primitives"]])
    pmin = all_pts.min(axis=0)
    pmax = all_pts.max(axis=0)
    center = 0.5 * (pmin + pmax)
    span = max(0.05, float(np.max(pmax - pmin)))
    cam_dist = span * 2.40
    focal = min(W, H) * 0.82 / (span / cam_dist)

    key_l = np.array([0.48, 0.68, 0.55], dtype=np.float32)
    key_l /= np.linalg.norm(key_l)
    fill_l = np.array([-0.62, 0.38, 0.68], dtype=np.float32)
    fill_l /= np.linalg.norm(fill_l)
    rim_l = np.array([-0.25, 0.55, -0.80], dtype=np.float32)
    rim_l /= np.linalg.norm(rim_l)
    view_v = np.array([0.0, 0.0, 1.0], dtype=np.float32)
    half_key = key_l + view_v
    half_key /= np.linalg.norm(half_key)

    rig = LIGHT_RIGS.get(theme) or LIGHT_RIGS.get("commercial_studio")
    if rig:
        amb_col = np.array(rig["ambient"][:3], dtype=np.float32) / 100000.0 * 0.52
        k_rgb = np.array(rig["lights"][0][:3], dtype=np.float32) / 100000.0
        f_rgb = np.array(rig["lights"][1][:3], dtype=np.float32) / 100000.0
        r_rgb = np.array(rig["lights"][2][:3], dtype=np.float32) / 100000.0
    else:
        amb_col = np.array([0.25, 0.26, 0.28], dtype=np.float32)
        k_rgb = np.array([1.0, 0.88, 0.74], dtype=np.float32)
        f_rgb = np.array([0.46, 0.70, 0.96], dtype=np.float32)
        r_rgb = np.array([0.90, 0.84, 1.0], dtype=np.float32)

    zbuf = np.full((H, W), np.inf, dtype=np.float32)
    cbuf = np.zeros((H, W, 3), dtype=np.float32)
    abuf = np.zeros((H, W), dtype=np.uint8)

    for prim in scene["primitives"]:
        pos = (prim["pos"] @ R) - center[None, :]
        norm = prim["norm"] @ R
        nl = np.linalg.norm(norm, axis=1, keepdims=True)
        norm = norm / np.where(nl < 1e-6, 1.0, nl)

        idx = prim["idx"]
        if len(idx) > max_tris_per_prim:
            step = max(1, len(idx) // max_tris_per_prim)
            idx = idx[::step]

        z_cam_v = np.maximum(0.05, cam_dist - pos[:, 2])
        sx = W * 0.5 + (pos[:, 0] / z_cam_v) * focal
        sy = H * 0.5 - (pos[:, 1] / z_cam_v) * focal

        # Per-vertex lighting factors (smooth Gouraud/Phong shading across triangles)
        nf = np.where(norm[:, 2:3] >= -0.05, norm, -norm)
        ndotl_k = np.clip(nf @ key_l, 0.0, 1.0)
        ndotl_f = np.clip(nf @ fill_l, 0.0, 1.0)
        ndotl_r = np.clip(nf @ rim_l, 0.0, 1.0)
        hemi = 0.5 + 0.5 * np.clip(nf[:, 1], -1.0, 1.0)
        fresnel = np.power(1.0 - np.clip(np.abs(nf[:, 2]), 0.0, 1.0), 2.6)
        ndoth = np.clip(nf @ half_key, 0.0, 1.0)
        shininess = max(10.0, 96.0 * (1.0 - prim["rough"] * 0.82))
        spec_str = 0.28 + 0.72 * prim["metal"]
        v_spec = (np.power(ndoth, shininess) * spec_str)[:, None] * k_rgb[None, :] + (
            (fresnel * 0.34 + ndotl_r * 0.22)[:, None] * r_rgb[None, :]
        )
        v_diff = (
            amb_col[None, :] * (0.72 + 0.50 * hemi[:, None])
            + (ndotl_k * 0.86)[:, None] * k_rgb[None, :]
            + (ndotl_f * 0.40)[:, None] * f_rgb[None, :]
        )

        # Triangle coordinates
        x0, x1, x2 = sx[idx[:, 0]], sx[idx[:, 1]], sx[idx[:, 2]]
        y0, y1, y2 = sy[idx[:, 0]], sy[idx[:, 1]], sy[idx[:, 2]]
        z0, z1, z2 = z_cam_v[idx[:, 0]], z_cam_v[idx[:, 1]], z_cam_v[idx[:, 2]]

        denom = (y1 - y2) * (x0 - x2) + (x2 - x1) * (y0 - y2)
        valid = np.abs(denom) > 0.04
        xmin = np.maximum(0, np.floor(np.minimum(np.minimum(x0, x1), x2)).astype(np.int32))
        xmax = np.minimum(W - 1, np.ceil(np.maximum(np.maximum(x0, x1), x2)).astype(np.int32))
        ymin = np.maximum(0, np.floor(np.minimum(np.minimum(y0, y1), y2)).astype(np.int32))
        ymax = np.minimum(H - 1, np.ceil(np.maximum(np.maximum(y0, y1), y2)).astype(np.int32))
        valid &= (xmax >= xmin) & (ymax >= ymin)
        if not np.any(valid):
            continue

        t_ids = np.nonzero(valid)[0]
        bw = xmax[t_ids] - xmin[t_ids] + 1
        bh = ymax[t_ids] - ymin[t_ids] + 1

        uv = prim["uv"]
        tex = prim["tex"]
        em_tex = prim["em_tex"]
        base_col = prim["base_col"]
        em_col = prim["em_col"]
        metal = prim["metal"]

        # Fast vectorized pass for small/micro triangles (<= 5x5 bbox)
        small_mask = (bw <= 5) & (bh <= 5)
        small_t = t_ids[small_mask]
        if len(small_t) > 0:
            # 5x5 grid offsets evaluated across all small triangles simultaneously in chunks
            gx, gy = np.meshgrid(np.arange(5, dtype=np.int32), np.arange(5, dtype=np.int32))
            gx_f = gx.ravel()[None, :]
            gy_f = gy.ravel()[None, :]
            chunk = 8000
            for c_start in range(0, len(small_t), chunk):
                st = small_t[c_start : c_start + chunk]
                px = xmin[st, None] + gx_f
                py = ymin[st, None] + gy_f
                in_b = (px <= xmax[st, None]) & (py <= ymax[st, None])

                pxc = px.astype(np.float32) + 0.5
                pyc = py.astype(np.float32) + 0.5
                inv_d = (1.0 / denom[st])[:, None]
                w0 = ((y1[st, None] - y2[st, None]) * (pxc - x2[st, None]) + (x2[st, None] - x1[st, None]) * (pyc - y2[st, None])) * inv_d
                w1 = ((y2[st, None] - y0[st, None]) * (pxc - x2[st, None]) + (x0[st, None] - x2[st, None]) * (pyc - y2[st, None])) * inv_d
                w2 = 1.0 - w0 - w1
                hit = in_b & (w0 >= -1e-4) & (w1 >= -1e-4) & (w2 >= -1e-4)
                if not np.any(hit):
                    continue

                ti_loc, gi_loc = np.nonzero(hit)
                tri_g = st[ti_loc]
                px_h = px[ti_loc, gi_loc]
                py_h = py[ti_loc, gi_loc]
                w0_h = w0[ti_loc, gi_loc]
                w1_h = w1[ti_loc, gi_loc]
                w2_h = w2[ti_loc, gi_loc]

                z_h = w0_h * z0[tri_g] + w1_h * z1[tri_g] + w2_h * z2[tri_g]
                # Sort back-to-front so minimum Z wins cleanly on colliding pixels
                ord_z = np.argsort(-z_h)
                tri_g = tri_g[ord_z]
                px_h, py_h, z_h = px_h[ord_z], py_h[ord_z], z_h[ord_z]
                w0_h, w1_h, w2_h = w0_h[ord_z], w1_h[ord_z], w2_h[ord_z]

                closer = z_h < zbuf[py_h, px_h]
                if not np.any(closer):
                    continue
                # Deduplicate to front-most fragment per pixel in this batch
                flat_pix = py_h * W + px_h
                # Since sorted descending by z_h, the LAST occurrence for each flat_pix is the minimum z_h!
                _, rev_idx = np.unique(flat_pix[::-1], return_index=True)
                win = len(flat_pix) - 1 - rev_idx
                win = win[z_h[win] < zbuf[py_h[win], px_h[win]]]
                if len(win) == 0:
                    continue

                tg = tri_g[win]
                pxw, pyw = px_h[win], py_h[win]
                b0, b1, b2 = w0_h[win], w1_h[win], w2_h[win]
                i0, i1, i2 = idx[tg, 0], idx[tg, 1], idx[tg, 2]

                if tex is not None:
                    u_p = b0 * uv[i0, 0] + b1 * uv[i1, 0] + b2 * uv[i2, 0]
                    v_p = b0 * uv[i0, 1] + b1 * uv[i1, 1] + b2 * uv[i2, 1]
                    alb = _sample_bilinear(tex, u_p, v_p) * base_col[None, :]
                else:
                    alb = np.broadcast_to(base_col[None, :], (len(win), 3))

                diff_p = b0[:, None] * v_diff[i0] + b1[:, None] * v_diff[i1] + b2[:, None] * v_diff[i2]
                spec_p = b0[:, None] * v_spec[i0] + b1[:, None] * v_spec[i1] + b2[:, None] * v_spec[i2]
                spec_tint = (1.0 - 0.65 * metal) + 0.65 * metal * alb
                col_lin = alb * diff_p * (1.0 - 0.35 * metal) + spec_p * spec_tint + em_col[None, :]
                if em_tex is not None:
                    u_e = b0 * uv[i0, 0] + b1 * uv[i1, 0] + b2 * uv[i2, 0]
                    v_e = b0 * uv[i0, 1] + b1 * uv[i1, 1] + b2 * uv[i2, 1]
                    col_lin = col_lin + _sample_bilinear(em_tex, u_e, v_e)

                zbuf[pyw, pxw] = z_h[win]
                cbuf[pyw, pxw] = col_lin
                abuf[pyw, pxw] = 255

        # Larger triangles (> 5x5 bbox): exact per-triangle NumPy bounding-box Z-buffer
        large_t = t_ids[~small_mask]
        for t in large_t:
            x_lo, x_hi = int(xmin[t]), int(xmax[t])
            y_lo, y_hi = int(ymin[t]), int(ymax[t])
            ys, xs = np.mgrid[y_lo : y_hi + 1, x_lo : x_hi + 1]
            pxc = xs.astype(np.float32) + 0.5
            pyc = ys.astype(np.float32) + 0.5
            inv_d = 1.0 / float(denom[t])
            w0 = ((y1[t] - y2[t]) * (pxc - x2[t]) + (x2[t] - x1[t]) * (pyc - y2[t])) * inv_d
            w1 = ((y2[t] - y0[t]) * (pxc - x2[t]) + (x0[t] - x2[t]) * (pyc - y2[t])) * inv_d
            w2 = 1.0 - w0 - w1
            inside = (w0 >= -1e-4) & (w1 >= -1e-4) & (w2 >= -1e-4)
            if not np.any(inside):
                continue

            zt = w0 * z0[t] + w1 * z1[t] + w2 * z2[t]
            z_sub = zbuf[y_lo : y_hi + 1, x_lo : x_hi + 1]
            vis = inside & (zt < z_sub)
            if not np.any(vis):
                continue

            b0, b1, b2 = w0[vis], w1[vis], w2[vis]
            i0, i1, i2 = int(idx[t, 0]), int(idx[t, 1]), int(idx[t, 2])
            if tex is not None:
                u_p = b0 * uv[i0, 0] + b1 * uv[i1, 0] + b2 * uv[i2, 0]
                v_p = b0 * uv[i0, 1] + b1 * uv[i1, 1] + b2 * uv[i2, 1]
                alb = _sample_bilinear(tex, u_p, v_p) * base_col[None, :]
            else:
                alb = np.broadcast_to(base_col[None, :], (len(b0), 3))

            diff_p = b0[:, None] * v_diff[i0] + b1[:, None] * v_diff[i1] + b2[:, None] * v_diff[i2]
            spec_p = b0[:, None] * v_spec[i0] + b1[:, None] * v_spec[i1] + b2[:, None] * v_spec[i2]
            spec_tint = (1.0 - 0.65 * metal) + 0.65 * metal * alb
            col_lin = alb * diff_p * (1.0 - 0.35 * metal) + spec_p * spec_tint + em_col[None, :]
            if em_tex is not None:
                u_e = b0 * uv[i0, 0] + b1 * uv[i1, 0] + b2 * uv[i2, 0]
                v_e = b0 * uv[i0, 1] + b1 * uv[i1, 1] + b2 * uv[i2, 1]
                col_lin = col_lin + _sample_bilinear(em_tex, u_e, v_e)

            z_sub[vis] = zt[vis]
            c_sub = cbuf[y_lo : y_hi + 1, x_lo : x_hi + 1]
            a_sub = abuf[y_lo : y_hi + 1, x_lo : x_hi + 1]
            c_sub[vis] = col_lin
            a_sub[vis] = 255

    # Aces-inspired filmic tone curve + sRGB gamma
    mapped = cbuf / (1.0 + cbuf * 0.22)
    srgb = np.power(np.clip(mapped, 0.0, 1.0), 1.0 / 2.2)
    rgba = np.dstack([(srgb * 255.0).astype(np.uint8), abuf])
    out = Image.fromarray(rgba, mode="RGBA")
    if ss > 1:
        out = out.resize((width, height), Image.Resampling.LANCZOS)
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a 2x-SSAA Z-buffered PBR studio preview PNG from any .glb file.")
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
    print(f"Rendered Z-buffered 3D GLB studio preview -> {args.out}")


if __name__ == "__main__":
    main()
