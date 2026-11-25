#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Procedural Multi-Material PBR 3D glTF 2.0 (.glb) Subject Model Generator
for morphdeck-3d.

Generates 100% original, domain-linked 3D hero models with analytical vertex
normals, UV coordinates, 5 PBR metallic-roughness materials, and custom-baked
self-lit textures across 10 archetypes:
  - architectural_pavilion
  - ai_quantum_server
  - energy_turbine_core
  - product_drop_capsule
  - aerospace_orbital_craft
  - executive_globe_monolith
  - biotech_helix_core
  - fintech_vault_prism
  - smartphone_pro  (alias: iphone)
  - laptop_studio   (alias: macbook)
"""

import argparse
import io
import json
import math
import os
import struct
import sys
import numpy as np
from PIL import Image, ImageDraw, ImageFont


ARCHETYPES = [
    "architectural_pavilion",
    "ai_quantum_server",
    "energy_turbine_core",
    "product_drop_capsule",
    "aerospace_orbital_craft",
    "executive_globe_monolith",
    "biotech_helix_core",
    "fintech_vault_prism",
    "smartphone_pro",
    "laptop_studio",
]

ARCHETYPE_ALIASES = {
    "iphone": "smartphone_pro",
    "macbook": "laptop_studio",
}


def canonical_archetype(name: str) -> str:
    key = (name or "").strip().lower()
    return ARCHETYPE_ALIASES.get(key, key)


def hex_to_rgb(hex_str: str, default=(37, 99, 235)) -> tuple:
    if not hex_str:
        return default
    h = str(hex_str).strip().lstrip("#")
    if len(h) != 6:
        return default
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def srgb_to_linear(rgb_tuple: tuple) -> list:
    return [((max(0, min(255, c)) / 255.0) ** 2.2) for c in rgb_tuple]


def load_font(size: int, bold: bool = True):
    """Cross-platform TrueType font loader with safe fallback."""
    candidates = []
    if sys.platform.startswith("win"):
        fdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        names = ["bahnschrift.ttf", "segoeuib.ttf", "arialbd.ttf"] if bold else ["consola.ttf", "segoeui.ttf", "arial.ttf"]
        candidates.extend(os.path.join(fdir, n) for n in names)
    elif sys.platform == "darwin":
        candidates.extend([
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
            "/Library/Fonts/Arial.ttf",
        ])
    else:
        candidates.extend([
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ])

    for p in candidates:
        if os.path.exists(p):
            try:
                f = ImageFont.truetype(p, size)
                if "bahnschrift" in p.lower():
                    try:
                        f.set_variation_by_axes([700 if bold else 400, 100])
                    except Exception:
                        pass
                return f
            except Exception:
                continue
    return ImageFont.load_default()


def rotate_points(pts: np.ndarray, rx_deg: float = 0.0, ry_deg: float = 0.0, rz_deg: float = 0.0) -> np.ndarray:
    out = pts.copy()
    if rx_deg != 0.0:
        a = math.radians(rx_deg)
        c, s = math.cos(a), math.sin(a)
        R = np.array([[1, 0, 0], [0, c, -s], [0, s, c]], dtype=np.float32)
        out = out @ R.T
    if ry_deg != 0.0:
        a = math.radians(ry_deg)
        c, s = math.cos(a), math.sin(a)
        R = np.array([[c, 0, s], [0, 1, 0], [-s, 0, c]], dtype=np.float32)
        out = out @ R.T
    if rz_deg != 0.0:
        a = math.radians(rz_deg)
        c, s = math.cos(a), math.sin(a)
        R = np.array([[c, -s, 0], [s, c, 0], [0, 0, 1]], dtype=np.float32)
        out = out @ R.T
    return out


class PrimBuilder:
    """Collects indexed triangles (positions, normals, UVs) for one glTF material."""

    def __init__(self, material_idx: int, name: str):
        self.material_idx = material_idx
        self.name = name
        self.positions = []
        self.normals = []
        self.uvs = []
        self.indices = []
        self._vcount = 0

    def _append_geom(self, pos: np.ndarray, norm: np.ndarray, uv: np.ndarray, idx: np.ndarray):
        norms = np.linalg.norm(norm, axis=1, keepdims=True)
        norms = np.where(norms < 1e-6, 1.0, norms)
        norm = (norm / norms).astype(np.float32)
        self.positions.append(pos.astype(np.float32))
        self.normals.append(norm)
        self.uvs.append(uv.astype(np.float32))
        self.indices.append((idx.astype(np.uint32) + self._vcount))
        self._vcount += len(pos)

    def add_quad(self, center=(0, 0, 0), size=(1.0, 1.0), rot=(0, 0, 0), uv_rect=(0.0, 0.0, 1.0, 1.0)):
        """Single forward-facing (+Z) quad with upright [0..1] planar UVs."""
        hx, hy = size[0] * 0.5, size[1] * 0.5
        u0, v0, u1, v1 = uv_rect
        verts = [(-hx, -hy, 0.0), (hx, -hy, 0.0), (hx, hy, 0.0), (-hx, hy, 0.0)]
        pos = rotate_points(np.array(verts, dtype=np.float32), *rot) + np.array(center, dtype=np.float32)
        norm = rotate_points(np.array([(0.0, 0.0, 1.0)] * 4, dtype=np.float32), *rot)
        uv = np.array([(u0, v1), (u1, v1), (u1, v0), (u0, v0)], dtype=np.float32)
        idx = np.array([0, 1, 2, 0, 2, 3], dtype=np.uint32)
        self._append_geom(pos, norm, uv, idx)

    def add_box(self, center=(0, 0, 0), size=(1, 1, 1), rot=(0, 0, 0), uv_rect=(0.0, 0.0, 1.0, 1.0)):
        hx, hy, hz = size[0] * 0.5, size[1] * 0.5, size[2] * 0.5
        u0, v0, u1, v1 = uv_rect
        faces = [
            ([(-hx, -hy, hz), (hx, -hy, hz), (hx, hy, hz), (-hx, hy, hz)], (0, 0, 1)),
            ([(hx, -hy, -hz), (-hx, -hy, -hz), (-hx, hy, -hz), (hx, hy, -hz)], (0, 0, -1)),
            ([(hx, -hy, hz), (hx, -hy, -hz), (hx, hy, -hz), (hx, hy, hz)], (1, 0, 0)),
            ([(-hx, -hy, -hz), (-hx, -hy, hz), (-hx, hy, hz), (-hx, hy, -hz)], (-1, 0, 0)),
            ([(-hx, hy, hz), (hx, hy, hz), (hx, hy, -hz), (-hx, hy, -hz)], (0, 1, 0)),
            ([(-hx, -hy, -hz), (hx, -hy, -hz), (hx, -hy, hz), (-hx, -hy, hz)], (0, -1, 0)),
        ]
        pos_list, norm_list, uv_list, idx_list = [], [], [], []
        for fi, (verts, n) in enumerate(faces):
            base = fi * 4
            pos_list.extend(verts)
            norm_list.extend([n] * 4)
            uv_list.extend([(u0, v1), (u1, v1), (u1, v0), (u0, v0)])
            idx_list.extend([base, base + 1, base + 2, base, base + 2, base + 3])

        pos = rotate_points(np.array(pos_list, dtype=np.float32), *rot) + np.array(center, dtype=np.float32)
        norm = rotate_points(np.array(norm_list, dtype=np.float32), *rot)
        uv = np.array(uv_list, dtype=np.float32)
        idx = np.array(idx_list, dtype=np.uint32)
        self._append_geom(pos, norm, uv, idx)

    def add_cylinder(self, center=(0, 0, 0), r_bottom=0.5, r_top=0.5, height=1.0, segments=48, rot=(0, 0, 0), capped=True):
        hy = height * 0.5
        pos_list, norm_list, uv_list, idx_list = [], [], [], []
        slope_y = (r_bottom - r_top) / max(1e-5, height)

        for i in range(segments + 1):
            frac = i / segments
            th = (1.0 - frac * 2.0) * math.pi
            ct, st = math.cos(th), math.sin(th)
            pos_list.append((r_bottom * ct, -hy, r_bottom * st))
            pos_list.append((r_top * ct, hy, r_top * st))
            norm_list.append((ct, slope_y, st))
            norm_list.append((ct, slope_y, st))
            uv_list.append((frac, 1.0))
            uv_list.append((frac, 0.0))

        for i in range(segments):
            b = i * 2
            idx_list.extend([b, b + 2, b + 1, b + 2, b + 3, b + 1])

        if capped:
            top_c_idx = len(pos_list)
            pos_list.append((0.0, hy, 0.0))
            norm_list.append((0.0, 1.0, 0.0))
            uv_list.append((0.5, 0.5))
            for i in range(segments + 1):
                th = (i / segments) * 2.0 * math.pi
                ct, st = math.cos(th), math.sin(th)
                pos_list.append((r_top * ct, hy, r_top * st))
                norm_list.append((0.0, 1.0, 0.0))
                uv_list.append((0.5 + 0.5 * ct, 0.5 + 0.5 * st))
            for i in range(segments):
                idx_list.extend([top_c_idx, top_c_idx + 2 + i, top_c_idx + 1 + i])

            bot_c_idx = len(pos_list)
            pos_list.append((0.0, -hy, 0.0))
            norm_list.append((0.0, -1.0, 0.0))
            uv_list.append((0.5, 0.5))
            for i in range(segments + 1):
                th = (i / segments) * 2.0 * math.pi
                ct, st = math.cos(th), math.sin(th)
                pos_list.append((r_bottom * ct, -hy, r_bottom * st))
                norm_list.append((0.0, -1.0, 0.0))
                uv_list.append((0.5 + 0.5 * ct, 0.5 + 0.5 * st))
            for i in range(segments):
                idx_list.extend([bot_c_idx, bot_c_idx + 1 + i, bot_c_idx + 2 + i])

        pos = rotate_points(np.array(pos_list, dtype=np.float32), *rot) + np.array(center, dtype=np.float32)
        norm = rotate_points(np.array(norm_list, dtype=np.float32), *rot)
        self._append_geom(pos, norm, np.array(uv_list, dtype=np.float32), np.array(idx_list, dtype=np.uint32))

    def add_torus(self, center=(0, 0, 0), major_r=1.0, minor_r=0.06, major_seg=64, minor_seg=24, rot=(0, 0, 0)):
        pos_list, norm_list, uv_list, idx_list = [], [], [], []
        for i in range(major_seg + 1):
            u = i / major_seg
            th = u * 2.0 * math.pi
            ct, st = math.cos(th), math.sin(th)
            for j in range(minor_seg + 1):
                v = j / minor_seg
                ph = v * 2.0 * math.pi
                cp, sp = math.cos(ph), math.sin(ph)
                r = major_r + minor_r * cp
                pos_list.append((r * ct, minor_r * sp, r * st))
                norm_list.append((cp * ct, sp, cp * st))
                uv_list.append((u, v))

        stride = minor_seg + 1
        for i in range(major_seg):
            for j in range(minor_seg):
                a = i * stride + j
                b = (i + 1) * stride + j
                idx_list.extend([a, a + 1, b, b, a + 1, b + 1])

        pos = rotate_points(np.array(pos_list, dtype=np.float32), *rot) + np.array(center, dtype=np.float32)
        norm = rotate_points(np.array(norm_list, dtype=np.float32), *rot)
        self._append_geom(pos, norm, np.array(uv_list, dtype=np.float32), np.array(idx_list, dtype=np.uint32))

    def add_sphere(self, center=(0, 0, 0), radius=0.8, lat_seg=32, lon_seg=48, rot=(0, 0, 0)):
        pos_list, norm_list, uv_list, idx_list = [], [], [], []
        for i in range(lat_seg + 1):
            v = i / lat_seg
            phi = v * math.pi
            sp, cp = math.sin(phi), math.cos(phi)
            for j in range(lon_seg + 1):
                u = j / lon_seg
                th = u * 2.0 * math.pi
                ct, st = math.cos(th), math.sin(th)
                nx, ny, nz = sp * ct, cp, sp * st
                pos_list.append((radius * nx, radius * ny, radius * nz))
                norm_list.append((nx, ny, nz))
                uv_list.append((u, v))

        stride = lon_seg + 1
        for i in range(lat_seg):
            for j in range(lon_seg):
                a = i * stride + j
                b = (i + 1) * stride + j
                idx_list.extend([a, b, a + 1, a + 1, b, b + 1])

        pos = rotate_points(np.array(pos_list, dtype=np.float32), *rot) + np.array(center, dtype=np.float32)
        norm = rotate_points(np.array(norm_list, dtype=np.float32), *rot)
        self._append_geom(pos, norm, np.array(uv_list, dtype=np.float32), np.array(idx_list, dtype=np.uint32))


def build_core_texture(archetype: str, brand_name: str, acc1: tuple, acc2: tuple, acc3: tuple) -> bytes:
    """Bake a crisp 1024x1024 self-lit PNG texture tailored to the 3D subject archetype."""
    W, H = 1024, 1024
    img = Image.new("RGB", (W, H), (10, 14, 22))
    d = ImageDraw.Draw(img)
    f_big = load_font(64, True)
    f_med = load_font(34, True)
    f_mono = load_font(24, False)

    if archetype == "architectural_pavilion":
        yy, _ = np.mgrid[0:H, 0:W]
        warm_top = np.array([255, 214, 155], dtype=np.float32)
        warm_bot = np.array([185, 95, 45], dtype=np.float32)
        t = (yy / float(H))[..., None]
        arr = warm_top * (1.0 - t) + warm_bot * t
        img = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), "RGB")
        d = ImageDraw.Draw(img)
        d.rectangle([0, 460, W, 505], fill=(42, 32, 26))
        d.rectangle([0, 880, W, 1024], fill=(34, 28, 24))
        for mx in range(64, W, 128):
            d.rectangle([mx - 6, 0, mx + 6, H], fill=(38, 30, 25))
        for lx in (192, 448, 576, 832):
            d.ellipse([lx - 26, 170, lx + 26, 222], fill=(255, 248, 225))
            d.ellipse([lx - 22, 610, lx + 22, 654], fill=(255, 242, 210))

    elif archetype == "energy_turbine_core":
        img = Image.new("RGB", (W, H), (12, 24, 48))
        d = ImageDraw.Draw(img)
        for gx in range(0, W, 64):
            d.line([(gx, 0), (gx, H)], fill=acc1, width=2)
        for gy in range(0, H, 64):
            d.line([(0, gy), (W, gy)], fill=acc2, width=2)
        for gx in range(128, W, 256):
            d.rectangle([gx - 4, 0, gx + 4, H], fill=(210, 230, 255))
        d.rounded_rectangle([140, 360, W - 140, 664], radius=36, fill=(10, 16, 28), outline=acc2, width=6)
        d.text((200, 420), brand_name.upper()[:14], font=f_big, fill=(255, 255, 255))
        d.text((204, 520), "MICRO-GRID // 99.9% UPTIME", font=f_med, fill=acc2)

    elif archetype == "executive_globe_monolith":
        img = Image.new("RGB", (W, H), (8, 16, 34))
        d = ImageDraw.Draw(img)
        for gx in range(0, W, 64):
            d.line([(gx, 0), (gx, H)], fill=(24, 52, 96), width=2)
        for gy in range(0, H, 64):
            d.line([(0, gy), (W, gy)], fill=(24, 52, 96), width=2)
        rng = np.random.default_rng(17)
        pts = [(int(rng.integers(80, W - 80)), int(rng.integers(140, H - 140))) for _ in range(44)]
        for i in range(len(pts) - 1):
            d.line([pts[i], pts[i + 1]], fill=acc1, width=3)
        for px, py in pts:
            d.ellipse([px - 12, py - 12, px + 12, py + 12], fill=acc2)
            d.ellipse([px - 5, py - 5, px + 5, py + 5], fill=(255, 255, 255))
        d.rectangle([0, 468, W, 556], fill=(10, 20, 40))
        d.text((120, 482), f"{brand_name.upper()[:14]}  //  GLOBAL NET", font=f_med, fill=(255, 255, 255))

    elif archetype == "biotech_helix_core":
        img = Image.new("RGB", (W, H), (6, 22, 18))
        d = ImageDraw.Draw(img)
        for gx in range(0, W, 64):
            d.line([(gx, 0), (gx, H)], fill=(14, 48, 40), width=2)
            d.line([(0, gx), (W, gx)], fill=(14, 48, 40), width=2)
        d.rectangle([32, 320, W - 32, 700], fill=(10, 32, 26), outline=acc1, width=6)
        d.text((96, 400), brand_name.upper()[:14], font=f_big, fill=(240, 253, 250))
        d.text((96, 510), "GENOMIC SEQUENCE // TARGET LOCKED", font=f_med, fill=acc1)

    elif archetype == "fintech_vault_prism":
        img = Image.new("RGB", (W, H), (10, 12, 18))
        d = ImageDraw.Draw(img)
        for gy in range(0, H, 48):
            d.line([(0, gy), (W, gy)], fill=(28, 32, 46), width=2)
        for gx in range(0, W, 48):
            d.line([(gx, 0), (gx, H)], fill=(28, 32, 46), width=2)
        d.rounded_rectangle([90, 280, W - 90, 740], radius=32, fill=(14, 18, 28), outline=acc1, width=6)
        d.text((140, 360), brand_name.upper()[:14], font=f_big, fill=(255, 255, 255))
        d.text((140, 470), "ZERO-KNOWLEDGE LEDGER CORE", font=f_med, fill=acc1)
        d.text((140, 560), "HARDWARE ENCLAVE // AES-GCM-256", font=f_mono, fill=acc2)

    elif archetype in ("smartphone_pro", "laptop_studio"):
        img = Image.new("RGB", (W, H), (11, 15, 25))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([40, 40, W - 40, H - 40], radius=48, fill=(15, 21, 35), outline=acc1, width=6)
        d.rectangle([40, 140, W - 40, 290], fill=acc1)
        d.text((88, 182), brand_name.upper()[:14], font=f_big, fill=(255, 255, 255))
        d.rounded_rectangle([84, 348, W - 84, 590], radius=24, fill=(10, 14, 24), outline=acc2, width=4)
        d.text((120, 396), "LIVE PLATFORM INTERFACE", font=f_med, fill=acc2)
        d.text((120, 468), "REAL-TIME TELEMETRY // 99.98% SLA", font=f_mono, fill=(226, 232, 240))
        for bi in range(4):
            bx = 84 + bi * 224
            d.rounded_rectangle([bx, 646, bx + 190, 910], radius=22, fill=(22, 30, 48), outline=acc1 if bi % 2 == 0 else acc2, width=3)

    else:
        img = Image.new("RGB", (W, H), (9, 12, 20))
        d = ImageDraw.Draw(img)
        for g in range(0, W, 64):
            d.line([(g, 0), (g, H)], fill=(22, 30, 48), width=2)
            d.line([(0, g), (W, g)], fill=(22, 30, 48), width=2)
        d.rectangle([60, 60, W - 60, H - 60], outline=acc1, width=6)
        d.rectangle([60, 180, W - 60, 340], fill=acc1)
        d.text((110, 222), brand_name.upper()[:14], font=f_big, fill=(255, 255, 255))
        d.rounded_rectangle([100, 420, W - 100, 640], radius=28, fill=(16, 22, 36), outline=acc2, width=4)
        d.text((144, 465), "CORE TELEMETRY // ACTIVE", font=f_med, fill=acc2)
        d.text((144, 535), "HIGH-THROUGHPUT ENGINE", font=f_mono, fill=(226, 232, 240))
        for bi in range(8):
            bx = 110 + bi * 104
            d.rounded_rectangle([bx, 710, bx + 76, 900], radius=12, fill=acc1 if bi % 2 == 0 else acc2)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def generate_subject_glb(
    archetype: str,
    out_glb: str,
    brand_name: str = "PROJECT",
    accent1_hex: str = "#2563EB",
    accent2_hex: str = "#10B981",
    accent3_hex: str = "#F59E0B",
    screen_png_path: str = "",
) -> str:
    """Build a valid binary glTF 2.0 (.glb) 3D model for any of the 10 archetypes."""
    archetype = canonical_archetype(archetype)
    acc1 = hex_to_rgb(accent1_hex, (37, 99, 235))
    acc2 = hex_to_rgb(accent2_hex, (16, 185, 129))
    acc3 = hex_to_rgb(accent3_hex, (245, 158, 11))

    acc1_lin = srgb_to_linear(acc1)
    acc2_lin = srgb_to_linear(acc2)
    acc3_lin = srgb_to_linear(acc3)

    p0 = PrimBuilder(0, "mat_primary")
    p1 = PrimBuilder(1, "mat_dark_metal")
    p2 = PrimBuilder(2, "mat_secondary")
    p3 = PrimBuilder(3, "mat_textured_core")
    p4 = PrimBuilder(4, "mat_emissive_glow")

    if archetype == "architectural_pavilion":
        mat0_col = srgb_to_linear((234, 224, 208))
        mat0_metal, mat0_rough = 0.08, 0.55
        mat1_col = srgb_to_linear((38, 32, 28))
        mat2_col = srgb_to_linear((168, 98, 52))
        mat2_metal, mat2_rough = 0.12, 0.45

        p0.add_box((0.0, -0.78, 0.0), (2.55, 0.16, 1.85))
        p4.add_box((0.0, -0.68, 0.0), (2.42, 0.04, 1.72))
        p0.add_box((0.0, -0.60, 0.0), (2.36, 0.14, 1.66))

        p2.add_box((-0.78, -0.02, -0.10), (0.48, 1.04, 1.24))
        p0.add_box((0.76, -0.02, -0.24), (0.52, 1.04, 0.96))
        p3.add_box((0.04, -0.02, 0.06), (1.18, 1.02, 1.02))

        for mx in (-0.54, -0.16, 0.22, 0.60):
            p1.add_box((mx, -0.02, 0.58), (0.04, 1.04, 0.05))
        p1.add_box((0.04, -0.02, 0.58), (1.20, 0.05, 0.06))

        for li in range(14):
            lx = -0.98 + li * 0.15
            p2.add_box((lx, -0.02, 0.72), (0.035, 1.04, 0.11), rot=(0, 18, 0))

        for cx in (-0.82, -0.28, 0.26, 0.80):
            p1.add_cylinder((cx, -0.02, 0.78), 0.042, 0.042, 1.04, segments=28)
            p4.add_torus((cx, 0.45, 0.78), 0.055, 0.014)
        p2.add_cylinder((-0.34, -0.02, 0.18), 0.19, 0.19, 1.00, segments=32)

        p0.add_torus((0.54, -0.50, 0.56), 0.34, 0.042)
        p4.add_cylinder((0.54, -0.51, 0.56), 0.31, 0.31, 0.025, segments=36)

        p0.add_box((0.0, 0.56, 0.04), (2.68, 0.14, 1.92))
        p1.add_box((0.0, 0.48, 0.04), (2.54, 0.04, 1.78))
        p2.add_box((-0.15, 0.68, -0.05), (1.45, 0.12, 1.05))
        p1.add_torus((0.04, 0.52, 0.10), 0.42, 0.036)
        p4.add_torus((0.04, 0.48, 0.10), 0.36, 0.022)
        p4.add_box((0.04, 0.46, 0.42), (1.50, 0.025, 0.08))

    elif archetype == "energy_turbine_core":
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.78, 0.22
        mat1_col = srgb_to_linear((18, 22, 32))
        mat2_col = acc2_lin
        mat2_metal, mat2_rough = 0.65, 0.25

        p1.add_cylinder((0.0, -0.95, 0.0), 1.15, 1.05, 0.22, segments=48)
        p4.add_torus((0.0, -0.82, 0.0), 1.04, 0.035)
        p1.add_cylinder((0.0, -0.70, 0.0), 0.85, 0.78, 0.24, segments=48)

        p3.add_cylinder((0.0, -0.15, 0.0), 0.44, 0.44, 0.92, segments=48)
        for ry in (-0.45, -0.15, 0.15):
            p0.add_torus((0.0, ry, 0.0), 0.47, 0.032)

        p3.add_box((-0.92, -0.18, 0.0), (0.86, 0.62, 0.05), rot=(28, -12, 8))
        p1.add_box((-0.92, -0.18, -0.03), (0.90, 0.66, 0.03), rot=(28, -12, 8))
        p3.add_box((0.92, -0.18, 0.0), (0.86, 0.62, 0.05), rot=(28, 12, -8))
        p1.add_box((0.92, -0.18, -0.03), (0.90, 0.66, 0.03), rot=(28, 12, -8))

        p0.add_cylinder((0.0, 0.56, 0.0), 0.11, 0.08, 0.65, segments=32)
        p0.add_cylinder((0.0, 0.88, 0.08), 0.15, 0.12, 0.46, segments=32, rot=(90, 0, 0))
        p4.add_sphere((0.0, 0.88, 0.32), 0.12)
        for bi in range(3):
            ang = bi * 120.0 + 18.0
            rad = math.radians(ang)
            bx = 0.48 * math.cos(rad)
            by = 0.88 + 0.48 * math.sin(rad)
            p2.add_box((bx, by, 0.28), (0.88, 0.11, 0.03), rot=(14, 0, ang))

    elif archetype == "product_drop_capsule":
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.82, 0.18
        mat1_col = srgb_to_linear((16, 17, 22))
        mat2_col = acc2_lin
        mat2_metal, mat2_rough = 0.75, 0.20

        p1.add_cylinder((0.0, -0.92, 0.0), 1.02, 0.90, 0.22, segments=56)
        p4.add_torus((0.0, -0.79, 0.0), 0.91, 0.035)
        p0.add_cylinder((0.0, -0.66, 0.0), 0.72, 0.68, 0.14, segments=56)

        p3.add_cylinder((0.0, 0.04, 0.0), 0.56, 0.56, 1.12, segments=64)
        p0.add_sphere((0.0, 0.60, 0.0), 0.56)
        p1.add_torus((0.0, 0.60, 0.0), 0.57, 0.04)
        p1.add_torus((0.0, -0.52, 0.0), 0.57, 0.04)

        p2.add_torus((0.0, 0.06, 0.0), 0.88, 0.038, rot=(24, 18, -22))
        p0.add_torus((0.0, 0.06, 0.0), 0.96, 0.028, rot=(-20, -14, 28))
        p4.add_torus((0.0, 0.28, 0.0), 0.58, 0.024)

    elif archetype == "aerospace_orbital_craft":
        mat0_col = srgb_to_linear((228, 234, 242))
        mat0_metal, mat0_rough = 0.75, 0.22
        mat1_col = srgb_to_linear((20, 24, 32))
        mat2_col = acc1_lin
        mat2_metal, mat2_rough = 0.80, 0.20

        p1.add_torus((0.0, -1.02, 0.0), 0.78, 0.05)
        for ex, ez in ((-0.18, -0.15), (0.18, -0.15), (0.0, 0.20)):
            p1.add_cylinder((ex, -0.88, ez), 0.16, 0.09, 0.26, segments=32)
            p4.add_sphere((ex, -0.96, ez), 0.11)
        p0.add_cylinder((0.0, -0.32, 0.0), 0.42, 0.42, 0.92, segments=56)
        p1.add_cylinder((0.0, 0.20, 0.0), 0.43, 0.43, 0.16, segments=56)
        p3.add_cylinder((0.0, 0.56, 0.0), 0.42, 0.36, 0.58, segments=56)
        p0.add_cylinder((0.0, 1.04, 0.0), 0.36, 0.04, 0.42, segments=56)
        p2.add_sphere((0.0, 1.25, 0.0), 0.05)
        for fi in range(4):
            ang = fi * 90.0 + 45.0
            rad = math.radians(ang)
            fx, fz = 0.52 * math.cos(rad), 0.52 * math.sin(rad)
            p2.add_box((fx, -0.52, fz), (0.38, 0.52, 0.04), rot=(0, -ang, -14))

    elif archetype == "executive_globe_monolith":
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.84, 0.18
        mat1_col = srgb_to_linear((16, 22, 36))
        mat2_col = acc3_lin
        mat2_metal, mat2_rough = 0.88, 0.16

        p1.add_cylinder((0.0, -0.92, 0.0), 0.96, 0.86, 0.20, segments=56)
        p0.add_cylinder((0.0, -0.74, 0.0), 0.68, 0.52, 0.20, segments=56)
        p4.add_torus((0.0, -0.62, 0.0), 0.54, 0.028)

        p3.add_sphere((0.0, 0.12, 0.0), 0.66, lat_seg=36, lon_seg=56, rot=(12, -25, 0))
        p0.add_torus((0.0, 0.12, 0.0), 0.80, 0.032, rot=(18, 0, 14))
        p2.add_torus((0.0, 0.12, 0.0), 0.88, 0.026, rot=(74, 22, 0))
        p4.add_torus((0.0, 0.12, 0.0), 0.72, 0.018, rot=(-24, 15, 0))

    elif archetype == "biotech_helix_core":
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.72, 0.22
        mat1_col = srgb_to_linear((12, 24, 22))
        mat2_col = acc2_lin
        mat2_metal, mat2_rough = 0.68, 0.24

        # Containment bio-pedestal & crown rings
        p1.add_cylinder((0.0, -0.96, 0.0), 0.96, 0.86, 0.18, segments=48)
        p3.add_cylinder((0.0, -0.74, 0.0), 0.72, 0.72, 0.26, segments=48)
        p4.add_torus((0.0, -0.60, 0.0), 0.74, 0.028)
        p1.add_torus((0.0, 0.86, 0.0), 0.66, 0.045)
        p4.add_torus((0.0, 0.86, 0.0), 0.54, 0.022)

        # Double-helix backbone strands & base-pair cross-rungs
        steps = 20
        r_hlx = 0.40
        for hi in range(steps):
            frac = hi / max(1, steps - 1)
            hy = -0.46 + frac * 1.24
            th = frac * 2.6 * math.pi
            deg = math.degrees(th)
            hx = r_hlx * math.cos(th)
            hz = r_hlx * math.sin(th)
            p0.add_sphere((hx, hy, hz), 0.068, lat_seg=14, lon_seg=20)
            p2.add_sphere((-hx, hy, -hz), 0.068, lat_seg=14, lon_seg=20)
            target = p4 if hi % 3 == 0 else p1
            target.add_box((0.0, hy, 0.0), (r_hlx * 1.92, 0.036, 0.036), rot=(0, -deg, 0))

        # Dual orbiting biosensor gimbal rings
        p0.add_torus((0.0, 0.18, 0.0), 0.78, 0.024, rot=(22, 16, -28))
        p2.add_torus((0.0, 0.18, 0.0), 0.84, 0.020, rot=(-24, -18, 22))

    elif archetype == "fintech_vault_prism":
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.88, 0.15
        mat1_col = srgb_to_linear((16, 18, 26))
        mat2_col = srgb_to_linear((220, 226, 236))
        mat2_metal, mat2_rough = 0.90, 0.14

        # Security plinth
        p1.add_cylinder((0.0, -0.92, 0.0), 1.02, 0.90, 0.20, segments=6)
        p4.add_torus((0.0, -0.80, 0.0), 0.90, 0.032)
        p2.add_cylinder((0.0, -0.70, 0.0), 0.76, 0.70, 0.16, segments=6)

        # Hexagonal encrypted core prism & corner armor ribs
        p3.add_cylinder((0.0, 0.08, 0.0), 0.56, 0.56, 1.18, segments=6)
        p0.add_cylinder((0.0, 0.74, 0.0), 0.56, 0.16, 0.22, segments=6)
        for ci in range(6):
            ang = math.radians(ci * 60.0)
            px, pz = 0.58 * math.cos(ang), 0.58 * math.sin(ang)
            p1.add_box((px, 0.08, pz), (0.09, 1.22, 0.09), rot=(0, -ci * 60.0, 0))

        # Tri-axis armored vault rings
        p0.add_torus((0.0, 0.08, 0.0), 0.88, 0.038, rot=(20, 12, 18))
        p2.add_torus((0.0, 0.08, 0.0), 0.96, 0.030, rot=(-26, 18, -22))
        p4.add_torus((0.0, 0.08, 0.0), 0.80, 0.022, rot=(68, -24, 12))

    elif archetype == "smartphone_pro":
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.86, 0.20
        mat1_col = srgb_to_linear((12, 14, 20))
        mat2_col = srgb_to_linear((210, 218, 230))
        mat2_metal, mat2_rough = 0.88, 0.16

        # Anodized smartphone body with rounded tubular side/corner rails
        bw, bh, bd = 1.04, 2.12, 0.11
        cr = 0.065
        p0.add_box((0.0, 0.0, 0.0), (bw - cr * 1.4, bh, bd))
        p0.add_box((0.0, 0.0, 0.0), (bw, bh - cr * 1.4, bd))
        for cx in (-(bw * 0.5 - cr), (bw * 0.5 - cr)):
            p0.add_cylinder((cx, 0.0, 0.0), cr, cr, bh - cr * 1.2, segments=24)
        for cy in (-(bh * 0.5 - cr), (bh * 0.5 - cr)):
            p0.add_cylinder((0.0, cy, 0.0), cr, cr, bw - cr * 1.2, segments=24, rot=(0, 0, 90))

        # Side titanium buttons
        p2.add_box((bw * 0.5 + 0.015, 0.34, 0.0), (0.024, 0.24, 0.04))
        p2.add_box((-bw * 0.5 - 0.015, 0.42, 0.0), (0.024, 0.16, 0.04))
        p2.add_box((-bw * 0.5 - 0.015, 0.18, 0.0), (0.024, 0.16, 0.04))

        # Front OLED glass surround (cleanly inset inside rounded titanium perimeter) & Dynamic Island notch
        p1.add_box((0.0, 0.0, 0.054), (bw - 0.08, bh - 0.08, 0.012))
        p3.add_quad((0.0, 0.0, 0.064), (bw - 0.11, bh - 0.11))
        p1.add_box((0.0, bh * 0.5 - 0.13, 0.066), (0.26, 0.062, 0.008))

        # Rear triple-lens camera island
        p1.add_box((-0.22, 0.66, -0.068), (0.42, 0.46, 0.024))
        for lx, ly in ((-0.32, 0.76), (-0.32, 0.56), (-0.12, 0.66)):
            p2.add_cylinder((lx, ly, -0.088), 0.082, 0.082, 0.028, segments=24, rot=(90, 0, 0))
            p4.add_cylinder((lx, ly, -0.095), 0.052, 0.052, 0.016, segments=20, rot=(90, 0, 0))

    elif archetype == "laptop_studio":
        mat0_col = srgb_to_linear((132, 138, 152))
        mat0_metal, mat0_rough = 0.86, 0.28
        mat1_col = srgb_to_linear((16, 18, 26))
        mat2_col = acc1_lin
        mat2_metal, mat2_rough = 0.78, 0.24

        # Unibody base chassis (tilted slightly forward visually, center y=-0.42)
        p0.add_box((0.0, -0.44, 0.26), (1.88, 0.075, 1.24))
        p1.add_box((0.0, -0.398, 0.12), (1.58, 0.014, 0.58))
        # Individual keycaps grid (5 rows x 12 cols)
        for kr in range(5):
            kz = -0.09 + kr * 0.105
            for kc in range(12):
                kx = -0.68 + kc * 0.124
                p1.add_box((kx, -0.386, kz), (0.096, 0.018, 0.082))
        # Glass trackpad & accent strip
        p1.add_box((0.0, -0.396, 0.62), (0.64, 0.012, 0.36))
        p4.add_box((0.0, -0.44, 0.885), (0.36, 0.018, 0.012))

        # Hinge barrel
        p1.add_cylinder((0.0, -0.38, -0.36), 0.052, 0.052, 1.56, segments=24, rot=(0, 0, 90))

        # Display lid tilted slightly back (-14 deg around X)
        lid_c = (0.0, 0.18, -0.52)
        lid_rot = (-14.0, 0.0, 0.0)
        p0.add_box(lid_c, (1.88, 1.20, 0.044), rot=lid_rot)
        # Front bezel + raised display quad (offset along tilted lid normal +Z)
        a = math.radians(-14.0)
        nz_y = -math.sin(a)
        nz_z = math.cos(a)
        bezel_c = (0.0, lid_c[1] + 0.024 * nz_y, lid_c[2] + 0.024 * nz_z)
        scr_c = (0.0, lid_c[1] + 0.030 * nz_y, lid_c[2] + 0.030 * nz_z)
        p1.add_box(bezel_c, (1.82, 1.14, 0.010), rot=lid_rot)
        p3.add_quad(scr_c, (1.72, 1.06), rot=lid_rot)
        # Accent emblem on outer lid back
        back_c = (0.0, lid_c[1] - 0.026 * nz_y, lid_c[2] - 0.026 * nz_z)
        p2.add_cylinder(back_c, 0.11, 0.11, 0.010, segments=32, rot=(90 - 14, 0, 0))

    else:
        # ai_quantum_server
        mat0_col = acc1_lin
        mat0_metal, mat0_rough = 0.85, 0.18
        mat1_col = srgb_to_linear((15, 20, 32))
        mat2_col = acc2_lin
        mat2_metal, mat2_rough = 0.78, 0.22

        p1.add_cylinder((0.0, -0.92, 0.0), 0.94, 0.88, 0.20, segments=48)
        p4.add_torus((0.0, -0.80, 0.0), 0.88, 0.032)
        p1.add_cylinder((0.0, 0.92, 0.0), 0.88, 0.94, 0.20, segments=48)
        p4.add_torus((0.0, 0.80, 0.0), 0.88, 0.032)

        for px, pz in ((-0.56, -0.56), (0.56, -0.56), (0.56, 0.56), (-0.56, 0.56)):
            p1.add_box((px, 0.0, pz), (0.14, 1.66, 0.14))
            p4.add_box((px * 1.08, 0.0, pz * 1.08), (0.035, 1.52, 0.035))

        p3.add_box((0.0, 0.0, 0.0), (0.86, 1.44, 0.86))
        for bi, by in enumerate(np.linspace(-0.60, 0.60, 7)):
            target = p0 if bi % 2 == 0 else p2
            target.add_box((0.0, float(by), 0.0), (1.16, 0.065, 1.16))

        p0.add_torus((0.0, 0.0, 0.0), 1.08, 0.032, rot=(18, 20, -16))
        p2.add_torus((0.0, 0.0, 0.0), 1.02, 0.024, rot=(-22, -15, 24))

    # Pack binary glTF 2.0 buffers
    bin_data = bytearray()
    buffer_views = []
    accessors = []
    mesh_primitives = []

    for pb in (p0, p1, p2, p3, p4):
        if not pb.positions:
            continue
        pos = np.vstack(pb.positions).astype("<f4")
        norm = np.vstack(pb.normals).astype("<f4")
        uv = np.vstack(pb.uvs).astype("<f4")
        idx = np.concatenate(pb.indices).astype("<u4")

        def add_bv(raw_bytes: bytes, target: int) -> int:
            while len(bin_data) % 4 != 0:
                bin_data.append(0)
            off = len(bin_data)
            bin_data.extend(raw_bytes)
            bv_i = len(buffer_views)
            buffer_views.append({
                "buffer": 0,
                "byteOffset": off,
                "byteLength": len(raw_bytes),
                "target": target,
            })
            return bv_i

        pos_bv = add_bv(pos.tobytes(), 34962)
        norm_bv = add_bv(norm.tobytes(), 34962)
        uv_bv = add_bv(uv.tobytes(), 34962)
        idx_bv = add_bv(idx.tobytes(), 34963)

        pos_acc = len(accessors)
        accessors.append({
            "bufferView": pos_bv,
            "byteOffset": 0,
            "componentType": 5126,
            "count": len(pos),
            "type": "VEC3",
            "min": [float(x) for x in pos.min(axis=0)],
            "max": [float(x) for x in pos.max(axis=0)],
        })
        norm_acc = len(accessors)
        accessors.append({
            "bufferView": norm_bv,
            "byteOffset": 0,
            "componentType": 5126,
            "count": len(norm),
            "type": "VEC3",
        })
        uv_acc = len(accessors)
        accessors.append({
            "bufferView": uv_bv,
            "byteOffset": 0,
            "componentType": 5126,
            "count": len(uv),
            "type": "VEC2",
        })
        idx_acc = len(accessors)
        accessors.append({
            "bufferView": idx_bv,
            "byteOffset": 0,
            "componentType": 5125,
            "count": len(idx),
            "type": "SCALAR",
        })

        mesh_primitives.append({
            "attributes": {
                "POSITION": pos_acc,
                "NORMAL": norm_acc,
                "TEXCOORD_0": uv_acc,
            },
            "indices": idx_acc,
            "material": pb.material_idx,
            "mode": 4,
        })

    if screen_png_path and os.path.exists(screen_png_path):
        with open(screen_png_path, "rb") as sf:
            png_bytes = sf.read()
    else:
        png_bytes = build_core_texture(archetype, brand_name, acc1, acc2, acc3)

    while len(bin_data) % 4 != 0:
        bin_data.append(0)
    img_off = len(bin_data)
    bin_data.extend(png_bytes)
    img_bv = len(buffer_views)
    buffer_views.append({
        "buffer": 0,
        "byteOffset": img_off,
        "byteLength": len(png_bytes),
    })

    while len(bin_data) % 4 != 0:
        bin_data.append(0)

    materials = [
        {
            "name": "mat_primary",
            "doubleSided": True,
            "pbrMetallicRoughness": {
                "baseColorFactor": [*mat0_col, 1.0],
                "metallicFactor": mat0_metal,
                "roughnessFactor": mat0_rough,
            },
        },
        {
            "name": "mat_dark_metal",
            "doubleSided": True,
            "pbrMetallicRoughness": {
                "baseColorFactor": [*mat1_col, 1.0],
                "metallicFactor": 0.82,
                "roughnessFactor": 0.26,
            },
        },
        {
            "name": "mat_secondary",
            "doubleSided": True,
            "pbrMetallicRoughness": {
                "baseColorFactor": [*mat2_col, 1.0],
                "metallicFactor": mat2_metal,
                "roughnessFactor": mat2_rough,
            },
        },
        {
            "name": "mat_textured_core",
            "doubleSided": True,
            "pbrMetallicRoughness": {
                "baseColorFactor": [1.0, 1.0, 1.0, 1.0],
                "baseColorTexture": {"index": 0},
                "metallicFactor": 0.05,
                "roughnessFactor": 0.14,
            },
            "emissiveTexture": {"index": 0},
            "emissiveFactor": [0.85, 0.85, 0.85],
        },
        {
            "name": "mat_emissive_glow",
            "doubleSided": True,
            "pbrMetallicRoughness": {
                "baseColorFactor": [*acc1_lin, 1.0],
                "metallicFactor": 0.1,
                "roughnessFactor": 0.1,
            },
            "emissiveFactor": [min(1.0, c * 1.35 + 0.15) for c in acc1_lin],
        },
    ]

    gltf = {
        "asset": {
            "version": "2.0",
            "generator": "morphdeck-3d-glb-engine-1.0",
            "copyright": "Copyright 2026 Mohamed S. Sissoko (7AK3SH1) - Apache-2.0",
        },
        "scene": 0,
        "scenes": [{"nodes": [0]}],
        "nodes": [{"mesh": 0, "name": f"hero_{archetype}"}],
        "meshes": [{"name": f"mesh_{archetype}", "primitives": mesh_primitives}],
        "materials": materials,
        "textures": [{"sampler": 0, "source": 0}],
        "images": [{"bufferView": img_bv, "mimeType": "image/png", "name": "core_tex"}],
        "samplers": [{"magFilter": 9729, "minFilter": 9729, "wrapS": 33071, "wrapT": 33071}],
        "accessors": accessors,
        "bufferViews": buffer_views,
        "buffers": [{"byteLength": len(bin_data)}],
    }

    json_bytes = bytearray(json.dumps(gltf, separators=(",", ":")).encode("utf-8"))
    while len(json_bytes) % 4 != 0:
        json_bytes.append(0x20)

    total_len = 12 + 8 + len(json_bytes) + 8 + len(bin_data)
    os.makedirs(os.path.dirname(os.path.abspath(out_glb)), exist_ok=True)
    with open(out_glb, "wb") as out:
        out.write(struct.pack("<4sII", b"glTF", 2, total_len))
        out.write(struct.pack("<I4s", len(json_bytes), b"JSON"))
        out.write(json_bytes)
        out.write(struct.pack("<I4s", len(bin_data), b"BIN\x00"))
        out.write(bin_data)

    return out_glb


def main():
    choices = ARCHETYPES + sorted(ARCHETYPE_ALIASES.keys())
    parser = argparse.ArgumentParser(description="Generate a procedural PBR glTF 2.0 (.glb) 3D model for morphdeck-3d.")
    parser.add_argument("--archetype", choices=choices, required=True, help="3D subject archetype")
    parser.add_argument("--output", required=True, help="Output .glb path")
    parser.add_argument("--brand", default="MORPHDECK", help="Brand name baked onto self-lit core texture")
    parser.add_argument("--accent1", default="#2563EB", help="Primary hex accent")
    parser.add_argument("--accent2", default="#10B981", help="Secondary hex accent")
    parser.add_argument("--accent3", default="#F59E0B", help="Tertiary hex accent")
    parser.add_argument("--screen-image", default="", help="Optional custom PNG image for screen/core texture")
    args = parser.parse_args()

    generate_subject_glb(
        archetype=args.archetype,
        out_glb=args.output,
        brand_name=args.brand,
        accent1_hex=args.accent1,
        accent2_hex=args.accent2,
        accent3_hex=args.accent3,
        screen_png_path=args.screen_image,
    )
    print(f"Generated 3D GLB ({canonical_archetype(args.archetype)}) -> {args.output} ({os.path.getsize(args.output)} bytes)")


if __name__ == "__main__":
    main()
