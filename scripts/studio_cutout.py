#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Museum-Grade Studio Cutout, True Architectural Portal Pavilion & Mullion Synthesizer,
Cloud-Proof 1D Seam Sky Splitter, Backdrop-Refracted Glass Dock, Mitered Museum Frame,
Gallery Spotlight Wall & Archival Broadsheet Engine for morphdeck-3d v4.0.0.
"""

import argparse
import os
import sys
import textwrap
import numpy as np
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont


def hex_to_rgb(hex_str: str, default=(185, 28, 28)) -> tuple[int, int, int]:
    h = (hex_str or "").strip().lstrip("#")
    if len(h) != 6:
        return default
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def load_font(style: str, size: int, weight: int = 400):
    candidates = []
    if sys.platform.startswith("win"):
        fdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        win_map = {
            "georgia_bold": ["georgiab.ttf", "georgia.ttf", "timesbd.ttf"],
            "georgia": ["georgia.ttf", "times.ttf"],
            "bahnschrift": ["bahnschrift.ttf", "arialbd.ttf"],
            "impact": ["impact.ttf", "bahnschrift.ttf"],
            "consolas_bold": ["consolab.ttf", "consola.ttf"],
            "segoe": ["segoeui.ttf", "arial.ttf"],
            "segoe_semi": ["seguisb.ttf", "segoeuib.ttf", "arialbd.ttf"],
        }
        candidates.extend(os.path.join(fdir, f) for f in win_map.get(style.lower(), ["segoeui.ttf", "arial.ttf"]))
    elif sys.platform == "darwin":
        if "georgia" in style.lower():
            candidates.append(
                "/System/Library/Fonts/Supplemental/Georgia Bold.ttf"
                if "bold" in style.lower()
                else "/System/Library/Fonts/Supplemental/Georgia.ttf"
            )
        candidates.append(
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
            if weight >= 600
            else "/System/Library/Fonts/Supplemental/Arial.ttf"
        )
    else:
        if "georgia" in style.lower():
            candidates.append(
                "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
                if "bold" in style.lower()
                else "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
            )
        candidates.append(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
            if weight >= 600
            else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
        )

    for fp in candidates:
        if os.path.exists(fp):
            try:
                f = ImageFont.truetype(fp, size)
                if "bahnschrift" in fp.lower():
                    try:
                        f.set_variation_by_axes([weight, 100])
                    except Exception:
                        pass
                return f
            except Exception:
                continue
    return ImageFont.load_default()


def cover_crop(img: Image.Image, target_w: int, target_h: int, bias_y: float = 0.5) -> Image.Image:
    src_r = img.width / max(1, img.height)
    dst_r = target_w / max(1, target_h)
    if src_r > dst_r:
        nw = max(1, int(img.height * dst_r))
        ox = max(0, (img.width - nw) // 2)
        img = img.crop((ox, 0, ox + nw, img.height))
    else:
        nh = max(1, int(img.width / dst_r))
        oy = max(0, min(img.height - nh, int((img.height - nh) * bias_y)))
        img = img.crop((0, oy, img.width, oy + nh))
    return img.resize((target_w, target_h), Image.Resampling.LANCZOS)


def apply_tonal_text_scrim(
    img_or_path,
    out_path: str | None = None,
    zone: str = "left",
    strength: float = 0.66,
) -> Image.Image:
    """
    Apply an optical graduated neutral-density exposure ramp directly into a full-bleed photo
    so frameless white typography achieves crisp contrast without clunky rounded scrim boxes.
    """
    im = Image.open(img_or_path).convert("RGB") if isinstance(img_or_path, (str, os.PathLike)) else img_or_path.convert("RGB")
    arr = np.asarray(im, dtype=np.float32)
    h, w, _ = arr.shape
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    xn = xx / max(1.0, float(w - 1))
    yn = yy / max(1.0, float(h - 1))

    if zone == "left":
        mask = np.clip(1.0 - (xn - 0.08) / 0.54, 0.0, 1.0) ** 1.35
    elif zone == "right":
        mask = np.clip((xn - 0.38) / 0.54, 0.0, 1.0) ** 1.35
    elif zone == "bottom":
        mask = np.clip((yn - 0.42) / 0.58, 0.0, 1.0) ** 1.30
    elif zone == "top_bottom":
        m_top = np.clip(1.0 - yn / 0.34, 0.0, 1.0) ** 1.5 * 0.65
        m_bot = np.clip((yn - 0.58) / 0.42, 0.0, 1.0) ** 1.4
        mask = np.maximum(m_top, m_bot)
    else:
        mask = np.clip(1.0 - (xn - 0.06) / 0.56, 0.0, 1.0) ** 1.35

    darken = 1.0 - (mask[..., None] * float(strength))
    out_arr = np.clip(arr * darken, 0, 255).astype(np.uint8)
    res = Image.fromarray(out_arr, mode="RGB")
    if out_path:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        res.save(out_path, "PNG")
    return res


def _flood_outer_bg(cand_u8: np.ndarray, seed_top_row: bool = False) -> np.ndarray:
    h, w = cand_u8.shape
    padded = np.zeros((h + 2, w + 2), dtype=np.uint8) if not seed_top_row else np.full((h + 2, w + 2), 255, dtype=np.uint8)
    padded[1 : h + 1, 1 : w + 1] = cand_u8
    if seed_top_row:
        padded[0, :] = 0
    mask_rgb = Image.fromarray(np.dstack([padded, padded, padded]), mode="RGB")
    ImageDraw.floodfill(mask_rgb, (0, 0), (128, 0, 0))
    res = np.asarray(mask_rgb)[1 : h + 1, 1 : w + 1, 0]
    return res == 128


def make_studio_cutout(
    src_path: str,
    out_path: str,
    tolerance: float = 28.0,
    feather: float = 1.2,
) -> str:
    """Segment an isolated object on a clean studio backdrop into a crisp transparent RGBA cutout."""
    im = Image.open(src_path).convert("RGBA")
    arr = np.asarray(im, dtype=np.float32)
    h, w, _ = arr.shape

    # Try rembg neural matting if installed in the buyer's Python environment
    try:
        from rembg import remove  # type: ignore

        cut = remove(im)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        cut.save(out_path, "PNG")
        return out_path
    except Exception:
        pass

    p = max(4, min(16, w // 20, h // 20))
    corners = np.vstack([
        arr[:p, :p, :3].reshape(-1, 3),
        arr[:p, -p:, :3].reshape(-1, 3),
        arr[-p:, :p, :3].reshape(-1, 3),
        arr[-p:, -p:, :3].reshape(-1, 3),
    ])
    bg_rgb = np.median(corners, axis=0)
    dist = np.linalg.norm(arr[..., :3] - bg_rgb[None, None, :], axis=-1)

    cand = np.where(dist <= tolerance, 0, 255).astype(np.uint8)
    cand_sealed = np.asarray(Image.fromarray(cand, mode="L").filter(ImageFilter.MaxFilter(3)))
    exterior_bg = _flood_outer_bg(cand_sealed, seed_top_row=False) & (cand == 0)

    alpha = np.where(exterior_bg, 0.0, 255.0).astype(np.float32)
    edge_zone = exterior_bg & (dist > tolerance * 0.72) & (dist <= tolerance * 1.45)
    ramp = np.clip((dist - tolerance * 0.72) / max(1e-3, tolerance * 0.73), 0.0, 1.0) * 180.0
    alpha = np.where(edge_zone, ramp, alpha)

    alpha_im = Image.fromarray(np.clip(alpha, 0, 255).astype(np.uint8), mode="L")
    if feather > 0.1:
        alpha_im = alpha_im.filter(ImageFilter.GaussianBlur(feather))

    rgba = np.dstack([arr[..., :3].astype(np.uint8), np.asarray(alpha_im)])
    out_im = Image.fromarray(rgba, mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    out_im.save(out_path, "PNG")
    return out_path


def render_architectural_portal_exterior(
    bg_photo_path: str,
    out_path: str,
    is_light: bool = True,
    stone_hex: str = "#F2EDE4",
    box_frac: tuple[float, float, float, float] = (0.31, 0.25, 0.69, 0.65),
) -> str:
    """
    Synthesize a photorealistic cantilevered architectural pavilion & infinity reflection pool
    whose sky and central panoramic glass opening are 100% alpha-transparent (modeled on
    high-craft architectural fly-throughs, never sticker cutouts over tourist photos).
    """
    W, H = 1920, 1080
    rng = np.random.default_rng(404)
    bg_im = cover_crop(Image.open(bg_photo_path).convert("RGB"), W, H) if (bg_photo_path and os.path.exists(bg_photo_path)) else Image.new("RGB", (W, H), (32, 68, 98))
    bg_arr = np.asarray(bg_im, dtype=np.float32)

    stone_rgb = np.array(hex_to_rgb(stone_hex, (242, 237, 228) if is_light else (28, 32, 40)), dtype=np.float32)
    if not is_light and np.mean(stone_rgb) > 160:
        stone_rgb = np.array([30.0, 33.0, 40.0], dtype=np.float32)

    canvas = np.zeros((H, W, 4), dtype=np.float32)
    yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)

    # Pavilion geometry coordinates
    pav_x0, pav_x1 = int(W * 0.17), int(W * 0.83)
    pav_y0, pav_y1 = int(H * 0.11), int(H * 0.66)
    win_x0, win_y0 = int(W * box_frac[0]), int(H * box_frac[1])
    win_x1, win_y1 = int(W * box_frac[2]), int(H * box_frac[3])
    soffit_x0, soffit_x1 = win_x0 - 54, win_x1 + 54
    soffit_y0 = win_y0 - 52
    pool_y = int(H * 0.675)

    # Fine travertine/architectural concrete micro-texture
    grain = rng.normal(0.0, 2.6, size=(H, W, 1)).astype(np.float32)
    sun_grad = (1.04 - 0.12 * (yy / H))[..., None] - 0.05 * ((xx / W) - 0.35)[..., None]
    wall_col = np.clip(stone_rgb[None, None, :] * sun_grad + grain, 0, 255)

    # Main pavilion volume
    pav_mask = (xx >= pav_x0) & (xx <= pav_x1) & (yy >= pav_y0) & (yy <= pool_y)
    # Right wing lower roof step for realistic asymmetric architectural massing
    wing_x0, wing_x1 = pav_x1, int(W * 0.91)
    wing_y0 = int(H * 0.27)
    wing_mask = (xx > wing_x0) & (xx <= wing_x1) & (yy >= wing_y0) & (yy <= pool_y)
    # Left natural coastal rock / plinth silhouette
    left_y_curve = H * 0.44 + 42.0 * np.sin(xx / 95.0) + 0.18 * np.maximum(0.0, pav_x0 - xx)
    rock_mask = (xx < pav_x0) & (yy >= left_y_curve) & (yy <= pool_y)

    canvas[pav_mask, :3] = wall_col[pav_mask]
    canvas[pav_mask, 3] = 255.0
    canvas[wing_mask, :3] = wall_col[wing_mask] * 0.90
    canvas[wing_mask, 3] = 255.0
    rock_col = np.clip(np.array([62.0, 58.0, 54.0])[None, None, :] * (0.75 + 0.35 * (1.0 - yy / pool_y)[..., None]) + grain * 2.4, 0, 255)
    canvas[rock_mask, :3] = rock_col[rock_mask]
    canvas[rock_mask, 3] = 255.0

    img = Image.fromarray(canvas.astype(np.uint8), mode="RGBA")
    d = ImageDraw.Draw(img)

    # Recessed perspective soffit ceiling & side reveals leading into the window
    soffit_col = tuple(np.clip(stone_rgb * 0.76, 0, 255).astype(int)) + (255,)
    reveal_l_col = tuple(np.clip(stone_rgb * 0.84, 0, 255).astype(int)) + (255,)
    reveal_r_col = tuple(np.clip(stone_rgb * 0.68, 0, 255).astype(int)) + (255,)
    d.polygon([(soffit_x0, soffit_y0), (soffit_x1, soffit_y0), (win_x1, win_y0), (win_x0, win_y0)], fill=soffit_col)
    d.polygon([(soffit_x0, soffit_y0), (win_x0, win_y0), (win_x0, win_y1), (soffit_x0, pool_y)], fill=reveal_l_col)
    d.polygon([(soffit_x1, soffit_y0), (win_x1, win_y0), (win_x1, win_y1), (soffit_x1, pool_y)], fill=reveal_r_col)

    # Subtle recessed architectural downlights in soffit
    for lx_f in (0.37, 0.50, 0.63):
        lx = int(W * lx_f)
        ly = (soffit_y0 + win_y0) // 2
        d.rectangle([lx - 10, ly - 3, lx + 10, ly + 3], fill=(38, 36, 34, 255))
        d.rectangle([lx - 6, ly - 1, lx + 6, ly + 1], fill=(255, 244, 214, 255))

    # Left vertical architectural accent window
    lw_x0, lw_x1 = int(W * 0.215), int(W * 0.248)
    lw_y0, lw_y1 = int(H * 0.30), int(H * 0.49)
    d.rectangle([lw_x0 - 4, lw_y0 - 4, lw_x1 + 4, lw_y1 + 4], fill=(24, 26, 30, 255))
    d.rectangle([lw_x0, lw_y0, lw_x1, lw_y1], fill=(38, 108, 128, 215))

    # Terrace steps & coping ledge above infinity pool
    terrace_col = tuple(np.clip(stone_rgb * 0.94, 0, 255).astype(int)) + (255,)
    shadow_ledge = tuple(np.clip(stone_rgb * 0.58, 0, 255).astype(int)) + (255,)
    d.rectangle([0, win_y1, W, pool_y], fill=terrace_col)
    d.line([(0, pool_y - 2), (W, pool_y - 2)], fill=shadow_ledge, width=4)

    # Anodized architectural window frame & 100% transparent portal aperture
    d.rectangle([win_x0 - 6, win_y0 - 6, win_x1 + 6, win_y1 + 2], fill=(22, 24, 28, 255))
    out_arr = np.array(img, dtype=np.float32)
    out_arr[win_y0:win_y1, win_x0:win_x1, :] = 0.0

    # Infinity reflection pool (y = pool_y .. H): reflects pavilion + real background landscape with water ripples
    pool_h = H - pool_y
    # Composite what stands above pool_y over bg_arr so reflection matches exact scene colors
    above_alpha = out_arr[:pool_y, :, 3:4] / 255.0
    above_comp = out_arr[:pool_y, :, :3] * above_alpha + bg_arr[:pool_y, :, :3] * (1.0 - above_alpha)
    flipped = above_comp[::-1, :, :]
    refl_src = np.array(
        Image.fromarray(flipped.astype(np.uint8)).resize((W, pool_h), Image.Resampling.BILINEAR).filter(ImageFilter.GaussianBlur(2.2)),
        dtype=np.float32,
    )

    py_rel = np.linspace(0.0, 1.0, pool_h, dtype=np.float32)[:, None]
    ripple_shift = (np.sin(np.linspace(0, 46.0, pool_h)[:, None] + xx[:pool_h, :] * 0.015) * (2.0 + 7.0 * py_rel)).astype(np.int32)
    rx_idx = np.clip(xx[:pool_h, :].astype(np.int32) + ripple_shift, 0, W - 1)
    ry_idx = np.arange(pool_h, dtype=np.int32)[:, None]
    rippled = refl_src[ry_idx, rx_idx]

    water_tint = np.array([12.0, 92.0, 108.0], dtype=np.float32) if is_light else np.array([8.0, 26.0, 38.0], dtype=np.float32)
    fresnel_mix = np.clip(0.32 + 0.36 * py_rel, 0.0, 0.78)[..., None]
    pool_rgb = rippled * (1.0 - fresnel_mix) + water_tint[None, None, :] * fresnel_mix
    out_arr[pool_y:, :, :3] = np.clip(pool_rgb, 0, 255)
    out_arr[pool_y:, :, 3] = 255.0

    final_im = Image.fromarray(out_arr.astype(np.uint8), mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    final_im.save(out_path, "PNG")
    return out_path


def render_architectural_portal_interior(
    out_path: str,
    is_light: bool = True,
    wall_hex: str = "#EFECE6",
    box_frac: tuple[float, float, float, float] = (0.12, 0.24, 0.88, 0.78),
) -> str:
    """
    Synthesize a minimalist architectural gallery/villa interior with floor-to-ceiling steel mullions
    and 100% transparent panoramic window glass so zooming through reveals the 4K horizon seamlessly.
    """
    W, H = 1920, 1080
    base_rgb = np.array(hex_to_rgb(wall_hex, (239, 236, 230) if is_light else (22, 25, 32)), dtype=np.float32)
    if not is_light and np.mean(base_rgb) > 150:
        base_rgb = np.array([22.0, 25.0, 32.0], dtype=np.float32)

    yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
    wx0, wy0 = int(W * box_frac[0]), int(H * box_frac[1])
    wx1, wy1 = int(W * box_frac[2]), int(H * box_frac[3])

    arr = np.zeros((H, W, 4), dtype=np.float32)
    # Perspective ceiling (y = 0 .. wy0)
    ceil_shade = 0.82 + 0.18 * (yy / max(1.0, float(wy0)))[..., None]
    arr[:wy0, :, :3] = np.clip(base_rgb[None, None, :] * ceil_shade[:wy0], 0, 255)
    arr[:wy0, :, 3] = 255.0

    # Polished architectural floor (y = wy1 .. H) with soft window bounce light
    floor_t = ((yy - wy1) / max(1.0, float(H - wy1)))[..., None]
    center_glow = np.exp(-(((xx / W) - 0.5) ** 2) / 0.14)[..., None]
    floor_col = base_rgb[None, None, :] * (0.76 + 0.20 * (1.0 - floor_t) + 0.10 * center_glow)
    arr[wy1:, :, :3] = np.clip(floor_col[wy1:], 0, 255)
    arr[wy1:, :, 3] = 255.0

    # Left & right structural columns/walls framing the panoramic glass
    arr[wy0:wy1, :wx0, :3] = np.clip(base_rgb * 0.78, 0, 255)
    arr[wy0:wy1, :wx0, 3] = 255.0
    arr[wy0:wy1, wx1:, :3] = np.clip(base_rgb * 0.74, 0, 255)
    arr[wy0:wy1, wx1:, 3] = 255.0

    im = Image.fromarray(arr.astype(np.uint8), mode="RGBA")
    d = ImageDraw.Draw(im)

    # Recessed architectural slot cove & linear ceiling luminaires
    cove_col = tuple(np.clip(base_rgb * 0.45, 0, 255).astype(int)) + (255,)
    d.rectangle([int(W * 0.08), int(wy0 * 0.22), int(W * 0.92), int(wy0 * 0.28)], fill=cove_col)
    for lx in (int(W * 0.30), int(W * 0.64)):
        d.rectangle([lx - 28, int(wy0 * 0.66), lx + 28, int(wy0 * 0.70)], fill=(28, 28, 32, 255))

    # Floor-to-ceiling slim black steel mullions (keep center wide open!)
    mullion_col = (20, 22, 26, 255)
    d.rectangle([wx0 - 10, wy0 - 8, wx1 + 10, wy0 + 4], fill=mullion_col)
    d.rectangle([wx0 - 10, wy1 - 4, wx1 + 10, wy1 + 8], fill=mullion_col)
    d.rectangle([wx0 - 10, wy0, wx0 + 6, wy1], fill=mullion_col)
    d.rectangle([wx1 - 6, wy0, wx1 + 10, wy1], fill=mullion_col)
    for mx_frac in (0.34, 0.72):
        mx = int(W * mx_frac)
        d.rectangle([mx - 5, wy0, mx + 5, wy1], fill=mullion_col)

    # Minimalist low-profile architectural lounge bench along the lower sill (adds scale & realism)
    sofa_col = tuple(np.clip(base_rgb * 0.88, 0, 255).astype(int)) + (255,)
    sofa_sh = tuple(np.clip(base_rgb * 0.62, 0, 255).astype(int)) + (255,)
    d.rounded_rectangle([int(W * 0.36), wy1 + 26, int(W * 0.84), H - 14], radius=10, fill=sofa_col, outline=sofa_sh, width=2)
    d.line([(int(W * 0.52), wy1 + 28), (int(W * 0.52), H - 16)], fill=sofa_sh, width=2)
    d.line([(int(W * 0.68), wy1 + 28), (int(W * 0.68), H - 16)], fill=sofa_sh, width=2)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    im.save(out_path, "PNG")
    return out_path


def punch_portal_mask(
    photo_path: str,
    out_path: str,
    box_frac: tuple[float, float, float, float] = (0.26, 0.17, 0.74, 0.83),
    radius: int = 8,
    bevel: bool = True,
) -> str:
    """Punch a clean dark-mullioned architectural window aperture out of an image (zero white sticker border)."""
    im = cover_crop(Image.open(photo_path).convert("RGB"), 1920, 1080).convert("RGBA")
    w, h = im.size
    x0, y0 = int(w * box_frac[0]), int(h * box_frac[1])
    x1, y1 = int(w * box_frac[2]), int(h * box_frac[3])

    mask = Image.new("L", (w * 2, h * 2), 255)
    ImageDraw.Draw(mask).rounded_rectangle([x0 * 2, y0 * 2, x1 * 2, y1 * 2], radius=radius * 2, fill=0)
    mask = mask.resize((w, h), Image.Resampling.LANCZOS)

    arr = np.array(im)
    arr[..., 3] = np.array(mask)
    out_im = Image.fromarray(arr, mode="RGBA")

    if bevel:
        d = ImageDraw.Draw(out_im)
        d.rounded_rectangle([x0 - 5, y0 - 5, x1 + 5, y1 + 5], radius=radius + 3, outline=(18, 20, 24, 245), width=8)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    out_im.save(out_path, "PNG")
    return out_path


def split_horizon_cutout(
    photo_path: str,
    out_path: str,
    sky_tolerance: float = 40.0,
    ridge_frac: float | None = None,
    size: tuple[int, int] | None = None,
) -> str:
    """
    Cloud-proof 1D horizon seam tracker: finds a single-valued ridge curve y_ridge(x)
    separating upper sky/clouds (100% alpha=0 so giant sky display letters never get blotched)
    from opaque foreground mountain/terrain silhouette below.
    """
    im = Image.open(photo_path).convert("RGBA")
    if size:
        im = cover_crop(im.convert("RGB"), size[0], size[1]).convert("RGBA")
    w, h = im.size
    arr = np.asarray(im, dtype=np.float32)

    if ridge_frac is not None:
        xn = np.linspace(0.0, 1.0, w, dtype=np.float32)
        ridge_px = (
            float(ridge_frac) + 0.018 * np.sin(xn * 9.0 + 0.6) - 0.025 * np.exp(-((xn - 0.5) ** 2) / 0.03)
        ) * float(h)
    else:
        # Analyze vertical edge contrast + sky chromaticity in [0.16*H .. 0.44*H]
        y_min = int(h * 0.16)
        y_max = int(h * 0.44)
        rgb = arr[y_min:y_max, :, :3]
        lum = 0.299 * rgb[..., 0] + 0.587 * rgb[..., 1] + 0.114 * rgb[..., 2]
        # Vertical gradient + local texture variance (clouds are smooth, rock/ridge has sharp gradient)
        gy = np.abs(np.diff(lum, axis=0, prepend=lum[:1, :]))
        gx = np.abs(np.diff(lum, axis=1, prepend=lum[:, :1]))
        edge_score = gy * 1.4 + gx * 0.7

        # Smooth horizontally to track coherent mountain/horizon ridge
        kernel_w = 9
        pad_e = np.pad(edge_score, ((0, 0), (kernel_w // 2, kernel_w // 2)), mode="edge")
        smooth_e = sum(pad_e[:, k : k + w] for k in range(kernel_w)) / float(kernel_w)

        # Dynamic-programming left-to-right seam with max step +-3 px per column
        sh_h = smooth_e.shape[0]
        dp = smooth_e.copy()
        back = np.zeros((sh_h, w), dtype=np.int32)
        for x in range(1, w):
            prev = dp[:, x - 1]
            # Vectorized max over shifts [-2, -1, 0, 1, 2]
            best_v = prev.copy()
            best_s = np.zeros(sh_h, dtype=np.int32)
            for s in (-2, -1, 1, 2):
                shifted = np.empty_like(prev)
                if s < 0:
                    shifted[:s] = prev[-s:] - 1.5 * abs(s)
                    shifted[s:] = -1e9
                else:
                    shifted[s:] = prev[:-s] - 1.5 * abs(s)
                    shifted[:s] = -1e9
                better = shifted > best_v
                best_v = np.where(better, shifted, best_v)
                best_s = np.where(better, s, best_s)
            dp[:, x] += best_v
            back[:, x] = np.arange(sh_h, dtype=np.int32) - best_s

        seam = np.zeros(w, dtype=np.float32)
        cur = int(np.argmax(dp[:, w - 1]))
        seam[w - 1] = cur
        for x in range(w - 1, 0, -1):
            cur = int(np.clip(back[cur, x], 0, sh_h - 1))
            seam[x - 1] = cur
        # Light 1D smoothing for sub-pixel anti-aliased horizon
        k15 = np.ones(15, dtype=np.float32) / 15.0
        ridge_px = np.convolve(np.pad(seam + y_min, (7, 7), mode="edge"), k15, mode="valid")

    yy = np.arange(h, dtype=np.float32)[:, None]
    alpha = np.clip((yy - ridge_px[None, :]) / 1.8, 0.0, 1.0) * 255.0
    out_rgba = np.dstack([arr[..., :3].astype(np.uint8), alpha.astype(np.uint8)])
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    Image.fromarray(out_rgba, "RGBA").save(out_path, "PNG")
    return out_path


def render_refracted_glass_dock(
    bg_path: str,
    out_path: str,
    crop_box: tuple[int, int, int, int] = (80, 772, 1840, 1028),
    blur_radius: int = 28,
    tint_rgba: tuple[int, int, int, int] = (14, 18, 28, 116),
    radius: int = 32,
) -> str:
    """Create a sleek borderless backdrop-refracted glass dock cropped and blurred directly from bg_path."""
    bg = cover_crop(Image.open(bg_path).convert("RGB"), 1920, 1080)
    w = crop_box[2] - crop_box[0]
    h = crop_box[3] - crop_box[1]
    sliced = bg.crop(crop_box).filter(ImageFilter.GaussianBlur(blur_radius))
    glass = Image.alpha_composite(sliced.convert("RGBA"), Image.new("RGBA", (w, h), tint_rgba))

    iw, ih = w - 12, h - 16
    mask = Image.new("L", (iw * 2, ih * 2), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, iw * 2 - 1, ih * 2 - 1], radius=radius * 2, fill=246)
    mask = mask.resize((iw, ih), Image.Resampling.LANCZOS)

    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([12, 14, w - 12, h - 4], radius=radius + 2, fill=(0, 0, 0, 128))
    canvas = Image.alpha_composite(canvas, sh.filter(ImageFilter.GaussianBlur(10)))
    canvas.paste(glass.crop((6, 6, w - 6, h - 10)), (6, 6), mask)

    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([6, 6, w - 6, h - 10], radius=radius, outline=(255, 255, 255, 58), width=1)
    d.line([(6 + radius, 7), (w - 6 - radius, 7)], fill=(255, 255, 255, 145), width=2)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    canvas.save(out_path, "PNG")
    return out_path


def render_organic_ink_drips(
    out_path: str,
    width: int = 1920,
    height: int = 240,
    color_hex: str = "#B91C1C",
    start_x_frac: float = 0.0,
) -> str:
    """
    Generate an authentic archival deckle-edge torn header banner with fine letterpress fiber grain
    (replaces cartoonish blobs with museum documentary craft).
    """
    rgb = hex_to_rgb(color_hex, (185, 28, 28))
    rng = np.random.default_rng(74)
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
    x_norm = np.linspace(0.0, 1.0, width, dtype=np.float32)

    # Multi-scale Fractal Brownian deckle tear profile along bottom of archival header band
    tear = (
        92.0
        + 11.0 * np.sin(x_norm * 13.0 + 0.4)
        + 6.5 * np.cos(x_norm * 37.0 + 1.1)
        + 3.5 * np.sin(x_norm * 97.0)
        + rng.normal(0.0, 1.8, size=width).astype(np.float32)
    )
    # White fibrous deckle under-edge (4px offset) + crisp crimson over-layer
    alpha_crimson = np.clip((tear[None, :] - yy) * 0.85, 0.0, 1.0)
    alpha_fiber = np.clip((tear[None, :] + 5.5 - yy) * 0.65, 0.0, 1.0) * (1.0 - alpha_crimson)

    x0 = int(width * start_x_frac)
    if x0 > 0:
        fade_x = np.clip((xx - x0) / 80.0, 0.0, 1.0)
        alpha_crimson *= fade_x
        alpha_fiber *= fade_x

    grain = rng.integers(-8, 9, size=(height, width, 1), dtype=np.int16)
    crim_rgb = np.clip(np.full((height, width, 3), rgb, dtype=np.int16) + grain, 0, 255).astype(np.float32)
    fiber_rgb = np.full((height, width, 3), (244.0, 239.0, 228.0), dtype=np.float32)

    tot_a = np.clip(alpha_crimson + alpha_fiber * 0.92, 0.0, 1.0)
    comp_rgb = np.where(
        tot_a[..., None] > 1e-4,
        (crim_rgb * alpha_crimson[..., None] + fiber_rgb * (alpha_fiber * 0.92)[..., None]) / np.maximum(tot_a[..., None], 1e-4),
        0.0,
    )
    rgba = np.dstack([np.clip(comp_rgb, 0, 255).astype(np.uint8), (tot_a * 250.0).astype(np.uint8)])
    canvas = Image.fromarray(rgba, mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    canvas.save(out_path, "PNG")
    return out_path


def render_soft_paper_cloud(
    out_path: str,
    width: int = 1920,
    height: int = 620,
    paper_hex: str = "#F5EFE2",
) -> str:
    """Create a clean archival torn-edge vellum reading sheet with subtle drop shadow (never a blurry smudge)."""
    rgb = hex_to_rgb(paper_hex, (245, 239, 226))
    rng = np.random.default_rng(108)
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)

    # Deckle-edged rectangular sheet inside [36..width-36, 28..height-36]
    jx = 4.0 * np.sin(yy[:, 0] * 0.09) + rng.normal(0, 1.2, size=height)
    jy = 4.0 * np.cos(xx[0, :] * 0.08) + rng.normal(0, 1.2, size=width)
    in_x = (xx >= (38.0 + jx[:, None])) & (xx <= (width - 38.0 + jx[:, None]))
    in_y = (yy >= (28.0 + jy[None, :])) & (yy <= (height - 38.0 + jy[None, :]))
    sheet = (in_x & in_y).astype(np.uint8) * 248

    sh_im = Image.fromarray(sheet, mode="L").filter(ImageFilter.GaussianBlur(12))
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    sh_rgba = np.zeros((height, width, 4), dtype=np.uint8)
    sh_rgba[..., 3] = (np.asarray(sh_im, dtype=np.float32) * 0.28).astype(np.uint8)
    canvas = Image.alpha_composite(canvas, Image.fromarray(sh_rgba, "RGBA"))

    grain = rng.integers(-4, 5, size=(height, width, 1), dtype=np.int16)
    p_rgb = np.clip(np.full((height, width, 3), rgb, dtype=np.int16) + grain, 0, 255).astype(np.uint8)
    sheet_smooth = np.asarray(Image.fromarray(sheet, mode="L").filter(ImageFilter.GaussianBlur(0.8)))
    fg = Image.fromarray(np.dstack([p_rgb, sheet_smooth]), mode="RGBA")
    canvas = Image.alpha_composite(canvas, fg)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    canvas.save(out_path, "PNG")
    return out_path


def render_cloud_sweep(
    out_path: str,
    width: int = 1920,
    height: int = 1080,
    tint_rgb: tuple[int, int, int] = (244, 247, 252),
) -> str:
    """Create a realistic multi-octave atmospheric cloud bank for airborne Morph fly-by transitions."""
    rng = np.random.default_rng(19)
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
    xn = xx / float(width)
    yn = yy / float(height)

    density = np.zeros((height, width), dtype=np.float32)
    for cx, cy, rx, ry, wgt in [
        (0.24, 0.54, 0.22, 0.26, 0.88),
        (0.44, 0.44, 0.26, 0.30, 1.00),
        (0.62, 0.50, 0.24, 0.28, 0.94),
        (0.78, 0.40, 0.20, 0.24, 0.78),
        (0.36, 0.66, 0.22, 0.20, 0.72),
        (0.58, 0.64, 0.22, 0.20, 0.70),
    ]:
        d2 = ((xn - cx) / rx) ** 2 + ((yn - cy) / ry) ** 2
        density += wgt * np.exp(-d2 * 1.85)

    # Multi-octave soft turbulence
    for gh, gw, amp in [(24, 42, 0.16), (64, 112, 0.09)]:
        n_low = Image.fromarray(((rng.random((gh, gw)) * 255).astype(np.uint8)), mode="L").resize((width, height), Image.Resampling.BICUBIC)
        density += ((np.asarray(n_low, dtype=np.float32) / 255.0) - 0.5) * amp

    alpha = np.clip(density * 0.82, 0.0, 1.0)
    alpha_im = Image.fromarray((alpha * 228.0).astype(np.uint8), mode="L").filter(ImageFilter.GaussianBlur(14))

    shade = np.clip(1.03 - (yn - 0.30) * 0.22, 0.82, 1.03)[..., None]
    rgb_arr = np.clip(np.full((height, width, 3), tint_rgb, dtype=np.float32) * shade, 0, 255).astype(np.uint8)
    out = Image.fromarray(np.dstack([rgb_arr, np.asarray(alpha_im)]), mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    out.save(out_path, "PNG")
    return out_path


def render_silhouette_horizon(
    out_path: str,
    width: int = 1920,
    height: int = 320,
    color_hex: str = "#B91C1C",
) -> str:
    """
    Create a moody 3-layer atmospheric twilight mountain ridge & horizon rim-glow
    for odometer counter slides (replaces flat single-polygon bump).
    """
    acc = np.array(hex_to_rgb(color_hex, (185, 28, 28)), dtype=np.float32)
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
    xn = xx[0, :] / float(width)

    canvas = np.zeros((height, width, 4), dtype=np.float32)
    # Soft crimson atmospheric horizon glow behind the ridges
    glow_a = np.exp(-((yy - height * 0.58) ** 2) / (2.0 * (height * 0.24) ** 2)) * 0.55
    canvas[..., :3] = acc[None, None, :] * 0.85
    canvas[..., 3] = glow_a * 255.0

    # Three layered parallax mountain ridges (back crimson-lit -> front deep obsidian)
    ridges = [
        (0.48, [18.0, 9.0, 4.0], [5.2, 13.5, 29.0], acc * 0.62, 0.88),
        (0.62, [22.0, 11.0, 5.0], [4.1, 11.2, 23.0], acc * 0.32, 0.95),
        (0.75, [15.0, 8.0, 3.5], [6.4, 15.0, 34.0], np.array([10.0, 11.0, 15.0]), 1.00),
    ]
    for base_f, amps, freqs, r_col, r_op in ridges:
        ry = np.full(width, height * base_f, dtype=np.float32)
        for a, f in zip(amps, freqs):
            ry -= a * np.sin(xn * f + f * 0.37) + (a * 0.5) * np.cos(xn * (f * 1.7))
        mask = np.clip((yy - ry[None, :]) / 1.8, 0.0, 1.0) * r_op
        prev_a = canvas[..., 3] / 255.0
        new_a = mask + prev_a * (1.0 - mask)
        canvas[..., :3] = np.where(
            new_a[..., None] > 1e-4,
            (r_col[None, None, :] * mask[..., None] + canvas[..., :3] * prev_a[..., None] * (1.0 - mask[..., None]))
            / np.maximum(new_a[..., None], 1e-4),
            0.0,
        )
        canvas[..., 3] = np.clip(new_a * 255.0, 0, 255)

    out = Image.fromarray(canvas.astype(np.uint8), mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    out.save(out_path, "PNG")
    return out_path


def render_hollow_gold_frame(
    out_path: str,
    width: int = 1920,
    height: int = 1080,
    frame_style: str = "ornate_gold",
    border_px: int = 84,
) -> str:
    """
    Synthesize a 3D mitered museum gallery picture frame with directional metallic shading,
    linen passe-partout mat, cast shadow, and 100% transparent center aperture.
    """
    mat_px = max(24, border_px // 2)
    in_pad = border_px + mat_px
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)

    # Distance from nearest outer edge + miter quadrant (top, bottom, left, right)
    d_top = yy
    d_bot = (height - 1) - yy
    d_left = xx
    d_right = (width - 1) - xx
    d_edge = np.minimum(np.minimum(d_top, d_bot), np.minimum(d_left, d_right))

    # Moulding profile wave across [0 .. border_px]
    prof = np.clip(d_edge / max(1.0, float(border_px)), 0.0, 1.0)
    flute = 0.78 + 0.22 * np.sin(prof * np.pi * 3.5) + 0.12 * np.cos(prof * np.pi * 9.0)

    # Directional top-left gallery key light across mitered rails
    is_top = (d_top <= d_left) & (d_top <= d_right) & (d_top <= d_bot)
    is_left = (d_left < d_top) & (d_left <= d_right) & (d_left <= d_bot)
    is_right = (d_right < d_top) & (d_right < d_left) & (d_right <= d_bot)
    dir_light = np.where(is_top, 1.14, np.where(is_left, 1.04, np.where(is_right, 0.78, 0.68)))

    if frame_style == "ornate_gold":
        base_mould = np.array([188.0, 148.0, 68.0], dtype=np.float32)
    else:
        base_mould = np.array([34.0, 36.0, 42.0], dtype=np.float32)

    mould_rgb = np.clip(base_mould[None, None, :] * (flute * dir_light)[..., None], 0, 255)
    # Linen museum mat board in [border_px .. in_pad]
    rng = np.random.default_rng(55)
    linen = np.clip(np.full((height, width, 3), (244, 240, 232), dtype=np.int16) + rng.integers(-3, 4, size=(height, width, 1)), 0, 255)

    is_mould = d_edge < border_px
    rgb = np.where(is_mould[..., None], mould_rgb.astype(np.uint8), linen.astype(np.uint8))
    alpha = np.full((height, width), 255, dtype=np.uint8)

    # Inner bevel accent line around aperture
    im = Image.fromarray(np.dstack([rgb, alpha]), mode="RGBA")
    d = ImageDraw.Draw(im)
    bev_col = (156, 122, 56, 255) if frame_style == "ornate_gold" else (150, 154, 162, 255)
    d.rectangle([in_pad - 4, in_pad - 4, width - in_pad + 3, height - in_pad + 3], outline=bev_col, width=4)

    arr = np.array(im)
    arr[in_pad : height - in_pad, in_pad : width - in_pad, 3] = 0
    out_im = Image.fromarray(arr, mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    out_im.save(out_path, "PNG")
    return out_im and out_path


def render_gallery_wall_room(
    out_path: str,
    width: int = 1920,
    height: int = 1080,
    wall_hex: str = "#1B1E26",
    spot_cx_frac: float = 0.50,
    spot_cy_frac: float = 0.26,
) -> str:
    """
    Synthesize a serene, tourist-free museum gallery wall with parabolic track-spotlight illumination,
    architectural dado wainscot rail, and oak gallery flooring for `picture_frame_pullback`.
    """
    rng = np.random.default_rng(301)
    base_rgb = np.array(hex_to_rgb(wall_hex, (27, 30, 38)), dtype=np.float32)
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
    xn = xx / float(width)
    yn = yy / float(height)

    # Parabolic overhead museum spotlight cone centered on the framed artwork
    dx = (xn - spot_cx_frac) / (0.22 + 0.28 * np.clip(yn, 0.0, 1.0))
    dy = (yn - spot_cy_frac) / 0.44
    spot = np.exp(-(dx * dx + dy * dy) * 1.15)
    warm_light = np.array([255.0, 242.0, 218.0], dtype=np.float32)

    plaster = base_rgb[None, None, :] * (0.68 + 0.55 * spot[..., None]) + warm_light[None, None, :] * (0.14 * spot[..., None])
    plaster += rng.normal(0.0, 1.8, size=(height, width, 1))

    dado_y = int(height * 0.77)
    floor_y = int(height * 0.88)
    # Wainscot lower paneling
    plaster[dado_y:floor_y, :, :] *= 0.82
    # Warm museum parquet floor
    floor_t = ((yy[floor_y:] - floor_y) / max(1.0, float(height - floor_y)))[..., None]
    wood = np.array([54.0, 39.0, 26.0], dtype=np.float32)[None, None, :] * (0.72 + 0.35 * spot[floor_y:, :, None]) * (1.0 - 0.28 * floor_t)
    plaster[floor_y:, :, :] = wood

    im = Image.fromarray(np.clip(plaster, 0, 255).astype(np.uint8), mode="RGB")
    d = ImageDraw.Draw(im)
    d.line([(0, dado_y), (width, dado_y)], fill=(78, 74, 68), width=3)
    d.line([(0, dado_y + 4), (width, dado_y + 4)], fill=(18, 20, 24), width=2)
    d.line([(0, floor_y), (width, floor_y)], fill=(16, 14, 12), width=3)

    # Subtle recessed wainscot molding boxes
    for bx in range(60, width - 120, 360):
        d.rectangle([bx, dado_y + 22, bx + 280, floor_y - 18], outline=(52, 52, 58), width=2)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    im.save(out_path, "PNG")
    return out_path


def generate_newspaper_sheet(
    headline: str,
    subhead: str,
    dateline: str,
    columns: list[str],
    photo_paths: list[str],
    out_path: str,
    size: tuple[int, int] = (1920, 1080),
    paper_hex: str = "#ECE4D3",
    accent_hex: str = "#B91C1C",
) -> str:
    """Synthesize a rich, non-repetitive archival broadsheet newspaper with distinct editorial dispatches."""
    w, h = size
    paper_rgb = hex_to_rgb(paper_hex, (236, 228, 211))
    acc_rgb = hex_to_rgb(accent_hex, (185, 28, 28))
    rng = np.random.default_rng(1945)

    noise = rng.integers(-7, 8, size=(h, w, 1), dtype=np.int16)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    vig = ((((xx - w * 0.5) / (w * 0.68)) ** 2 + ((yy - h * 0.5) / (h * 0.68)) ** 2) * 16).clip(0, 24).astype(np.int16)[..., None]
    base_arr = np.clip(np.full((h, w, 3), paper_rgb, dtype=np.int16) + noise - vig, 0, 255).astype(np.uint8)
    img = Image.fromarray(base_arr, "RGB")
    d = ImageDraw.Draw(img)

    ink = (22, 22, 24)
    muted = (84, 78, 70)
    rule_col = (56, 52, 48)

    f_date = load_font("georgia", 22)
    f_mast = load_font("georgia_bold", 90)
    f_sub = load_font("georgia_bold", 32)
    f_col_t = load_font("georgia_bold", 22)
    f_body = load_font("georgia", 18)

    mx = 92
    d.line([(mx, 58), (w - mx, 58)], fill=ink, width=4)
    d.line([(mx, 65), (w - mx, 65)], fill=rule_col, width=1)
    dbb = d.textbbox((0, 0), dateline.upper(), font=f_date)
    d.text((w // 2 - (dbb[2] - dbb[0]) // 2, 74), dateline.upper(), font=f_date, fill=muted)
    d.line([(mx, 108), (w - mx, 108)], fill=rule_col, width=2)

    h_lines = textwrap.wrap(headline.upper(), width=26)[:2]
    hy = 122
    for ln in h_lines:
        lbb = d.textbbox((0, 0), ln, font=f_mast)
        d.text((w // 2 - (lbb[2] - lbb[0]) // 2, hy), ln, font=f_mast, fill=ink)
        hy += 98

    sbb = d.textbbox((0, 0), subhead.upper(), font=f_sub)
    d.text((w // 2 - (sbb[2] - sbb[0]) // 2, hy + 6), subhead.upper(), font=f_sub, fill=acc_rgb)
    top_body = hy + 64
    d.line([(mx, top_body - 12), (w - mx, top_body - 12)], fill=ink, width=3)

    editorial_continuations = [
        (
            " Primary archival logbooks, instrument calibration tables, and station dispatches compiled "
            "across regional bureaus confirm that structural tolerances held steady throughout peak operational "
            "cycles. Engineers documented every milestone with synchronized chronometers, preserving a verified "
            "record of technical execution under demanding field conditions."
        ),
        (
            " Contemporary observers noted an immediate shift in long-range coordination: routes that previously "
            "required multi-stage relay handoffs consolidated into continuous corridors. Archival photographs "
            "and telemetry ledgers preserved in national collections illustrate how material precision reshaped "
            "modern industrial and cultural scale."
        ),
        (
            " Subsequent technical audits cross-checked primary witness accounts against physical hardware "
            "specimens, establishing an unbroken chain of provenance. Each recorded metric in the archive "
            "reflects measured field performance rather than theoretical projection, anchoring the historical "
            "narrative in verifiable evidence."
        ),
    ]

    n_cols = 3
    col_gap = 42
    col_w = (w - 2 * mx - col_gap * (n_cols - 1)) // n_cols
    base_cols = columns if columns else [
        "DISPATCH ARCHIVE // Coordinated logistics and precision instrumentation reshaped transcontinental operations across four consecutive seasons.",
        "FIELD TELEMETRY // Synchronized navigation fixes and calibrated optical assemblies established reliable all-weather performance benchmarks.",
        "VERIFIED RECORD // Archival inspections confirm sustained structural integrity across every documented deployment phase.",
    ]

    valid_photos = [p for p in (photo_paths or []) if p and os.path.exists(p)]
    sec_titles = ["PRIMARY DISPATCH", "TECHNICAL LEDGER", "ARCHIVAL BUREAU"]
    for ci in range(n_cols):
        cx = mx + ci * (col_w + col_gap)
        if ci > 0:
            d.line([(cx - col_gap // 2, top_body), (cx - col_gap // 2, h - 68)], fill=(142, 134, 122), width=1)

        cy = top_body + 8
        if ci == 1 and valid_photos:
            ph = ImageEnhance.Contrast(cover_crop(Image.open(valid_photos[0]).convert("L"), col_w, 336).convert("RGB")).enhance(1.18)
            tint = Image.new("RGB", ph.size, paper_rgb)
            img.paste(Image.blend(ph, tint, 0.12), (cx, cy))
            d.rectangle([cx, cy, cx + col_w, cy + 336], outline=ink, width=2)
            cy += 358

        lead = base_cols[ci % len(base_cols)].strip()
        full_article = lead + editorial_continuations[ci % len(editorial_continuations)]
        d.text((cx, cy), f"0{ci + 1}  //  {sec_titles[ci]}", font=f_col_t, fill=ink)
        cy += 34
        for line in textwrap.wrap(full_article, width=max(28, col_w // 10)):
            if cy > h - 86:
                break
            d.text((cx, cy), line, font=f_body, fill=(38, 36, 34))
            cy += 26

    d.line([(mx, h - 58), (w - mx, h - 58)], fill=ink, width=3)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    img.save(out_path, "PNG")
    return out_path


def make_contact_shadow(out_path: str, width: int = 640, height: int = 140, opacity: int = 155) -> str:
    sh = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    ImageDraw.Draw(sh).ellipse(
        [int(width * 0.12), int(height * 0.28), int(width * 0.88), int(height * 0.72)],
        fill=(0, 0, 0, opacity),
    )
    sh = sh.filter(ImageFilter.GaussianBlur(radius=int(height * 0.20)))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    sh.save(out_path, "PNG")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Museum-grade studio cutouts, architectural portals, horizons, and archival textures.")
    sub = parser.add_subparsers(dest="op", required=True)

    p1 = sub.add_parser("cutout", help="Segment an isolated studio image to transparent RGBA PNG")
    p1.add_argument("--src", required=True)
    p1.add_argument("--out", required=True)
    p1.add_argument("--tolerance", type=float, default=28.0)
    p1.add_argument("--feather", type=float, default=1.2)

    p2 = sub.add_parser("portal", help="Synthesize or punch an architectural portal aperture")
    p2.add_argument("--src", required=True)
    p2.add_argument("--out", required=True)
    p2.add_argument("--box", default="0.29,0.22,0.71,0.66", help="x0,y0,x1,y1 fractions")

    p3 = sub.add_parser("split-sky", help="Cloud-proof 1D horizon seam sky splitter")
    p3.add_argument("--src", required=True)
    p3.add_argument("--out", required=True)
    p3.add_argument("--tolerance", type=float, default=40.0)

    p4 = sub.add_parser("refract", help="Create Gaussian backdrop-refracted glass dock from background photo")
    p4.add_argument("--src", required=True)
    p4.add_argument("--out", required=True)

    args = parser.parse_args()
    if args.op == "cutout":
        make_studio_cutout(args.src, args.out, tolerance=args.tolerance, feather=args.feather)
    elif args.op == "portal":
        vals = tuple(float(v) for v in args.box.split(","))
        punch_portal_mask(args.src, args.out, box_frac=vals)
    elif args.op == "split-sky":
        split_horizon_cutout(args.src, args.out, sky_tolerance=args.tolerance)
    elif args.op == "refract":
        render_refracted_glass_dock(args.src, args.out)
    print(f"Studio asset op '{args.op}' -> {args.out}")


if __name__ == "__main__":
    main()
