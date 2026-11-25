#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Photorealistic Studio Cutout, Architectural Portal Punch, Sky Splitter,
Backdrop-Refracted Glass, Hollow Picture Frame, Organic Ink Drips, Cloud Sweeps,
Horizon Silhouettes & Broadsheet Newspaper Engine for morphdeck-3d.
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
            candidates.append("/System/Library/Fonts/Supplemental/Georgia Bold.ttf" if "bold" in style.lower() else "/System/Library/Fonts/Supplemental/Georgia.ttf")
        candidates.append("/System/Library/Fonts/Supplemental/Arial Bold.ttf" if weight >= 600 else "/System/Library/Fonts/Supplemental/Arial.ttf")
    else:
        if "georgia" in style.lower():
            candidates.append("/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf" if "bold" in style.lower() else "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf")
        candidates.append("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if weight >= 600 else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")

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


def _flood_outer_bg(cand_u8: np.ndarray, seed_top_row: bool = False) -> np.ndarray:
    """Fast C-level perimeter flood fill via Pillow RGB floodfill on a padded binary mask."""
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


def punch_portal_mask(
    photo_path: str,
    out_path: str,
    box_frac: tuple[float, float, float, float] = (0.26, 0.17, 0.74, 0.83),
    radius: int = 22,
    bevel: bool = True,
) -> str:
    """Punch a transparent architectural window/doorway aperture out of a real photograph."""
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
        d.rounded_rectangle([x0 - 2, y0 - 2, x1 + 2, y1 + 2], radius=radius + 2, outline=(12, 14, 18, 225), width=5)
        d.rounded_rectangle([x0, y0, x1, y1], radius=radius, outline=(255, 255, 255, 128), width=2)

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
    """Split a real landscape/architecture photo into a transparent sky layer over opaque terrain."""
    im = Image.open(photo_path).convert("RGBA")
    if size:
        im = cover_crop(im.convert("RGB"), size[0], size[1]).convert("RGBA")
    w, h = im.size

    if ridge_frac is not None:
        yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
        xn = xx / float(w)
        yn = yy / float(h)
        ridge_y = float(ridge_frac) + 0.018 * np.sin(xn * 9.0 + 0.6) - 0.025 * np.exp(-((xn - 0.5) ** 2) / 0.03)
        alpha = np.clip((yn - ridge_y) / 0.004, 0.0, 1.0) * 255.0
        arr = np.array(im)
        arr[..., 3] = alpha.astype(np.uint8)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        Image.fromarray(arr, "RGBA").save(out_path, "PNG")
        return out_path

    arr = np.asarray(im, dtype=np.float32)
    sky_sample = np.median(arr[: max(4, h // 18), :, :3].reshape(-1, 3), axis=0)
    dist = np.linalg.norm(arr[..., :3] - sky_sample[None, None, :], axis=-1)
    cand = np.where((dist <= sky_tolerance) & (np.arange(h)[:, None] < int(h * 0.78)), 0, 255).astype(np.uint8)
    is_sky = _flood_outer_bg(cand, seed_top_row=True)

    alpha_u8 = np.where(is_sky, 0, 255).astype(np.uint8)
    alpha_im = Image.fromarray(alpha_u8, "L").filter(ImageFilter.GaussianBlur(1.2))
    out_rgba = np.dstack([arr[..., :3].astype(np.uint8), np.asarray(alpha_im)])
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    Image.fromarray(out_rgba, "RGBA").save(out_path, "PNG")
    return out_path


def render_refracted_glass_dock(
    bg_path: str,
    out_path: str,
    crop_box: tuple[int, int, int, int] = (80, 772, 1840, 1028),
    blur_radius: int = 24,
    tint_rgba: tuple[int, int, int, int] = (18, 22, 34, 138),
    radius: int = 34,
) -> str:
    """Create a genuine backdrop-refracted glass dock cropped and blurred directly from bg_path."""
    bg = cover_crop(Image.open(bg_path).convert("RGB"), 1920, 1080)
    w = crop_box[2] - crop_box[0]
    h = crop_box[3] - crop_box[1]
    sliced = bg.crop(crop_box).filter(ImageFilter.GaussianBlur(blur_radius))
    glass = Image.alpha_composite(sliced.convert("RGBA"), Image.new("RGBA", (w, h), tint_rgba))

    iw, ih = w - 12, h - 16
    mask = Image.new("L", (iw * 2, ih * 2), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, iw * 2 - 1, ih * 2 - 1], radius=radius * 2, fill=244)
    mask = mask.resize((iw, ih), Image.Resampling.LANCZOS)

    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([12, 14, w - 12, h - 6], radius=radius + 2, fill=(0, 0, 0, 160))
    canvas = Image.alpha_composite(canvas, sh.filter(ImageFilter.GaussianBlur(8)))
    canvas.paste(glass.crop((6, 6, w - 6, h - 10)), (6, 6), mask)

    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([6, 6, w - 6, h - 10], radius=radius, outline=(255, 255, 255, 86), width=2)
    d.line([(6 + radius, 8), (w - 6 - radius, 8)], fill=(255, 255, 255, 175), width=2)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    canvas.save(out_path, "PNG")
    return out_path


def render_organic_ink_drips(
    out_path: str,
    width: int = 1920,
    height: int = 380,
    color_hex: str = "#B91C1C",
    start_x_frac: float = 0.0,
) -> str:
    """
    Generate organic wavy liquid ink drips hanging from the top edge, inspired by
    classic animated editorial print sequences (replacing rigid bars).
    """
    rgb = hex_to_rgb(color_hex, (185, 28, 28))
    rng = np.random.default_rng(74)
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)

    x0 = int(width * start_x_frac)
    y_profile = np.full(width, 28.0, dtype=np.float32)
    # Multi-frequency smooth organic waves + rounded hanging lobes
    x_norm = np.linspace(0, 1, width, dtype=np.float32)
    y_profile += 24.0 * np.sin(x_norm * 22.0 + 0.8) + 16.0 * np.cos(x_norm * 47.0)
    for center_f, amp, sigma in [
        (0.18, 125.0, 0.016),
        (0.31, 185.0, 0.019),
        (0.46, 145.0, 0.017),
        (0.58, 225.0, 0.022),
        (0.71, 165.0, 0.018),
        (0.84, 240.0, 0.021),
        (0.93, 150.0, 0.015),
    ]:
        dist2 = ((x_norm - center_f) ** 2) / (sigma ** 2)
        y_profile += amp * np.exp(-dist2)

    alpha = np.clip((y_profile[None, :] - yy) * 0.55, 0.0, 1.0) * 246.0
    if x0 > 0:
        fade_x = np.clip((xx - x0) / 80.0, 0.0, 1.0)
        alpha = alpha * fade_x

    # Subtle fluid pigment grain
    grain = rng.integers(-6, 7, size=(height, width, 1), dtype=np.int16)
    col_arr = np.clip(np.full((height, width, 3), rgb, dtype=np.int16) + grain, 0, 255).astype(np.uint8)
    rgba = np.dstack([col_arr, alpha.astype(np.uint8)])
    canvas = Image.fromarray(rgba, mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    canvas.save(out_path, "PNG")
    return out_path


def render_soft_paper_cloud(
    out_path: str,
    width: int = 1920,
    height: int = 620,
    paper_hex: str = "#EFE8D8",
) -> str:
    """Create a soft feathered paper-toned reading window floating over dense newspaper texture."""
    rgb = hex_to_rgb(paper_hex, (239, 232, 216))
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
    nx = (xx - width * 0.50) / (width * 0.44)
    ny = (yy - height * 0.50) / (height * 0.42)
    r2 = nx ** 2 + ny ** 2
    alpha = np.clip((1.0 - r2) * 1.45, 0.0, 1.0)
    alpha_u8 = (np.power(alpha, 0.7) * 238.0).astype(np.uint8)
    rgb_arr = np.full((height, width, 3), rgb, dtype=np.uint8)
    im = Image.fromarray(np.dstack([rgb_arr, alpha_u8]), mode="RGBA").filter(ImageFilter.GaussianBlur(14))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    im.save(out_path, "PNG")
    return out_path


def render_cloud_sweep(
    out_path: str,
    width: int = 1920,
    height: int = 1080,
    tint_rgb: tuple[int, int, int] = (242, 245, 250),
) -> str:
    """Create a diagonal volumetric cloud bank used as an airborne Morph fly-by wipe."""
    rng = np.random.default_rng(19)
    yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
    xn = xx / float(width)
    yn = yy / float(height)

    density = np.zeros((height, width), dtype=np.float32)
    for cx, cy, rx, ry, wgt in [
        (0.28, 0.52, 0.24, 0.30, 0.90),
        (0.46, 0.44, 0.28, 0.34, 1.00),
        (0.64, 0.52, 0.26, 0.31, 0.95),
        (0.80, 0.40, 0.22, 0.28, 0.80),
        (0.38, 0.68, 0.24, 0.24, 0.75),
        (0.60, 0.66, 0.24, 0.24, 0.72),
    ]:
        d2 = ((xn - cx) / rx) ** 2 + ((yn - cy) / ry) ** 2
        density += wgt * np.exp(-d2 * 1.65)

    noise = rng.normal(0.0, 0.06, size=(height, width)).astype(np.float32)
    alpha = np.clip((density + noise) * 0.78, 0.0, 1.0)
    alpha_im = Image.fromarray((alpha * 225.0).astype(np.uint8), mode="L").filter(ImageFilter.GaussianBlur(16))

    # Subtle bottom atmospheric shading inside the cloud bank
    shade = np.clip(1.0 - (yn - 0.35) * 0.18, 0.84, 1.0)[..., None]
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
    Create a crisp high-contrast crimson silhouette skyline/landscape horizon strip
    for the lower edge of odometer counter slides.
    """
    rgb = hex_to_rgb(color_hex, (185, 28, 28))
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    d = ImageDraw.Draw(canvas)

    # Undulating rocky ridge + subtle architectural monolith peaks
    pts = [(0, height)]
    for x in range(0, width + 16, 16):
        t = x / float(width)
        y = int(
            height * 0.62
            - 36 * np.sin(t * 7.5 + 0.4)
            - 22 * np.cos(t * 17.0)
            - (46 if 0.22 < t < 0.26 else 0)
            - (58 if 0.72 < t < 0.75 else 0)
        )
        pts.append((x, max(28, min(height - 16, y))))
    pts.append((width, height))
    d.polygon(pts, fill=(*rgb, 245))
    # Subtle glow gradient along ridge
    glow = canvas.filter(ImageFilter.GaussianBlur(6))
    glow_arr = np.array(glow)
    glow_arr[..., 3] = (glow_arr[..., 3].astype(np.float32) * 0.38).astype(np.uint8)
    out = Image.alpha_composite(Image.fromarray(glow_arr, "RGBA"), canvas)
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
    Create a HOLLOW picture frame RGBA overlay whose center aperture is 100% transparent,
    enabling true 4.4x concentric pull-back Morph transitions without occluding the live poster.
    """
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    d = ImageDraw.Draw(canvas)
    mat_px = max(24, border_px // 2)
    in_pad = border_px + mat_px

    # Outer frame molding
    if frame_style == "ornate_gold":
        d.rectangle([0, 0, width - 1, height - 1], fill=(108, 78, 26, 255), outline=(58, 40, 12, 255), width=6)
        d.rectangle([16, 16, width - 17, height - 17], outline=(218, 178, 86, 255), width=14)
        d.rectangle([38, 38, width - 39, height - 39], outline=(162, 122, 46, 255), width=16)
        d.rectangle([border_px - 6, border_px - 6, width - border_px + 5, height - border_px + 5], outline=(235, 204, 122, 255), width=6)
        # Cream museum mat board
        d.rectangle([border_px, border_px, width - border_px - 1, height - border_px - 1], fill=(242, 237, 226, 255))
        # Inner bevel line
        d.rectangle([in_pad - 3, in_pad - 3, width - in_pad + 2, height - in_pad + 2], outline=(120, 96, 54, 255), width=4)
    else:
        d.rectangle([0, 0, width - 1, height - 1], fill=(20, 22, 26, 255), outline=(58, 62, 72, 255), width=6)
        d.rectangle([border_px, border_px, width - border_px - 1, height - border_px - 1], fill=(244, 244, 244, 255))
        d.rectangle([in_pad - 2, in_pad - 2, width - in_pad + 1, height - in_pad + 1], outline=(180, 184, 190, 255), width=3)

    # Punch transparent aperture strictly in the center!
    arr = np.array(canvas)
    arr[in_pad : height - in_pad, in_pad : width - in_pad, 3] = 0
    out_im = Image.fromarray(arr, mode="RGBA")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    out_im.save(out_path, "PNG")
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
    """Synthesize a high-resolution editorial broadsheet newspaper texture with serif masthead and halftone photos."""
    w, h = size
    paper_rgb = hex_to_rgb(paper_hex, (236, 228, 211))
    acc_rgb = hex_to_rgb(accent_hex, (185, 28, 28))
    rng = np.random.default_rng(1945)

    noise = rng.integers(-8, 9, size=(h, w, 1), dtype=np.int16)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    vig = ((((xx - w * 0.5) / (w * 0.65)) ** 2 + ((yy - h * 0.5) / (h * 0.65)) ** 2) * 18).clip(0, 26).astype(np.int16)[..., None]
    base_arr = np.clip(np.full((h, w, 3), paper_rgb, dtype=np.int16) + noise - vig, 0, 255).astype(np.uint8)
    img = Image.fromarray(base_arr, "RGB")
    d = ImageDraw.Draw(img)

    ink = (22, 22, 24)
    muted = (84, 78, 70)
    rule_col = (56, 52, 48)

    f_date = load_font("georgia", 24)
    f_mast = load_font("georgia_bold", 94)
    f_sub = load_font("georgia_bold", 34)
    f_col_t = load_font("georgia_bold", 24)
    f_body = load_font("georgia", 18)

    mx = 96
    d.line([(mx, 64), (w - mx, 64)], fill=ink, width=4)
    dbb = d.textbbox((0, 0), dateline.upper(), font=f_date)
    d.text((w // 2 - (dbb[2] - dbb[0]) // 2, 76), dateline.upper(), font=f_date, fill=muted)
    d.line([(mx, 114), (w - mx, 114)], fill=rule_col, width=2)

    h_lines = textwrap.wrap(headline.upper(), width=26)[:2]
    hy = 128
    for ln in h_lines:
        lbb = d.textbbox((0, 0), ln, font=f_mast)
        d.text((w // 2 - (lbb[2] - lbb[0]) // 2, hy), ln, font=f_mast, fill=ink)
        hy += 102

    sbb = d.textbbox((0, 0), subhead.upper(), font=f_sub)
    d.text((w // 2 - (sbb[2] - sbb[0]) // 2, hy + 8), subhead.upper(), font=f_sub, fill=acc_rgb)
    top_body = hy + 68
    d.line([(mx, top_body - 12), (w - mx, top_body - 12)], fill=ink, width=4)

    n_cols = 3
    col_gap = 42
    col_w = (w - 2 * mx - col_gap * (n_cols - 1)) // n_cols
    col_texts = columns if columns else [
        "DISPATCH ARCHIVE // Strategic industrial mobilization reshaped global logistics networks across four consecutive quarters while maintaining strict delivery tolerances.",
        "FIELD REPORT // Autonomous telemetry and synchronized staging turned fragmented regional operations into unified continental throughput.",
        "SPECIAL EDITION // Primary witnesses and structural verification records confirm sustained resilience across extreme thermal and mechanical stress profiles.",
    ]

    valid_photos = [p for p in (photo_paths or []) if p and os.path.exists(p)]
    for ci in range(n_cols):
        cx = mx + ci * (col_w + col_gap)
        if ci > 0:
            d.line([(cx - col_gap // 2, top_body), (cx - col_gap // 2, h - 72)], fill=(140, 132, 120), width=2)

        cy = top_body + 10
        if ci == 1 and valid_photos:
            ph = ImageEnhance.Contrast(cover_crop(Image.open(valid_photos[0]).convert("L"), col_w, 340).convert("RGB")).enhance(1.15)
            tint = Image.new("RGB", ph.size, paper_rgb)
            img.paste(Image.blend(ph, tint, 0.14), (cx, cy))
            d.rectangle([cx, cy, cx + col_w, cy + 340], outline=ink, width=3)
            cy += 364

        txt = col_texts[ci % len(col_texts)]
        d.text((cx, cy), f"SECTION 0{ci + 1}  //  RECORD", font=f_col_t, fill=ink)
        cy += 36
        para = (txt + " ") * 4
        for line in textwrap.wrap(para, width=max(28, col_w // 10)):
            if cy > h - 90:
                break
            d.text((cx, cy), line, font=f_body, fill=(38, 36, 34))
            cy += 26

    d.line([(mx, h - 62), (w - mx, h - 62)], fill=ink, width=3)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    img.save(out_path, "PNG")
    return out_path


def make_contact_shadow(out_path: str, width: int = 640, height: int = 140, opacity: int = 180) -> str:
    sh = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    ImageDraw.Draw(sh).ellipse(
        [int(width * 0.10), int(height * 0.24), int(width * 0.90), int(height * 0.76)],
        fill=(0, 0, 0, opacity),
    )
    sh = sh.filter(ImageFilter.GaussianBlur(radius=int(height * 0.18)))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    sh.save(out_path, "PNG")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Photorealistic studio cutouts, portals, horizons, and archival textures.")
    sub = parser.add_subparsers(dest="op", required=True)

    p1 = sub.add_parser("cutout", help="Segment an isolated studio image to transparent RGBA PNG")
    p1.add_argument("--src", required=True)
    p1.add_argument("--out", required=True)
    p1.add_argument("--tolerance", type=float, default=28.0)
    p1.add_argument("--feather", type=float, default=1.2)

    p2 = sub.add_parser("portal", help="Punch an architectural window aperture out of a real photograph")
    p2.add_argument("--src", required=True)
    p2.add_argument("--out", required=True)
    p2.add_argument("--box", default="0.26,0.17,0.74,0.83", help="x0,y0,x1,y1 fractions")

    p3 = sub.add_parser("split-sky", help="Separate sky from foreground landscape/architecture photograph")
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
