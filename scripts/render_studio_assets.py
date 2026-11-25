#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Theme-Adaptive, Subject-Linked 3D (.glb) & Frosted-Glass Studio Asset Engine
for morphdeck-3d.

Generates production-grade visual assets for any project:
- 10 procedural subject-linked 3D `.glb` archetypes (or custom user `.glb`)
- 7 visual theme systems
- Multi-layered cinematic horizon landscapes when external photography is omitted
  (and color-graded foreground horizon cutouts when photography is supplied)
- Translucent frosted-glass containers with Gaussian drop shadows and specular rims
- 2x-supersampled translucent frosted-glass charts & system architecture diagrams
"""

import argparse
import copy
import json
import math
import os
import shutil
import sys
import textwrap
import numpy as np
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from bake_glb_screen import bake_glb_screen  # noqa: E402
from generate_subject_glb import ARCHETYPES, canonical_archetype, generate_subject_glb  # noqa: E402


THEMES = {
    "obsidian_kinetic": {
        "is_dark": True,
        "bg": "#080A0F",
        "bg_alt": "#0D111A",
        "surface": "#141824",
        "surface_alt": "#1C2233",
        "ink": "#F8FAFC",
        "ink_muted": "#94A3B8",
        "border": "#2E384D",
        "default_accents": ["#E07A5F", "#38BDF8", "#F59E0B"],
        "display_font": "Bahnschrift",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "smartphone_pro",
        "scene_mood": "golden_city",
    },
    "editorial_ivory": {
        "is_dark": False,
        "bg": "#F4EFE6",
        "bg_alt": "#12100E",
        "surface": "#181512",
        "surface_alt": "#241F1B",
        "ink": "#F8F5F0",
        "ink_muted": "#D6CFC5",
        "border": "#4A3F35",
        "default_accents": ["#C2410C", "#D97706", "#0D9488"],
        "display_font": "Georgia",
        "heading_font": "Georgia",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "architectural_pavilion",
        "scene_mood": "travertine_sunset",
    },
    "cyber_blueprint": {
        "is_dark": True,
        "bg": "#070C17",
        "bg_alt": "#0B1324",
        "surface": "#101B30",
        "surface_alt": "#162540",
        "ink": "#F1F5F9",
        "ink_muted": "#94A3B8",
        "border": "#243B61",
        "default_accents": ["#06B6D4", "#818CF8", "#10B981"],
        "display_font": "Bahnschrift",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "ai_quantum_server",
        "scene_mood": "cyber_horizon",
    },
    "kinetic_poster": {
        "is_dark": True,
        "bg": "#0B0C10",
        "bg_alt": "#12131A",
        "surface": "#171923",
        "surface_alt": "#212432",
        "ink": "#FFFFFF",
        "ink_muted": "#A1A1AA",
        "border": "#3F3F46",
        "default_accents": ["#E11D48", "#3B82F6", "#F59E0B"],
        "display_font": "Bahnschrift",
        "heading_font": "Impact",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "product_drop_capsule",
        "scene_mood": "crimson_stage",
    },
    "executive_clean": {
        "is_dark": True,
        "bg": "#0A1120",
        "bg_alt": "#0F172A",
        "surface": "#152036",
        "surface_alt": "#1E2D4A",
        "ink": "#F8FAFC",
        "ink_muted": "#94A3B8",
        "border": "#334155",
        "default_accents": ["#3B82F6", "#14B8A6", "#F59E0B"],
        "display_font": "Bahnschrift",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "executive_globe_monolith",
        "scene_mood": "alpine_twilight",
    },
    "biotech_luminance": {
        "is_dark": True,
        "bg": "#061412",
        "bg_alt": "#091F1C",
        "surface": "#0E2925",
        "surface_alt": "#143832",
        "ink": "#F0FDFA",
        "ink_muted": "#99F6E4",
        "border": "#1F5249",
        "default_accents": ["#10B981", "#06B6D4", "#A3E635"],
        "display_font": "Bahnschrift",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "biotech_helix_core",
        "scene_mood": "emerald_horizon",
    },
    "monochrome_luxury": {
        "is_dark": True,
        "bg": "#09090B",
        "bg_alt": "#111115",
        "surface": "#17171D",
        "surface_alt": "#22222B",
        "ink": "#FAFAFA",
        "ink_muted": "#A1A1AA",
        "border": "#383846",
        "default_accents": ["#D4AF37", "#E5E7EB", "#60A5FA"],
        "display_font": "Georgia",
        "heading_font": "Georgia",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
        "default_archetype": "fintech_vault_prism",
        "scene_mood": "champagne_obsidian",
    },
}


def hex_to_rgb(hex_str: str, default=(37, 99, 235)) -> tuple:
    if not hex_str:
        return default
    h = str(hex_str).strip().lstrip("#")
    if len(h) != 6:
        return default
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb_to_hex(rgb: tuple) -> str:
    return "#{:02X}{:02X}{:02X}".format(
        max(0, min(255, int(rgb[0]))),
        max(0, min(255, int(rgb[1]))),
        max(0, min(255, int(rgb[2]))),
    )


def darken(rgb: tuple, factor: float = 0.78) -> tuple:
    return tuple(max(0, int(c * factor)) for c in rgb)


def load_font(style: str, size: int, weight: int = 400):
    """Cross-platform font resolver (Windows, macOS, Linux)."""
    candidates = []
    if sys.platform.startswith("win"):
        fdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        win_map = {
            "bahnschrift": ["bahnschrift.ttf", "arialbd.ttf"],
            "georgia": ["georgia.ttf", "cambria.ttc", "segoeui.ttf"],
            "georgia_bold": ["georgiab.ttf", "georgia.ttf", "segoeuib.ttf"],
            "consolas": ["consola.ttf", "CascadiaMono.ttf", "segoeui.ttf"],
            "consolas_bold": ["consolab.ttf", "consola.ttf"],
            "impact": ["impact.ttf", "bahnschrift.ttf", "arialbd.ttf"],
            "segoe": ["segoeui.ttf", "arial.ttf"],
            "segoe_bold": ["segoeuib.ttf", "arial.ttf"],
            "segoe_semi": ["seguisb.ttf", "segoeuib.ttf", "arialbd.ttf"],
        }
        candidates.extend(os.path.join(fdir, n) for n in win_map.get(style.lower(), ["segoeui.ttf", "arial.ttf"]))
    elif sys.platform == "darwin":
        if "georgia" in style.lower():
            candidates.append("/System/Library/Fonts/Supplemental/Georgia Bold.ttf" if "bold" in style.lower() else "/System/Library/Fonts/Supplemental/Georgia.ttf")
        elif "consolas" in style.lower():
            candidates.append("/System/Library/Fonts/Menlo.ttc")
        elif "impact" in style.lower():
            candidates.append("/System/Library/Fonts/Supplemental/Impact.ttf")
        candidates.extend([
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if weight >= 600 or "bold" in style.lower() else "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
        ])
    else:
        if "georgia" in style.lower():
            candidates.append("/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf" if "bold" in style.lower() else "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf")
        elif "consolas" in style.lower():
            candidates.append("/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf" if "bold" in style.lower() else "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf")
        candidates.extend([
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if weight >= 600 or "bold" in style.lower() else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if weight >= 600 else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ])

    for fpath in candidates:
        if os.path.exists(fpath):
            try:
                font = ImageFont.truetype(fpath, size)
                if "bahnschrift" in fpath.lower():
                    try:
                        font.set_variation_by_axes([weight, 100])
                    except Exception:
                        pass
                return font
            except Exception:
                continue
    return ImageFont.load_default()


def theme_font(theme_cfg: dict, role: str, size: int, weight: int = 400):
    name = theme_cfg.get(f"{role}_font", "Bahnschrift").lower()
    if "georgia" in name:
        style = "georgia_bold" if weight >= 600 else "georgia"
    elif "impact" in name:
        style = "impact"
    elif "consolas" in name:
        style = "consolas_bold" if weight >= 600 else "consolas"
    elif "segoe" in name:
        style = "segoe_bold" if weight >= 700 else ("segoe_semi" if weight >= 600 else "segoe")
    else:
        style = "bahnschrift"
    return load_font(style, size, weight=weight)


def cover_crop(img: Image.Image, target_w: int, target_h: int, bias_y: float = 0.5) -> Image.Image:
    src_r = img.width / max(1, img.height)
    dst_r = target_w / max(1, target_h)
    if src_r > dst_r:
        new_w = max(1, int(img.height * dst_r))
        off_x = max(0, (img.width - new_w) // 2)
        img = img.crop((off_x, 0, off_x + new_w, img.height))
    else:
        new_h = max(1, int(img.width / dst_r))
        off_y = max(0, min(img.height - new_h, int((img.height - new_h) * bias_y)))
        img = img.crop((0, off_y, img.width, off_y + new_h))
    return img.resize((target_w, target_h), Image.Resampling.LANCZOS)


def round_image(img: Image.Image, radius: int) -> Image.Image:
    img = img.convert("RGBA")
    mask = Image.new("L", (img.width * 2, img.height * 2), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, img.width * 2 - 1, img.height * 2 - 1],
        radius=radius * 2,
        fill=255,
    )
    mask = mask.resize(img.size, Image.Resampling.LANCZOS)
    out = Image.new("RGBA", img.size, (0, 0, 0, 0))
    out.paste(img, (0, 0), mask)
    return out


def draw_star(draw: ImageDraw.ImageDraw, cx: float, cy: float, r_outer: float, fill) -> None:
    pts = []
    r_inner = r_outer * 0.42
    for i in range(10):
        angle = math.radians(-90 + i * 36)
        r = r_outer if i % 2 == 0 else r_inner
        pts.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))
    draw.polygon(pts, fill=fill)


def add_left_gradient(img: Image.Image, base_col=(9, 10, 14), max_alpha=240, fade_frac=0.66, power=1.25) -> Image.Image:
    w, h = img.size
    fade_w = max(1, int(w * fade_frac))
    x_arr = np.linspace(1.0, 0.0, fade_w, dtype=np.float32) ** power
    alpha_full = np.zeros((h, w), dtype=np.uint8)
    alpha_full[:, :fade_w] = (max_alpha * x_arr).astype(np.uint8)[None, :]
    overlay = np.zeros((h, w, 4), dtype=np.uint8)
    overlay[..., 0], overlay[..., 1], overlay[..., 2] = base_col
    overlay[..., 3] = alpha_full
    return Image.alpha_composite(img.convert("RGBA"), Image.fromarray(overlay, "RGBA"))


def add_bottom_gradient(img: Image.Image, base_col=(9, 10, 14), max_alpha=225, start_frac=0.52) -> Image.Image:
    w, h = img.size
    start_y = int(h * start_frac)
    span = max(1, h - start_y)
    y_arr = np.linspace(0.0, 1.0, span, dtype=np.float32) ** 1.35
    alpha_full = np.zeros((h, w), dtype=np.uint8)
    alpha_full[start_y:, :] = (max_alpha * y_arr[:, None]).astype(np.uint8)
    overlay = np.zeros((h, w, 4), dtype=np.uint8)
    overlay[..., 0], overlay[..., 1], overlay[..., 2] = base_col
    overlay[..., 3] = alpha_full
    return Image.alpha_composite(img.convert("RGBA"), Image.fromarray(overlay, "RGBA"))


def render_cinematic_scene(w: int, h: int, c1: tuple, c2: tuple, seed: int = 42) -> Image.Image:
    """Synthesize a rich multi-layer horizon landscape with sky bloom, 3 ridges, skyline, reflections, and bokeh."""
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    yn = yy / max(1.0, float(h))
    xn = xx / max(1.0, float(w))

    horizon_y = 0.60
    sky_top = np.array([10, 14, 24], dtype=np.float32)
    sky_mid = np.array([min(255, int(c1[0] * 0.55 + 24)), min(255, int(c1[1] * 0.45 + 22)), min(255, int(c2[2] * 0.50 + 36))], dtype=np.float32)
    sky_hor = np.array([min(255, int(c1[0] * 0.92 + 55)), min(255, int(c1[1] * 0.72 + 45)), min(255, int(c1[2] * 0.55 + 38))], dtype=np.float32)

    t_sky = np.clip(yn / horizon_y, 0.0, 1.0)[..., None]
    arr = sky_top * (1.0 - t_sky) ** 1.4 + sky_mid * (1.0 - (1.0 - t_sky) ** 1.4)

    sun_x = 0.64 + (seed % 3 - 1) * 0.12
    sun_y = horizon_y - 0.06
    d_sun = np.sqrt(((xn - sun_x) * 1.45) ** 2 + ((yn - sun_y) * 2.1) ** 2)
    bloom = np.exp(-d_sun * 3.2)[..., None]
    core_bloom = np.exp(-(d_sun ** 2) * 28.0)[..., None]
    arr = arr * (1.0 - bloom * 0.55) + sky_hor * (bloom * 0.85) + np.array([255, 242, 215], dtype=np.float32) * (core_bloom * 0.65)

    ray_angle = (xn - sun_x) / (np.abs(yn - sun_y) + 0.18)
    rays = (0.5 + 0.5 * np.sin(ray_angle * 9.0 + seed)) ** 3.0
    ray_mask = (rays * np.exp(-d_sun * 2.2) * 0.18)[..., None]
    arr += sky_hor * ray_mask

    x_1d = np.linspace(0.0, 1.0, w, dtype=np.float32)
    for layer_i, (base_hy, amp, freq, dark_f) in enumerate([
        (horizon_y - 0.11, 0.055, 5.0, 0.42),
        (horizon_y - 0.05, 0.038, 9.0, 0.26),
        (horizon_y, 0.022, 15.0, 0.14),
    ]):
        prof = base_hy + amp * np.sin(x_1d * freq + seed * (layer_i + 1)) + (amp * 0.5) * np.cos(x_1d * freq * 2.3 - seed)
        if layer_i == 2:
            n_blocks = 26
            for bi in range(n_blocks):
                bx0 = int((bi / n_blocks) * w)
                bw = max(6, int(w / (n_blocks * rng.uniform(1.1, 2.2))))
                bh = rng.uniform(0.012, 0.068)
                prof[bx0:min(w, bx0 + bw)] -= bh
        ridge_mask = (yn >= prof[None, :])[..., None]
        ridge_col = sky_top * (1.0 - dark_f) + sky_mid * dark_f
        arr = np.where(ridge_mask, ridge_col, arr)

    water_mask = yn >= horizon_y
    t_water = np.clip((yn - horizon_y) / (1.0 - horizon_y), 0.0, 1.0)[..., None]
    mirror_y = np.clip((2.0 * horizon_y - yn) * h, 0, h - 1).astype(np.int32)
    ripple = (np.sin(yn * 180.0 + xn * 12.0) * 6.0).astype(np.int32)
    mirror_x = np.clip(xx.astype(np.int32) + ripple, 0, w - 1)
    reflected = arr[mirror_y, mirror_x] * (0.72 - 0.48 * t_water)
    deep_water = np.array([8, 11, 18], dtype=np.float32)
    water_arr = reflected + deep_water * (0.28 + 0.52 * t_water)
    arr = np.where(water_mask[..., None], water_arr, arr)

    img = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), "RGB")
    d = ImageDraw.Draw(img, "RGBA")

    hy_px = int(horizon_y * h)
    for _ in range(140):
        lx = int(rng.integers(20, w - 20))
        ly = hy_px - int(rng.integers(-6, 36))
        col = c1 if rng.random() < 0.55 else (c2 if rng.random() < 0.5 else (255, 225, 165))
        r_sz = int(rng.integers(1, 4))
        d.ellipse([lx - r_sz, ly - r_sz, lx + r_sz, ly + r_sz], fill=(*col, 215))
        if ly >= hy_px - 10:
            ry = hy_px + (hy_px - ly) + int(rng.integers(4, 48))
            d.line([(lx - r_sz * 2, ry), (lx + r_sz * 2, ry)], fill=(*col, 95), width=2)

    bokeh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bokeh)
    for _ in range(18):
        bx = int(rng.integers(60, w - 60))
        by = int(rng.integers(int(h * 0.20), int(h * 0.82)))
        br = int(rng.integers(14, 46))
        bcol = c1 if rng.random() < 0.6 else c2
        bd.ellipse([bx - br, by - br, bx + br, by + br], fill=(*bcol, 24))
    bokeh = bokeh.filter(ImageFilter.GaussianBlur(12))
    return Image.alpha_composite(img.convert("RGBA"), bokeh).convert("RGB")


def load_photo_or_fallback(path: str, w: int, h: int, c1: tuple, c2: tuple, bias_y: float = 0.5, seed: int = 42) -> Image.Image:
    if path and os.path.exists(path):
        return cover_crop(Image.open(path).convert("RGB"), w, h, bias_y=bias_y)
    return render_cinematic_scene(w, h, c1, c2, seed=seed)


def draw_frosted_glass_card(w: int, h: int, surf_rgb: tuple, accent_rgb: tuple, radius: int = 26, bar_mode: str = "left", alpha: int = 218) -> Image.Image:
    """Render a translucent frosted glass container with drop shadow, specular top glint, and accent bar."""
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([12, 14, w - 12, h - 6], radius=radius + 2, fill=(0, 0, 0, 148))
    canvas = Image.alpha_composite(canvas, sh.filter(ImageFilter.GaussianBlur(8)))

    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([6, 6, w - 6, h - 10], radius=radius, fill=(*surf_rgb, alpha), outline=(255, 255, 255, 62), width=2)
    d.line([(6 + radius, 8), (w - 6 - radius, 8)], fill=(255, 255, 255, 138), width=2)

    if bar_mode == "left":
        d.rounded_rectangle([6, 26, 15, h - 30], radius=4, fill=(*accent_rgb, 245))
    elif bar_mode == "top":
        d.rounded_rectangle([6 + radius, 6, w - 6 - radius, 14], radius=3, fill=(*accent_rgb, 245))
    elif bar_mode == "glow_border":
        d.rounded_rectangle([6, 6, w - 6, h - 10], radius=radius, outline=(*accent_rgb, 195), width=3)
    return canvas


def normalize_spec_to_v2(raw_spec: dict) -> dict:
    spec = copy.deepcopy(raw_spec)
    theme_name = spec.get("theme", "obsidian_kinetic")
    if theme_name not in THEMES:
        theme_name = "obsidian_kinetic"
    theme_cfg = THEMES[theme_name]

    pal_in = spec.get("palette", {})
    d_acc = theme_cfg["default_accents"]
    acc1 = pal_in.get("accent1") or pal_in.get("accent_primary") or d_acc[0]
    acc2 = pal_in.get("accent2") or pal_in.get("accent_secondary") or d_acc[1]
    acc3 = pal_in.get("accent3") or pal_in.get("accent_tertiary") or d_acc[2]
    acc1_dark = pal_in.get("accent_dark") or rgb_to_hex(darken(hex_to_rgb(acc1), 0.78))

    palette = {
        "bg": pal_in.get("bg") or theme_cfg["bg"],
        "bg_alt": pal_in.get("bg_alt") or theme_cfg["bg_alt"],
        "surface": pal_in.get("surface") or theme_cfg["surface"],
        "surface_alt": pal_in.get("surface_alt") or theme_cfg["surface_alt"],
        "ink": pal_in.get("ink") or theme_cfg["ink"],
        "ink_muted": pal_in.get("ink_muted") or theme_cfg["ink_muted"],
        "border": pal_in.get("border") or theme_cfg["border"],
        "accent1": acc1,
        "accent2": acc2,
        "accent3": acc3,
        "accent1_dark": acc1_dark,
        "chassis_color": pal_in.get("chassis_color") or acc1,
        "ivory": pal_in.get("ivory") or "#F5EFE6",
        "obsidian": pal_in.get("obsidian") or "#090A0D",
    }

    brand_in = spec.get("brand", {})
    name = brand_in.get("name", "Project")
    initials = "".join([w[0] for w in name.split()[:2]]).upper() or name[:2].upper()
    brand = {
        "name": name,
        "title_part1": brand_in.get("title_part1", name.upper()[:8]),
        "title_part2": brand_in.get("title_part2", ""),
        "tagline": brand_in.get("tagline") or brand_in.get("motto", ""),
        "motto": brand_in.get("motto") or brand_in.get("tagline", ""),
        "monogram": brand_in.get("monogram") or spec.get("assets", {}).get("logo_monogram") or initials,
        "logo_path": brand_in.get("logo_path") or spec.get("assets", {}).get("logo_path", ""),
    }

    slides = spec.get("slides")
    is_legacy_v1 = slides is None
    if is_legacy_v1:
        s1, s2, s3, s4 = spec.get("slide1", {}), spec.get("slide2", {}), spec.get("slide3", {}), spec.get("slide4", {})
        sps, dock = spec.get("spotlights", [{}, {}]), spec.get("dock", [])
        s7, s8 = spec.get("slide7", {}), spec.get("slide8", {})
        slides = [
            {"type": "cover_hero", "style": "3d_staged", "overline": s1.get("top_bar", ""), "subtitle": s1.get("subtitle", ""), "bottom_cue": s1.get("bottom_cue", "")},
            {"type": "split_horizon_3d", "giant_word": s2.get("giant_word", "PROJECT"), "left_label": s2.get("left_label", ""), "left_text": s2.get("left_text", ""), "right_label": s2.get("right_label", ""), "right_text": s2.get("right_text", ""), "kpis": s2.get("kpis", [])},
            {"type": "orbit_steps_3d", "variant": "upper_stats", "overline": s3.get("overline", ""), "headline": s3.get("headline", ""), "body": s3.get("body", ""), "stats": s3.get("stats", [])},
            {"type": "orbit_steps_3d", "variant": "obsidian_steps", "overline": s4.get("overline", ""), "headline": s4.get("headline", ""), "steps": s4.get("steps", [])},
            {"type": "spotlight_dock", "spotlight_index": 0, "dock_index": 0, "spotlight": sps[0] if len(sps) > 0 else {}, "dock": dock},
            {"type": "spotlight_dock", "spotlight_index": 1, "dock_index": 1, "spotlight": sps[1] if len(sps) > 1 else {}, "dock": dock},
            {"type": "macbook_studio_3d", "overline": s7.get("overline", ""), "headline": s7.get("headline", ""), "pillars": s7.get("pillars", [])},
            {"type": "closing_finale", "style": "dual_3d", "tagline": s8.get("tagline", ""), "summary": s8.get("summary", ""), "cta": s8.get("cta", "")},
        ]

    hero_in = spec.get("hero_3d", {})
    raw_arch = hero_in.get("archetype") or theme_cfg["default_archetype"]
    if is_legacy_v1 and not hero_in.get("archetype"):
        raw_arch = "smartphone_pro"
    hero_arch = "custom" if raw_arch == "custom" else canonical_archetype(raw_arch)

    needs_laptop = any(s.get("type") == "macbook_studio_3d" for s in slides) or any(
        s.get("type") == "closing_finale" and s.get("style") == "dual_3d" for s in slides
    )
    raw_sec = hero_in.get("secondary_archetype") or ("laptop_studio" if needs_laptop else "")
    second_arch = "" if not raw_sec else ("custom" if raw_sec == "custom" else canonical_archetype(raw_sec))

    return {
        "schema_version": 2,
        "theme": theme_name,
        "theme_cfg": theme_cfg,
        "is_dark": True,
        "theme_light_paper": (theme_name == "editorial_ivory"),
        "brand": brand,
        "palette": palette,
        "hero_3d": {
            "archetype": hero_arch,
            "secondary_archetype": second_arch,
            "custom_glb_path": hero_in.get("custom_glb_path") or spec.get("assets", {}).get("custom_glb_path", ""),
            "custom_mesh": hero_in.get("custom_mesh", ""),
            "custom_material": hero_in.get("custom_material"),
        },
        "fonts": {
            "display": theme_cfg["display_font"],
            "heading": theme_cfg["heading_font"],
            "body": theme_cfg["body_font"],
            "semi": theme_cfg["semi_font"],
            "mono": theme_cfg["mono_font"],
        },
        "assets": spec.get("assets", {}),
        "mobile_ui": spec.get("mobile_ui", {}),
        "desktop_ui": spec.get("desktop_ui", {}),
        "spotlights": spec.get("spotlights", []),
        "dock": spec.get("dock", []),
        "slides": slides,
        "requirements": {
            "needs_iphone": True,
            "needs_macbook": bool(second_arch),
            "has_dual_finale": any(s.get("type") == "closing_finale" and s.get("style") == "dual_3d" for s in slides),
        },
    }


def prepare_logo_badge(resolved: dict, out_path: str) -> Image.Image:
    brand = resolved["brand"]
    logo_path = brand.get("logo_path", "")
    if logo_path and os.path.exists(logo_path):
        im = Image.open(logo_path).convert("RGBA").resize((220, 220), Image.Resampling.LANCZOS)
        badge = Image.new("RGBA", (240, 240), (0, 0, 0, 0))
        badge.paste(im, (10, 10), im)
        badge.save(out_path, "PNG")
        return badge

    theme_name = resolved["theme"]
    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    obs = hex_to_rgb(pal["obsidian"])
    mono_txt = (brand.get("monogram") or "MD")[:3].upper()

    badge = Image.new("RGBA", (240, 240), (0, 0, 0, 0))
    d = ImageDraw.Draw(badge)

    if theme_name in ("editorial_ivory", "monochrome_luxury"):
        d.ellipse([16, 16, 224, 224], fill=(*obs, 255), outline=(*acc1, 255), width=6)
        d.ellipse([30, 30, 210, 210], outline=(255, 255, 255, 110), width=2)
        f = load_font("georgia_bold", 78)
    else:
        cx, cy, r = 120, 120, 104
        pts = [(cx + r * math.cos(math.radians(60 * i - 30)), cy + r * math.sin(math.radians(60 * i - 30))) for i in range(6)]
        d.polygon(pts, fill=(*obs, 255), outline=(*acc1, 255), width=8)
        f = load_font("bahnschrift", 80, 700)

    bb = d.textbbox((0, 0), mono_txt, font=f)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    d.text((120 - tw / 2, 120 - th / 2 - 8), mono_txt, font=f, fill=(248, 250, 252, 255))
    badge.save(out_path, "PNG")
    return badge


def generate_studio_backgrounds_and_glass(resolved: dict, out_dir: str) -> None:
    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    acc2 = hex_to_rgb(pal["accent2"])
    acc3 = hex_to_rgb(pal["accent3"])
    ivory = hex_to_rgb(pal["ivory"])
    obs = hex_to_rgb(pal["obsidian"])
    surf = hex_to_rgb(pal["surface"])
    assets = resolved.get("assets", {})

    rng = np.random.default_rng(42)
    yy, xx = np.mgrid[0:1080, 0:1920]

    r2 = ((xx - 960) / 1100.0) ** 2 + ((yy - 540) / 700.0) ** 2
    vig = (r2 * 14).clip(0, 22).astype(np.int16)[..., None]
    noise = rng.integers(-4, 5, size=(1080, 1920, 1), dtype=np.int16)
    ivory_img = Image.fromarray(np.clip(np.full((1080, 1920, 3), ivory, dtype=np.int16) + noise - vig, 0, 255).astype(np.uint8), "RGB")
    ivory_img.save(os.path.join(out_dir, "bg_ivory_grain.png"), "PNG")

    scene_main = render_cinematic_scene(1920, 1080, acc1, acc2, seed=42)
    dark_main = add_bottom_gradient(
        add_left_gradient(ImageEnhance.Brightness(scene_main).enhance(0.46), base_col=obs, max_alpha=225, fade_frac=0.72),
        base_col=obs, max_alpha=215, start_frac=0.45,
    ).convert("RGB")
    dark_main.save(os.path.join(out_dir, "bg_main.png"), "PNG")

    scene_alt = render_cinematic_scene(1920, 1080, acc2, acc1, seed=77)
    dark_alt = add_bottom_gradient(
        add_left_gradient(ImageEnhance.Brightness(scene_alt).enhance(0.44), base_col=obs, max_alpha=230, fade_frac=0.70),
        base_col=obs, max_alpha=220, start_frac=0.48,
    ).convert("RGB")
    dark_alt.save(os.path.join(out_dir, "bg_alt.png"), "PNG")

    low_src = assets.get("horizon_photo", "")
    up_src = assets.get("upper_photo", "")
    add_bottom_gradient(
        ImageEnhance.Contrast(load_photo_or_fallback(low_src, 1920, 530, acc1, acc2, bias_y=0.58, seed=101)).enhance(1.10),
        base_col=obs, max_alpha=150, start_frac=0.38,
    ).convert("RGB").save(os.path.join(out_dir, "split_lower.png"), "PNG")

    add_left_gradient(
        ImageEnhance.Brightness(load_photo_or_fallback(up_src, 1920, 636, acc2, acc1, bias_y=0.25, seed=202)).enhance(0.84),
        base_col=obs, max_alpha=232, fade_frac=0.66, power=1.15,
    ).convert("RGB").save(os.path.join(out_dir, "split_upper.png"), "PNG")

    add_left_gradient(
        ImageEnhance.Brightness(load_photo_or_fallback(low_src, 1920, 1080, acc1, acc2, bias_y=0.50, seed=303)).enhance(0.62),
        base_col=obs, max_alpha=242, fade_frac=0.56,
    ).convert("RGB").save(os.path.join(out_dir, "bg_finale.png"), "PNG")

    draw_frosted_glass_card(1060, 148, surf, acc1, radius=24, bar_mode="left").save(os.path.join(out_dir, "card_step.png"), "PNG")
    for i, ac in enumerate([acc1, acc2, acc3]):
        draw_frosted_glass_card(730, 192, surf, ac, radius=24, bar_mode="left").save(os.path.join(out_dir, f"card_spec_{i + 1}.png"), "PNG")

    draw_frosted_glass_card(672, 768, surf, acc1, radius=28, bar_mode="top", alpha=224).save(os.path.join(out_dir, "card_bento_hero.png"), "PNG")
    draw_frosted_glass_card(1732, 296, surf, acc1, radius=28, bar_mode="left", alpha=224).save(os.path.join(out_dir, "card_bento_strip.png"), "PNG")
    for i, ac in enumerate([acc1, acc2, acc3]):
        draw_frosted_glass_card(552, 420, surf, ac, radius=24, bar_mode="top", alpha=216).save(os.path.join(out_dir, f"card_bento_tile_{i + 1}.png"), "PNG")
    draw_frosted_glass_card(1024, 368, surf, acc2, radius=24, bar_mode="left", alpha=218).save(os.path.join(out_dir, "card_bento_wide.png"), "PNG")
    draw_frosted_glass_card(1180, 560, surf, acc1, radius=28, bar_mode="left", alpha=222).save(os.path.join(out_dir, "card_timeline_drawer.png"), "PNG")
    draw_frosted_glass_card(596, 780, surf, (100, 116, 139), radius=26, bar_mode="top", alpha=214).save(os.path.join(out_dir, "card_split_left.png"), "PNG")
    draw_frosted_glass_card(596, 780, surf, acc1, radius=26, bar_mode="glow_border", alpha=225).save(os.path.join(out_dir, "card_split_right.png"), "PNG")

    draw_frosted_glass_card(1760, 256, surf, acc1, radius=36, bar_mode="none", alpha=205).save(os.path.join(out_dir, "dock_bg.png"), "PNG")

    gw, gh = 416, 216
    glider = Image.new("RGBA", (gw, gh), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glider)
    gd.rounded_rectangle([4, 4, gw - 4, gh - 4], radius=28, fill=(*acc1, 56), outline=(*acc1, 235), width=3)
    gd.line([(36, 6), (gw - 36, 6)], fill=(255, 255, 255, 175), width=2)
    glider.save(os.path.join(out_dir, "dock_glider.png"), "PNG")

    for fname, w, h in [("shadow_phone.png", 640, 160), ("shadow_laptop.png", 960, 180)]:
        sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        ImageDraw.Draw(sh).ellipse([int(w * 0.12), int(h * 0.25), int(w * 0.88), int(h * 0.75)], fill=(0, 0, 0, 185))
        sh.filter(ImageFilter.GaussianBlur(radius=int(h * 0.18))).save(os.path.join(out_dir, fname), "PNG")


def render_pan_backdrop_and_cutout(resolved: dict, slide: dict, idx: int, out_dir: str) -> None:
    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    acc2 = hex_to_rgb(pal["accent2"])

    bg_out = os.path.join(out_dir, f"s{idx}_pan_bg.png")
    fg_out = os.path.join(out_dir, f"s{idx}_pan_fg.png")

    W, H = 2560, 1440
    src_bg = slide.get("backdrop_image", "")
    if src_bg and os.path.exists(src_bg):
        pan_bg = cover_crop(Image.open(src_bg).convert("RGB"), W, H, bias_y=0.45)
    else:
        pan_bg = render_cinematic_scene(W, H, acc1, acc2, seed=500 + idx)
    pan_bg.save(bg_out, "PNG")

    src_cut = slide.get("subject_cutout", "")
    if src_cut and os.path.exists(src_cut):
        cover_crop(Image.open(src_cut).convert("RGBA"), W, H, bias_y=0.5).save(fg_out, "PNG")
        return

    yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
    xn = xx / float(W)
    yn = yy / float(H)
    pinnacle = 0.085 * np.exp(-((xn - 0.48) ** 2) / 0.018)
    ridge_y = 0.515 - pinnacle + 0.028 * np.sin(xn * 11.0 + idx) + 0.016 * np.cos(xn * 23.0 - idx * 0.7)
    alpha = np.clip((yn - ridge_y) / 0.008, 0.0, 1.0)
    fg_arr = np.dstack([np.array(pan_bg, dtype=np.uint8), (alpha * 255).astype(np.uint8)])
    Image.fromarray(fg_arr, "RGBA").save(fg_out, "PNG")


def render_chart_infographic(resolved: dict, slide: dict, idx: int, out_dir: str) -> None:
    pal = resolved["palette"]
    tcfg = resolved["theme_cfg"]
    surf = hex_to_rgb(pal["surface"])
    surf_alt = hex_to_rgb(pal["surface_alt"])
    border = hex_to_rgb(pal["border"])
    ink = (248, 250, 252)
    muted = (148, 163, 184)
    accents = [hex_to_rgb(pal["accent1"]), hex_to_rgb(pal["accent2"]), hex_to_rgb(pal["accent3"])]

    W, H = 2360, 1520
    img = draw_frosted_glass_card(W, H, surf, accents[0], radius=48, bar_mode="none", alpha=216)
    d = ImageDraw.Draw(img)

    f_num = theme_font(tcfg, "display", 92, 700)
    f_lbl = theme_font(tcfg, "display", 42, 700)
    f_sub = theme_font(tcfg, "body", 34, 400)
    f_mono = load_font("consolas_bold", 34)

    chart_type = slide.get("chart_type", "comparison_bars")
    items = slide.get("items", [])[:4]

    if chart_type == "radial_rings":
        n = max(1, len(items))
        slot_w = (W - 160) // n
        cy = 620
        ring_r = 215 if n <= 3 else 175
        ring_w = 42 if n <= 3 else 36

        for j, it in enumerate(items):
            cx = 80 + slot_w * j + slot_w // 2
            pct = max(0.0, min(100.0, float(it.get("value_pct", 75))))
            col = accents[j % len(accents)]
            bbox = [cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r]
            d.arc(bbox, start=0, end=360, fill=(*surf_alt, 255), width=ring_w)
            d.arc(bbox, start=-90, end=-90 + int(3.6 * pct), fill=(*col, 255), width=ring_w)

            disp = str(it.get("display_value", f"{int(pct)}%"))[:10]
            bb = d.textbbox((0, 0), disp, font=f_num)
            d.text((cx - (bb[2] - bb[0]) / 2, cy - (bb[3] - bb[1]) / 2 - 12), disp, font=f_num, fill=ink)

            lbl = str(it.get("label", ""))[:22]
            lbb = d.textbbox((0, 0), lbl, font=f_lbl)
            d.text((cx - (lbb[2] - lbb[0]) / 2, cy + ring_r + 64), lbl, font=f_lbl, fill=col)

            sub = str(it.get("sublabel", ""))[:26]
            sbb = d.textbbox((0, 0), sub, font=f_sub)
            d.text((cx - (sbb[2] - sbb[0]) / 2, cy + ring_r + 132), sub, font=f_sub, fill=muted)
    else:
        n = max(1, len(items))
        top_y = 110
        row_h = (H - 200) // n
        bar_x0 = 100
        bar_max_w = W - 520

        d.rounded_rectangle([W - 760, 44, W - 460, 100], radius=24, fill=surf_alt, outline=border, width=2)
        d.ellipse([W - 736, 62, W - 716, 82], fill=muted)
        d.text((W - 700, 52), "BASELINE", font=f_mono, fill=muted)

        d.rounded_rectangle([W - 430, 44, W - 90, 100], radius=24, fill=(*accents[0], 48), outline=accents[0], width=2)
        d.ellipse([W - 406, 62, W - 386, 82], fill=accents[0])
        d.text((W - 370, 52), "SOLUTION", font=f_mono, fill=accents[0])

        max_val = max([max(float(it.get("value", 1)), float(it.get("baseline", 1))) for it in items] + [1.0])
        for j, it in enumerate(items):
            ry = top_y + j * row_h + 24
            col = accents[j % len(accents)]
            val = float(it.get("value", 0))
            base = float(it.get("baseline", 0))
            unit = str(it.get("unit", ""))
            delta = str(it.get("delta", ""))

            d.text((bar_x0, ry), str(it.get("label", ""))[:36], font=f_lbl, fill=ink)
            if delta:
                dbb = d.textbbox((0, 0), delta, font=f_mono)
                pw = (dbb[2] - dbb[0]) + 48
                d.rounded_rectangle([W - 100 - pw, ry - 4, W - 100, ry + 52], radius=14, fill=(*col, 46), outline=col, width=2)
                d.text((W - 100 - pw + 24, ry + 6), delta, font=f_mono, fill=col)

            bw = max(40, int((base / max_val) * bar_max_w))
            d.rounded_rectangle([bar_x0, ry + 78, bar_x0 + bw, ry + 118], radius=14, fill=surf_alt)
            d.text((bar_x0 + bw + 24, ry + 78), f"{it.get('baseline', '')} {unit}".strip(), font=f_sub, fill=muted)

            vw = max(60, int((val / max_val) * bar_max_w))
            d.rounded_rectangle([bar_x0, ry + 134, bar_x0 + vw, ry + 190], radius=16, fill=col)
            d.text((bar_x0 + vw + 24, ry + 140), f"{it.get('value', '')} {unit}".strip(), font=f_lbl, fill=ink)

    img.resize((1180, 760), Image.Resampling.LANCZOS).save(os.path.join(out_dir, f"s{idx}_chart.png"), "PNG")


def render_architecture_diagram(resolved: dict, slide: dict, idx: int, out_dir: str) -> None:
    pal = resolved["palette"]
    tcfg = resolved["theme_cfg"]
    surf = hex_to_rgb(pal["surface"])
    surf_alt = hex_to_rgb(pal["surface_alt"])
    border = hex_to_rgb(pal["border"])
    ink = (248, 250, 252)
    muted = (148, 163, 184)
    accents = [hex_to_rgb(pal["accent1"]), hex_to_rgb(pal["accent2"]), hex_to_rgb(pal["accent3"])]

    W, H = 2560, 1592
    img = draw_frosted_glass_card(W, H, surf, accents[0], radius=44, bar_mode="none", alpha=216)
    d = ImageDraw.Draw(img)

    layers = slide.get("layers", [])[:4]
    flows = slide.get("flows", [])
    n_layers = max(1, len(layers))

    f_layer = load_font("consolas_bold", 28)
    f_node_t = theme_font(tcfg, "display", 38, 700)
    f_node_s = theme_font(tcfg, "body", 28, 400)
    f_tag = load_font("consolas_bold", 24)

    margin_x, top_pad = 76, 64
    gap_y = 84 if n_layers <= 3 else 56
    layer_h = (H - top_pad * 2 - gap_y * (n_layers - 1)) // n_layers

    for li, layer in enumerate(layers):
        ly = top_pad + li * (layer_h + gap_y)
        col = accents[li % len(accents)]
        d.rounded_rectangle([margin_x, ly, W - margin_x, ly + layer_h], radius=24, fill=(*surf_alt, 215), outline=(*col, 175), width=3)
        d.text((margin_x + 34, ly + 22), str(layer.get("name", f"LAYER {li + 1}")).upper()[:34], font=f_layer, fill=col)

        nodes = layer.get("nodes", [])[:4]
        n_nodes = max(1, len(nodes))
        inner_top = ly + 74
        inner_h = layer_h - 96
        inner_gap = 36
        node_w = ((W - margin_x * 2 - 68) - inner_gap * (n_nodes - 1)) // n_nodes

        for ni, nd in enumerate(nodes):
            nx = margin_x + 34 + ni * (node_w + inner_gap)
            d.rounded_rectangle([nx, inner_top, nx + node_w, inner_top + inner_h], radius=18, fill=(*surf, 245), outline=(*border, 255), width=3)
            d.rounded_rectangle([nx, inner_top, nx + 10, inner_top + inner_h], radius=4, fill=col)
            title = str(nd.get("title", ""))[:26]
            sub = str(nd.get("subtitle", ""))
            tag = str(nd.get("tag", ""))[:18]

            d.text((nx + 32, inner_top + 22), title, font=f_node_t, fill=ink)
            for wi, line in enumerate(textwrap.wrap(sub, width=max(18, int((node_w - 60) / 17)))[:2]):
                d.text((nx + 32, inner_top + 78 + wi * 38), line, font=f_node_s, fill=muted)

            if tag and inner_h >= 165:
                tbb = d.textbbox((0, 0), tag, font=f_tag)
                tw = (tbb[2] - tbb[0]) + 32
                tx, ty = nx + 32, inner_top + inner_h - 52
                d.rounded_rectangle([tx, ty, tx + tw, ty + 36], radius=10, fill=(*col, 42), outline=col, width=2)
                d.text((tx + 16, ty + 5), tag, font=f_tag, fill=col)

        if li < n_layers - 1:
            ay1, ay2 = ly + layer_h + 6, ly + layer_h + gap_y - 6
            mid_y = (ay1 + ay2) // 2
            for ax in (W // 4, W // 2, (W * 3) // 4):
                d.line([(ax, ay1), (ax, ay2 - 12)], fill=col, width=5)
                d.polygon([(ax, ay2), (ax - 12, ay2 - 18), (ax + 12, ay2 - 18)], fill=col)
            flow_lbl = flows[li] if li < len(flows) else ""
            if flow_lbl:
                fbb = d.textbbox((0, 0), flow_lbl, font=f_tag)
                fw = (fbb[2] - fbb[0]) + 44
                d.rounded_rectangle([W // 2 - fw // 2, mid_y - 20, W // 2 + fw // 2, mid_y + 20], radius=16, fill=(*surf, 255), outline=col, width=3)
                d.text((W // 2 - (fbb[2] - fbb[0]) // 2, mid_y - 13), flow_lbl, font=f_tag, fill=col)

    img.resize((1280, 796), Image.Resampling.LANCZOS).save(os.path.join(out_dir, f"s{idx}_arch.png"), "PNG")


def render_editorial_visual(resolved: dict, slide: dict, idx: int, out_dir: str, out_name: str = None) -> None:
    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    acc2 = hex_to_rgb(pal["accent2"])
    obs = hex_to_rgb(pal["obsidian"])
    out_path = os.path.join(out_dir, out_name or f"s{idx}_split_visual.png")
    src = slide.get("image", "")
    if src and os.path.exists(src):
        im = cover_crop(Image.open(src).convert("RGB"), 960, 1080, bias_y=0.35)
    else:
        im = render_cinematic_scene(960, 1080, acc1, acc2, seed=700 + idx)

    framed = add_bottom_gradient(im, base_col=obs, max_alpha=210, start_frac=0.62)
    r_im = round_image(framed, 32)
    d = ImageDraw.Draw(r_im)
    d.rounded_rectangle([2, 2, 957, 1077], radius=32, outline=(255, 255, 255, 78), width=3)
    fig_lbl = str(slide.get("figure_caption", f"{resolved['brand']['name'].upper()}  //  EDITION 0{idx}"))[:48]
    f_cap = load_font("consolas_bold", 22)
    d.rounded_rectangle([44, 976, 916, 1036], radius=18, fill=(*obs, 215), outline=(*acc1, 185), width=2)
    d.text((74, 994), fig_lbl, font=f_cap, fill=(248, 250, 252, 255))
    r_im.save(out_path, "PNG")


def build_iphone_screen(resolved: dict, logo_badge: Image.Image, out_path: str) -> None:
    custom = resolved.get("assets", {}).get("mobile_ui_image", "")
    W, H = 960, 2020
    img = Image.new("RGB", (W, H), (5, 6, 8))
    d = ImageDraw.Draw(img)

    if custom and os.path.exists(custom):
        ui = cover_crop(Image.open(custom).convert("RGB"), W - 52, H - 52)
        ui_r = round_image(ui, 88)
        img.paste(ui_r, (26, 26), ui_r)
        d.rounded_rectangle([355, 58, 605, 110], radius=26, fill=(0, 0, 0))
        img.save(out_path, "PNG")
        return

    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    gold = hex_to_rgb(pal["accent3"])
    mob = resolved.get("mobile_ui", {})
    brand_name = resolved["brand"]["name"]

    d.rounded_rectangle([26, 26, W - 26, H - 26], radius=88, fill=(9, 10, 13))
    d.rounded_rectangle([355, 58, 605, 110], radius=26, fill=(0, 0, 0))
    d.ellipse([562, 72, 584, 94], fill=(18, 22, 34))

    f_time = load_font("bahnschrift", 30, 700)
    f_brand = load_font("bahnschrift", 34, 700)
    f_badge = load_font("segoe_semi", 21)
    f_hero = load_font("bahnschrift", 40, 700)
    f_sub = load_font("segoe", 23)
    f_chip = load_font("segoe_semi", 21)
    f_card_t = load_font("bahnschrift", 29, 700)
    f_card_s = load_font("segoe", 21)
    f_price = load_font("bahnschrift", 26, 700)

    d.text((104, 72), "9:41", font=f_time, fill=(248, 250, 252))
    for i, bh in enumerate([10, 15, 20, 25]):
        bx = 750 + i * 11
        d.rounded_rectangle([bx, 102 - bh, bx + 7, 102], radius=2, fill=(248, 250, 252))
    d.rounded_rectangle([808, 78, 858, 102], radius=6, outline=(148, 163, 184), width=2)
    d.rounded_rectangle([812, 82, 850, 98], radius=3, fill=(248, 250, 252))

    logo_sm = logo_badge.resize((58, 58), Image.Resampling.LANCZOS)
    img.paste(logo_sm, (62, 150), logo_sm)
    d.text((134, 159), brand_name[:18], font=f_brand, fill=(248, 250, 252))

    loc_label = mob.get("location_badge", "Live")[:12]
    d.rounded_rectangle([710, 154, 896, 204], radius=25, fill=(26, 28, 36), outline=acc1, width=2)
    lb = d.textbbox((0, 0), loc_label, font=f_badge)
    d.text((803 - (lb[2] - lb[0]) / 2, 164), loc_label, font=f_badge, fill=acc1)

    hero_src = resolved.get("assets", {}).get("upper_photo", "")
    hero = add_bottom_gradient(
        ImageEnhance.Brightness(load_photo_or_fallback(hero_src, 836, 400, acc1, (20, 30, 50), seed=55)).enhance(0.66),
        max_alpha=240, start_frac=0.25,
    )
    hero_r = round_image(hero, 36)
    img.paste(hero_r, (62, 240), hero_r)
    d.rounded_rectangle([62, 240, 898, 640], radius=36, outline=(60, 66, 82), width=2)
    d.text((96, 396), mob.get("hero_line1", brand_name)[:22], font=f_hero, fill=(255, 255, 255))
    d.text((96, 448), mob.get("hero_line2", "Live Platform")[:24], font=f_hero, fill=acc1)
    d.rounded_rectangle([96, 535, 864, 608], radius=24, fill=(14, 16, 22), outline=(64, 72, 90), width=2)
    d.text((128, 554), mob.get("search_placeholder", "Search...")[:34], font=f_sub, fill=(148, 163, 184))

    chips = mob.get("chips", ["Overview", "Live Feed", "Analytics"])[:4]
    active_chip = int(mob.get("active_chip_index", 1))
    cx = 62
    for idx_c, label in enumerate(chips):
        cw = len(label) * 14 + 54
        if cx + cw > 898:
            break
        active = idx_c == active_chip
        bg = acc1 if active else (22, 25, 34)
        fg = (255, 255, 255) if active else (203, 213, 225)
        bd = acc1 if active else (42, 48, 62)
        d.rounded_rectangle([cx, 678, cx + cw, 742], radius=32, fill=bg, outline=bd, width=2)
        d.text((cx + 24, 694), label, font=f_chip, fill=fg)
        cx += cw + 14

    d.text((64, 782), mob.get("section_label", "VERIFIED ITEMS")[:34], font=f_badge, fill=(148, 163, 184))

    for i, item in enumerate(mob.get("cards", [])[:3]):
        cy = 838 + i * 302
        d.rounded_rectangle([62, cy, 898, cy + 272], radius=32, fill=(20, 22, 29), outline=(44, 50, 64), width=2)
        p_img = load_photo_or_fallback(item.get("image", ""), 208, 224, acc1, gold, bias_y=0.2, seed=60 + i)
        p_rnd = round_image(p_img, 24)
        img.paste(p_rnd, (86, cy + 24), p_rnd)
        d.text((322, cy + 30), item.get("title", "Item")[:22], font=f_card_t, fill=(248, 250, 252))
        d.text((322, cy + 78), item.get("subtitle", "")[:32], font=f_card_s, fill=(148, 163, 184))
        draw_star(d, 336, cy + 134, 13, gold)
        d.text((360, cy + 118), item.get("rating", "5.0 / 5")[:26], font=f_badge, fill=gold)
        d.text((322, cy + 188), item.get("metric", "")[:18], font=f_price, fill=acc1)
        cta_txt = item.get("cta", "Open")[:8]
        d.rounded_rectangle([676, cy + 180, 868, cy + 236], radius=24, fill=acc1)
        cb = d.textbbox((0, 0), cta_txt, font=f_badge)
        d.text((772 - (cb[2] - cb[0]) / 2, cy + 193), cta_txt, font=f_badge, fill=(255, 255, 255))

    d.rounded_rectangle([62, 1775, 898, 1906], radius=44, fill=(16, 18, 24), outline=(44, 50, 64), width=2)
    nav = mob.get("nav_items", ["Home", "Catalog", "Track", "Profile"])[:4]
    for i, item in enumerate(nav):
        nx = 104 + i * 202
        col = acc1 if i == 0 else (148, 163, 184)
        if i == 0:
            d.rounded_rectangle([nx - 14, 1796, nx + 116, 1884], radius=30, fill=(32, 24, 24))
        d.text((nx + 10, 1824), item[:10], font=f_badge, fill=col)

    d.rounded_rectangle([350, 1954, 610, 1966], radius=6, fill=(226, 232, 240))
    img.save(out_path, "PNG")


def build_macbook_screen(resolved: dict, logo_badge: Image.Image, out_path: str) -> None:
    custom = resolved.get("assets", {}).get("desktop_ui_image", "")
    W, H = 1920, 1246
    img = Image.new("RGB", (W, H), (5, 6, 9))
    d = ImageDraw.Draw(img)

    if custom and os.path.exists(custom):
        ui = cover_crop(Image.open(custom).convert("RGB"), W - 36, H - 62)
        ui_r = round_image(ui, 24)
        img.paste(ui_r, (18, 44), ui_r)
        img.save(out_path, "PNG")
        return

    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    acc2 = hex_to_rgb(pal["accent2"])
    gold = hex_to_rgb(pal["accent3"])
    desk = resolved.get("desktop_ui", {})
    brand_name = resolved["brand"]["name"]

    d.rounded_rectangle([18, 44, W - 18, H - 18], radius=24, fill=(9, 10, 13), outline=(34, 38, 50), width=2)
    f_nav_b = load_font("bahnschrift", 30, 700)
    f_nav = load_font("segoe_semi", 22)
    f_h1 = load_font("bahnschrift", 46, 700)
    f_h2 = load_font("bahnschrift", 27, 700)
    f_kpi = load_font("bahnschrift", 44, 700)
    f_body = load_font("segoe", 20)
    f_small = load_font("segoe_semi", 18)

    d.rectangle([18, 44, W - 18, 134], fill=(14, 16, 22))
    d.line([(18, 134), (W - 18, 134)], fill=(34, 38, 50), width=2)
    logo_sm = logo_badge.resize((52, 52), Image.Resampling.LANCZOS)
    img.paste(logo_sm, (58, 63), logo_sm)
    d.text((124, 72), brand_name[:18], font=f_nav_b, fill=(248, 250, 252))

    for i, nav in enumerate(desk.get("nav_links", ["Home", "Modules", "Workflow", "API Docs", "Console"])[:5]):
        d.text((540 + i * 175, 76), nav[:14], font=f_nav, fill=acc1 if i == 3 else (203, 213, 225))
    d.rounded_rectangle([1590, 64, 1858, 114], radius=25, fill=acc1)
    d.text((1642, 74), desk.get("cta_label", "Console")[:16], font=f_nav, fill=(255, 255, 255))

    d.text((64, 168), desk.get("overline", "CONSOLE & TELEMETRY")[:45], font=f_small, fill=acc1)
    d.text((64, 200), desk.get("headline", "Real-time platform operations")[:46], font=f_h1, fill=(255, 255, 255))

    colors = [acc1, acc2, gold, (52, 211, 153)]
    for i, kpi in enumerate(desk.get("kpis", [])[:4]):
        kx = 64 + i * 452
        col = colors[i % len(colors)]
        d.rounded_rectangle([kx, 286, kx + 428, 412], radius=22, fill=(20, 22, 30), outline=(42, 48, 62), width=2)
        d.rounded_rectangle([kx, 286, kx + 10, 412], radius=4, fill=col)
        d.text((kx + 32, 304), str(kpi.get("value", "0"))[:10], font=f_kpi, fill=col)
        d.text((kx + 32, 366), str(kpi.get("label", ""))[:26], font=f_small, fill=(148, 163, 184))

    d.text((64, 444), desk.get("grid_heading", "MODULES & SERVICES")[:38], font=f_small, fill=(148, 163, 184))
    for i, item in enumerate(desk.get("cards", [])[:3]):
        ax = 64 + i * 392
        d.rounded_rectangle([ax, 484, ax + 370, 1180], radius=24, fill=(18, 20, 27), outline=(40, 46, 58), width=2)
        ph = load_photo_or_fallback(item.get("image", ""), 330, 380, colors[i % len(colors)], acc1, bias_y=0.18, seed=80 + i)
        ph_r = round_image(ph, 18)
        img.paste(ph_r, (ax + 20, 504), ph_r)
        d.text((ax + 24, 912), item.get("title", "")[:20], font=f_h2, fill=(255, 255, 255))
        d.text((ax + 24, 956), item.get("subtitle", "")[:28], font=f_body, fill=(148, 163, 184))
        draw_star(d, ax + 36, 1012, 11, gold)
        d.text((ax + 58, 998), item.get("badge", "Verified")[:22], font=f_small, fill=gold)
        d.rounded_rectangle([ax + 24, 1094, ax + 346, 1154], radius=16, fill=(28, 32, 42), outline=acc1, width=2)
        d.text((ax + 56, 1110), item.get("footer", "")[:22], font=f_nav, fill=acc1)

    px = 1256
    d.text((px, 444), desk.get("telemetry_heading", "LIVE EVENT STREAM")[:36], font=f_small, fill=acc2)
    d.rounded_rectangle([px, 484, 1856, 1180], radius=24, fill=(16, 20, 28), outline=(46, 56, 76), width=2)
    for i, ev in enumerate(desk.get("events", [])[:3]):
        ty = 520 + i * 136
        col = colors[(i + 3) % len(colors)]
        d.rounded_rectangle([px + 28, ty, 1828, ty + 112], radius=18, fill=(24, 28, 38), outline=(44, 52, 68), width=1)
        d.ellipse([px + 48, ty + 32, px + 68, ty + 52], fill=col)
        d.text((px + 84, ty + 24), ev.get("title", "")[:28], font=f_h2, fill=(255, 255, 255))
        d.text((px + 84, ty + 66), ev.get("detail", "")[:38], font=f_body, fill=(148, 163, 184))

    api = desk.get("api_box", {})
    d.rounded_rectangle([px + 28, 944, 1828, 1148], radius=18, fill=(10, 12, 18), outline=acc2, width=2)
    d.text((px + 52, 968), api.get("stack", "API & INFRASTRUCTURE")[:32], font=f_small, fill=acc2)
    d.text((px + 52, 1006), api.get("line1", "")[:40], font=f_nav, fill=(248, 250, 252))
    d.text((px + 52, 1044), api.get("line2", "")[:40], font=f_nav, fill=(203, 213, 225))
    d.text((px + 52, 1094), api.get("status", "")[:46], font=f_small, fill=(52, 211, 153))

    img.save(out_path, "PNG")


def build_one_3d_glb(archetype: str, resolved: dict, logo_badge: Image.Image, out_glb: str, out_dir: str) -> str:
    pal = resolved["palette"]
    hero_cfg = resolved["hero_3d"]
    custom_glb = hero_cfg.get("custom_glb_path", "")

    if archetype == "custom" and custom_glb and os.path.exists(custom_glb):
        mesh_name = hero_cfg.get("custom_mesh", "")
        mat_idx = hero_cfg.get("custom_material")
        ui_img = resolved.get("assets", {}).get("desktop_ui_image") or resolved.get("assets", {}).get("mobile_ui_image") or ""
        if mesh_name and mat_idx is not None and ui_img and os.path.exists(ui_img):
            bake_glb_screen(custom_glb, mesh_name, int(mat_idx), ui_img, out_glb, chassis_color=pal["chassis_color"])
        else:
            shutil.copyfile(custom_glb, out_glb)
        return out_glb

    arch_key = canonical_archetype(archetype)
    if arch_key == "smartphone_pro":
        ui_p = os.path.join(out_dir, "screen_iphone_ui.png")
        build_iphone_screen(resolved, logo_badge, ui_p)
        generate_subject_glb(
            archetype="smartphone_pro",
            out_glb=out_glb,
            brand_name=resolved["brand"]["name"],
            accent1_hex=pal["chassis_color"],
            accent2_hex=pal["accent2"],
            accent3_hex=pal["accent3"],
            screen_png_path=ui_p,
        )
        return out_glb

    if arch_key == "laptop_studio":
        ui_p = os.path.join(out_dir, "screen_macbook_ui.png")
        build_macbook_screen(resolved, logo_badge, ui_p)
        generate_subject_glb(
            archetype="laptop_studio",
            out_glb=out_glb,
            brand_name=resolved["brand"]["name"],
            accent1_hex=pal["accent1"],
            accent2_hex=pal["accent2"],
            accent3_hex=pal["accent3"],
            screen_png_path=ui_p,
        )
        return out_glb

    if arch_key not in ARCHETYPES:
        arch_key = "ai_quantum_server"

    generate_subject_glb(
        archetype=arch_key,
        out_glb=out_glb,
        brand_name=resolved["brand"]["name"],
        accent1_hex=pal["accent1"],
        accent2_hex=pal["accent2"],
        accent3_hex=pal["accent3"],
    )
    return out_glb


def generate_all(spec_path: str, out_dir: str) -> None:
    with open(spec_path, "r", encoding="utf-8") as f:
        raw_spec = json.load(f)

    os.makedirs(out_dir, exist_ok=True)
    resolved = normalize_spec_to_v2(raw_spec)
    pal = resolved["palette"]
    acc1 = hex_to_rgb(pal["accent1"])
    acc2 = hex_to_rgb(pal["accent2"])
    acc3 = hex_to_rgb(pal["accent3"])
    obs = hex_to_rgb(pal["obsidian"])
    assets = resolved.get("assets", {})

    logo_badge = prepare_logo_badge(resolved, os.path.join(out_dir, "logo_badge.png"))
    generate_studio_backgrounds_and_glass(resolved, out_dir)

    hero_arch = resolved["hero_3d"]["archetype"]
    hero_glb_path = os.path.join(out_dir, "device_iphone.glb")
    build_one_3d_glb(hero_arch, resolved, logo_badge, hero_glb_path, out_dir)
    shutil.copyfile(hero_glb_path, os.path.join(out_dir, "hero_subject.glb"))

    second_arch = resolved["hero_3d"]["secondary_archetype"]
    macbook_glb_path = ""
    if second_arch:
        macbook_glb_path = os.path.join(out_dir, "device_macbook.glb")
        build_one_3d_glb(second_arch, resolved, logo_badge, macbook_glb_path, out_dir)

    f_badge = load_font("bahnschrift", 22, 700)
    dock_thumbs_done = False

    for idx, sl in enumerate(resolved["slides"], start=1):
        st = sl.get("type", "")
        if st == "cover_hero" and sl.get("style") in ("asymmetric_split", "bold_poster"):
            render_editorial_visual(resolved, sl, idx, out_dir, out_name=f"s{idx}_cover_visual.png")
        elif st == "depth_sandwich_pan":
            render_pan_backdrop_and_cutout(resolved, sl, idx, out_dir)
        elif st == "chart_infographic":
            render_chart_infographic(resolved, sl, idx, out_dir)
        elif st == "architecture_diagram":
            render_architecture_diagram(resolved, sl, idx, out_dir)
        elif st == "editorial_split":
            render_editorial_visual(resolved, sl, idx, out_dir)
        elif st == "full_bleed_visual":
            img_src = sl.get("image", "")
            load_photo_or_fallback(img_src, 1920, 1080, acc1, acc2, bias_y=0.45, seed=900 + idx).save(
                os.path.join(out_dir, f"s{idx}_visual.png"), "PNG"
            )
        elif st == "spotlight_dock":
            bg_src = sl.get("background_image", "")
            if not bg_src:
                fb_list = assets.get("focus_backgrounds", [])
                sp_i = int(sl.get("spotlight_index", 0))
                if sp_i < len(fb_list):
                    bg_src = fb_list[sp_i]
            foc = add_bottom_gradient(
                add_left_gradient(
                    ImageEnhance.Brightness(load_photo_or_fallback(bg_src, 1920, 1080, acc1 if idx % 2 == 1 else acc2, obs, 0.35, seed=400 + idx)).enhance(0.74),
                    base_col=obs, max_alpha=242, fade_frac=0.66,
                ),
                base_col=obs, max_alpha=235, start_frac=0.58,
            ).convert("RGB")
            foc.save(os.path.join(out_dir, f"s{idx}_bg.png"), "PNG")

            sp = sl.get("spotlight", {})
            cw, ch = 520, 600
            canvas = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
            sh = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
            ImageDraw.Draw(sh).rounded_rectangle([24, 24, cw - 24, ch - 18], radius=38, fill=(0, 0, 0, 160))
            canvas = Image.alpha_composite(canvas, sh.filter(ImageFilter.GaussianBlur(14)))
            p_img = add_bottom_gradient(
                load_photo_or_fallback(sp.get("portrait_image", ""), cw - 48, ch - 48, acc1 if idx % 2 == 1 else acc2, obs, 0.18, seed=450 + idx),
                base_col=obs, max_alpha=215, start_frac=0.65,
            )
            p_r = round_image(p_img, 34)
            canvas.paste(p_r, (24, 20), p_r)
            cd = ImageDraw.Draw(canvas)
            cd.rounded_rectangle([24, 20, cw - 24, ch - 28], radius=34, outline=(255, 255, 255, 88), width=2)
            cd.rounded_rectangle([54, ch - 104, cw - 54, ch - 52], radius=24, fill=(*obs, 232), outline=(*acc3, 215), width=2)
            draw_star(cd, 96, ch - 78, 13, acc3)
            cd.text((124, ch - 90), str(sp.get("badge", "VERIFIED"))[:22], font=f_badge, fill=acc3)
            canvas.save(os.path.join(out_dir, f"s{idx}_portrait.png"), "PNG")

            if not dock_thumbs_done:
                dock_items = sl.get("dock") or resolved.get("dock", [])
                for di_i in range(4):
                    d_item = dock_items[di_i] if di_i < len(dock_items) else {}
                    th = load_photo_or_fallback(d_item.get("image", ""), 180, 140, acc1 if di_i % 2 == 0 else acc2, obs, 0.18, seed=480 + di_i)
                    th_r = round_image(th, 18)
                    ImageDraw.Draw(th_r).rounded_rectangle([0, 0, 179, 139], radius=18, outline=(255, 255, 255, 78), width=2)
                    th_r.save(os.path.join(out_dir, f"thumb_{di_i + 1}.png"), "PNG")
                dock_thumbs_done = True

    resolved_out = {k: v for k, v in resolved.items() if k != "theme_cfg"}
    resolved_out["requirements"]["iphone_glb"] = hero_glb_path
    resolved_out["requirements"]["macbook_glb"] = macbook_glb_path

    with open(os.path.join(out_dir, "theme_resolved.json"), "w", encoding="utf-8") as f:
        json.dump(resolved_out, f, indent=2, ensure_ascii=False)

    print(
        f"Studio assets + 3D GLB ({hero_arch}"
        + (f" + {second_arch}" if second_arch else "")
        + f") generated cleanly ({resolved['theme']}, {len(resolved['slides'])} slides) -> {out_dir}"
    )


def main():
    parser = argparse.ArgumentParser(description="Generate studio graphics, frosted glass & subject-linked 3D GLBs from deck_spec.json")
    parser.add_argument("--spec", required=True, help="Path to deck_spec.json")
    parser.add_argument("--out", required=True, help="Output assets working directory")
    args = parser.parse_args()
    generate_all(args.spec, args.out)


if __name__ == "__main__":
    main()
