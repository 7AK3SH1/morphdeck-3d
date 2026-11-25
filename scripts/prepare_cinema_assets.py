#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Cinema-Morph Asset Orchestrator for morphdeck-3d v3.0.0.

Resolves real high-poly CC0 .glb models (NEVER procedural toy meshes),
fetches real open-license landscape photography via Wikimedia Commons when needed,
and prepares studio cutouts, architectural portal masks, sky-split horizons,
archival broadsheets, organic ink drips, cloud sweeps, hollow museum frames,
frosted refracted UI containers, and 2x supersampled charts/architecture diagrams.
"""

import argparse
import copy
import json
import os
import shutil
import sys
import textwrap
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from fetch_real_glb import fetch_khronos_glb, fetch_nasa_glb, fetch_polyhaven_glb, resolve_auto_glb  # noqa: E402
from fetch_real_photo import fetch_url_photo, fetch_wikimedia_photo  # noqa: E402
from render_glb_preview import render_glb_to_image  # noqa: E402
from studio_cutout import (  # noqa: E402
    cover_crop,
    generate_newspaper_sheet,
    hex_to_rgb,
    load_font,
    make_contact_shadow,
    make_studio_cutout,
    punch_portal_mask,
    render_cloud_sweep,
    render_hollow_gold_frame,
    render_organic_ink_drips,
    render_refracted_glass_dock,
    render_silhouette_horizon,
    render_soft_paper_cloud,
    split_horizon_cutout,
)

UNIVERSES = {
    "documentary_press": {
        "is_light": True,
        "bg": "#EAE3D2",
        "surface": "#F6F1E6",
        "surface_alt": "#DFD6C3",
        "ink": "#161618",
        "ink_muted": "#544E46",
        "border": "#B8AE9E",
        "accents": ["#B91C1C", "#1F2937", "#B45309"],
        "display_font": "Georgia",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
    },
    "architectural_portal": {
        "is_light": True,
        "bg": "#F4EFE6",
        "surface": "#FAF7F2",
        "surface_alt": "#161412",
        "ink": "#141413",
        "ink_muted": "#57534E",
        "border": "#D6CFC4",
        "accents": ["#C2410C", "#0D9488", "#D97706"],
        "display_font": "Georgia",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
    },
    "sculptural_museum": {
        "is_light": False,
        "bg": "#0A0C12",
        "surface": "#141824",
        "surface_alt": "#1E2436",
        "ink": "#F8FAFC",
        "ink_muted": "#94A3B8",
        "border": "#334155",
        "accents": ["#E07A5F", "#38BDF8", "#D4AF37"],
        "display_font": "Georgia",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
    },
    "commercial_studio": {
        "is_light": False,
        "bg": "#080A0F",
        "surface": "#131722",
        "surface_alt": "#1B2130",
        "ink": "#FFFFFF",
        "ink_muted": "#94A3B8",
        "border": "#2D3748",
        "accents": ["#E11D48", "#38BDF8", "#F59E0B"],
        "display_font": "Bahnschrift",
        "heading_font": "Impact",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
    },
    "executive_swiss": {
        "is_light": True,
        "bg": "#F8FAFC",
        "surface": "#FFFFFF",
        "surface_alt": "#EFF4FA",
        "ink": "#0F172A",
        "ink_muted": "#475569",
        "border": "#CBD5E1",
        "accents": ["#2563EB", "#0F766E", "#D97706"],
        "display_font": "Bahnschrift",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
    },
    "editorial_pop": {
        "is_light": False,
        "bg": "#0F0C18",
        "surface": "#1A1529",
        "surface_alt": "#26203D",
        "ink": "#FAFAFA",
        "ink_muted": "#C4B5FD",
        "border": "#4C1D95",
        "accents": ["#EC4899", "#06B6D4", "#FBBF24"],
        "display_font": "Impact",
        "heading_font": "Bahnschrift",
        "body_font": "Segoe UI",
        "semi_font": "Segoe UI Semibold",
        "mono_font": "Consolas",
    },
}


def draw_glass_or_swiss_card(
    w: int, h: int, surf_rgb: tuple, acc_rgb: tuple, border_rgb: tuple, is_light: bool, radius: int = 24, bar_mode: str = "left"
) -> Image.Image:
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh_alpha = 46 if is_light else 148
    ImageDraw.Draw(sh).rounded_rectangle([10, 12, w - 10, h - 6], radius=radius + 2, fill=(0, 0, 0, sh_alpha))
    canvas = Image.alpha_composite(canvas, sh.filter(ImageFilter.GaussianBlur(8)))

    d = ImageDraw.Draw(canvas)
    fill_a = 244 if is_light else 218
    out_col = (*border_rgb, 210) if is_light else (255, 255, 255, 62)
    d.rounded_rectangle([6, 6, w - 6, h - 10], radius=radius, fill=(*surf_rgb, fill_a), outline=out_col, width=2)
    if not is_light:
        d.line([(6 + radius, 8), (w - 6 - radius, 8)], fill=(255, 255, 255, 138), width=2)

    if bar_mode == "left":
        d.rounded_rectangle([6, 22, 14, h - 26], radius=4, fill=(*acc_rgb, 245))
    elif bar_mode == "top":
        d.rounded_rectangle([6 + radius, 6, w - 6 - radius, 13], radius=3, fill=(*acc_rgb, 245))
    return canvas


def _build_textured_canvas(w: int, h: int, bg_rgb: tuple, seed: int = 7) -> Image.Image:
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    vig = ((((xx - w * 0.5) / (w * 0.62)) ** 2 + ((yy - h * 0.5) / (h * 0.62)) ** 2) * 16).clip(0, 26).astype(np.int16)[..., None]
    noise = rng.integers(-5, 6, size=(h, w, 1), dtype=np.int16)
    arr = np.clip(np.full((h, w, 3), bg_rgb, dtype=np.int16) + noise - vig, 0, 255).astype(np.uint8)
    return Image.fromarray(arr, "RGB")


def _render_chart_png(resolved: dict, slide: dict, out_path: str) -> None:
    pal = resolved["palette"]
    is_light = resolved["is_light"]
    surf = hex_to_rgb(pal["surface"])
    surf_alt = hex_to_rgb(pal["surface_alt"])
    border = hex_to_rgb(pal["border"])
    ink = hex_to_rgb(pal["ink"])
    muted = hex_to_rgb(pal["ink_muted"])
    accents = [hex_to_rgb(pal["accent1"]), hex_to_rgb(pal["accent2"]), hex_to_rgb(pal["accent3"])]

    W, H = 2360, 1520
    img = draw_glass_or_swiss_card(W, H, surf, accents[0], border, is_light, radius=44, bar_mode="none")
    d = ImageDraw.Draw(img)
    f_num = load_font("bahnschrift", 90, 700)
    f_lbl = load_font("bahnschrift", 40, 700)
    f_sub = load_font("segoe", 32)
    f_mono = load_font("consolas_bold", 32)

    items = slide.get("items", [])[:4]
    chart_type = slide.get("chart_type", "comparison_bars")

    if chart_type == "radial_rings":
        n = max(1, len(items))
        slot_w = (W - 160) // n
        cy, ring_r, ring_w = 620, (215 if n <= 3 else 175), (42 if n <= 3 else 36)
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
            lbl = str(it.get("label", ""))[:24]
            lbb = d.textbbox((0, 0), lbl, font=f_lbl)
            d.text((cx - (lbb[2] - lbb[0]) / 2, cy + ring_r + 64), lbl, font=f_lbl, fill=col)
            sub = str(it.get("sublabel", ""))[:28]
            sbb = d.textbbox((0, 0), sub, font=f_sub)
            d.text((cx - (sbb[2] - sbb[0]) / 2, cy + ring_r + 128), sub, font=f_sub, fill=muted)
    else:
        n = max(1, len(items))
        top_y, row_h, bar_x0, bar_max_w = 96, (H - 180) // n, 100, W - 540
        max_val = max([max(float(it.get("value", 1)), float(it.get("baseline", 1))) for it in items] + [1.0])
        for j, it in enumerate(items):
            ry = top_y + j * row_h + 24
            col = accents[j % len(accents)]
            val = float(it.get("value", 0))
            base = float(it.get("baseline", 0))
            unit = str(it.get("unit", ""))
            d.text((bar_x0, ry), str(it.get("label", ""))[:36], font=f_lbl, fill=ink)
            bw = max(40, int((base / max_val) * bar_max_w))
            d.rounded_rectangle([bar_x0, ry + 76, bar_x0 + bw, ry + 116], radius=14, fill=surf_alt)
            d.text((bar_x0 + bw + 24, ry + 76), f"{it.get('baseline', '')} {unit}".strip(), font=f_sub, fill=muted)
            vw = max(60, int((val / max_val) * bar_max_w))
            d.rounded_rectangle([bar_x0, ry + 132, bar_x0 + vw, ry + 188], radius=16, fill=col)
            d.text((bar_x0 + vw + 24, ry + 138), f"{it.get('value', '')} {unit}".strip(), font=f_lbl, fill=ink)
            delta = str(it.get("delta", ""))
            if delta:
                dbb = d.textbbox((0, 0), delta, font=f_mono)
                pw = (dbb[2] - dbb[0]) + 44
                d.rounded_rectangle([W - 100 - pw, ry - 4, W - 100, ry + 50], radius=14, fill=(*col, 42), outline=col, width=2)
                d.text((W - 100 - pw + 22, ry + 6), delta, font=f_mono, fill=col)

    img.resize((1180, 760), Image.Resampling.LANCZOS).save(out_path, "PNG")


def _render_arch_png(resolved: dict, slide: dict, out_path: str) -> None:
    pal = resolved["palette"]
    is_light = resolved["is_light"]
    surf = hex_to_rgb(pal["surface"])
    surf_alt = hex_to_rgb(pal["surface_alt"])
    border = hex_to_rgb(pal["border"])
    ink = hex_to_rgb(pal["ink"])
    muted = hex_to_rgb(pal["ink_muted"])
    accents = [hex_to_rgb(pal["accent1"]), hex_to_rgb(pal["accent2"]), hex_to_rgb(pal["accent3"])]

    W, H = 2560, 1592
    img = draw_glass_or_swiss_card(W, H, surf, accents[0], border, is_light, radius=44, bar_mode="none")
    d = ImageDraw.Draw(img)
    f_layer = load_font("consolas_bold", 28)
    f_node_t = load_font("bahnschrift", 38, 700)
    f_node_s = load_font("segoe", 28)
    f_tag = load_font("consolas_bold", 24)

    layers = slide.get("layers", [])[:4]
    flows = slide.get("flows", [])
    n_layers = max(1, len(layers))
    margin_x, top_pad = 76, 64
    gap_y = 84 if n_layers <= 3 else 56
    layer_h = (H - top_pad * 2 - gap_y * (n_layers - 1)) // n_layers

    for li, layer in enumerate(layers):
        ly = top_pad + li * (layer_h + gap_y)
        col = accents[li % len(accents)]
        d.rounded_rectangle([margin_x, ly, W - margin_x, ly + layer_h], radius=24, fill=(*surf_alt, 220), outline=(*col, 180), width=3)
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
            d.text((nx + 32, inner_top + 22), str(nd.get("title", ""))[:26], font=f_node_t, fill=ink)
            for wi, line in enumerate(textwrap.wrap(str(nd.get("subtitle", "")), width=max(18, int((node_w - 60) / 17)))[:2]):
                d.text((nx + 32, inner_top + 78 + wi * 38), line, font=f_node_s, fill=muted)

        if li < n_layers - 1:
            ay1, ay2 = ly + layer_h + 6, ly + layer_h + gap_y - 6
            mid_y = (ay1 + ay2) // 2
            for ax in (W // 4, W // 2, (W * 3) // 4):
                d.line([(ax, ay1), (ax, ay2 - 12)], fill=col, width=5)
                d.polygon([(ax, ay2), (ax - 12, ay2 - 18), (ax + 12, ay2 - 18)], fill=col)
            if li < len(flows) and flows[li]:
                fl = flows[li]
                fbb = d.textbbox((0, 0), fl, font=f_tag)
                fw = (fbb[2] - fbb[0]) + 44
                d.rounded_rectangle([W // 2 - fw // 2, mid_y - 20, W // 2 + fw // 2, mid_y + 20], radius=16, fill=(*surf, 255), outline=col, width=3)
                d.text((W // 2 - (fbb[2] - fbb[0]) // 2, mid_y - 13), fl, font=f_tag, fill=col)

    img.resize((1280, 796), Image.Resampling.LANCZOS).save(out_path, "PNG")


def normalize_spec(raw_spec: dict) -> dict:
    spec = copy.deepcopy(raw_spec)
    universe = spec.get("universe") or spec.get("theme") or "commercial_studio"
    if universe not in UNIVERSES:
        universe = "commercial_studio"
    ucfg = UNIVERSES[universe]

    pal_in = spec.get("palette", {})
    d_acc = ucfg["accents"]
    palette = {
        "bg": pal_in.get("bg") or ucfg["bg"],
        "surface": pal_in.get("surface") or ucfg["surface"],
        "surface_alt": pal_in.get("surface_alt") or ucfg["surface_alt"],
        "ink": pal_in.get("ink") or ucfg["ink"],
        "ink_muted": pal_in.get("ink_muted") or ucfg["ink_muted"],
        "border": pal_in.get("border") or ucfg["border"],
        "accent1": pal_in.get("accent1") or d_acc[0],
        "accent2": pal_in.get("accent2") or d_acc[1],
        "accent3": pal_in.get("accent3") or d_acc[2],
    }
    is_light = bool(spec.get("is_light", pal_in.get("is_light", ucfg["is_light"])))

    return {
        "schema_version": 3,
        "universe": universe,
        "is_light": is_light,
        "brand": spec.get("brand", {"name": "Presentation", "monogram": "PR"}),
        "palette": palette,
        "fonts": {
            "display": spec.get("fonts", {}).get("display") or ucfg["display_font"],
            "heading": spec.get("fonts", {}).get("heading") or ucfg["heading_font"],
            "body": spec.get("fonts", {}).get("body") or ucfg["body_font"],
            "semi": spec.get("fonts", {}).get("semi") or ucfg["semi_font"],
            "mono": spec.get("fonts", {}).get("mono") or ucfg["mono_font"],
        },
        "hero_3d": spec.get("hero_3d", {}),
        "assets": spec.get("assets", {}),
        "slides": spec.get("slides", []),
    }


def _resolve_glb_entry(cfg: dict, out_glb: str, rel_fn) -> str:
    if not cfg:
        return ""
    local = rel_fn(cfg.get("glb_path", ""))
    if local and os.path.exists(local):
        shutil.copyfile(local, out_glb)
        return out_glb
    if cfg.get("source") and cfg.get("name"):
        src, name = cfg["source"].lower(), cfg["name"]
        if src == "khronos":
            return fetch_khronos_glb(name, out_glb)
        if src == "polyhaven":
            return fetch_polyhaven_glb(name, out_glb, res=cfg.get("res", "1k"))
        if src == "nasa":
            return fetch_nasa_glb(name, out_glb)
    if cfg.get("auto_query"):
        resolve_auto_glb(cfg["auto_query"], out_glb)
        return out_glb
    return ""


def _resolve_cutout_if_needed(src_path: str, out_path: str, already_cutout: bool = False, tol: float = 28.0) -> str:
    if not src_path or not os.path.exists(src_path):
        return ""
    if already_cutout:
        shutil.copyfile(src_path, out_path)
        return out_path
    with Image.open(src_path) as im:
        if im.mode == "RGBA":
            a = np.asarray(im)[..., 3]
            if float(np.mean(a < 16)) > 0.08:
                im.save(out_path, "PNG")
                return out_path
    return make_studio_cutout(src_path, out_path, tolerance=tol)


def _resolve_photo_field(sl: dict, path_key: str, query_key: str, url_key: str, out_dir: str, tag: str, rel_fn) -> str:
    local = rel_fn(sl.get(path_key, ""))
    if local and os.path.exists(local):
        return local
    q = sl.get(query_key, "")
    if q:
        dst = os.path.join(out_dir, f"_wm_{tag}.png")
        try:
            return fetch_wikimedia_photo(q, dst)
        except Exception as exc:
            print(f"Warning: Wikimedia query '{q}' skipped ({exc})")
    u = sl.get(url_key, "")
    if u:
        dst = os.path.join(out_dir, f"_url_{tag}.png")
        try:
            return fetch_url_photo(u, dst)
        except Exception as exc:
            print(f"Warning: URL photo '{u}' skipped ({exc})")
    return ""


def prepare_all(spec_path: str, out_dir: str) -> dict:
    with open(spec_path, "r", encoding="utf-8") as f:
        raw_spec = json.load(f)

    os.makedirs(out_dir, exist_ok=True)
    spec_dir = os.path.dirname(os.path.abspath(spec_path))
    resolved = normalize_spec(raw_spec)
    pal = resolved["palette"]
    is_light = resolved["is_light"]
    universe = resolved["universe"]

    bg_rgb = hex_to_rgb(pal["bg"])
    surf_rgb = hex_to_rgb(pal["surface"])
    acc1 = hex_to_rgb(pal["accent1"])
    acc2 = hex_to_rgb(pal["accent2"])
    acc3 = hex_to_rgb(pal["accent3"])
    border_rgb = hex_to_rgb(pal["border"])

    def rel(p: str) -> str:
        if not p:
            return ""
        if os.path.isabs(p):
            return p
        c = os.path.join(spec_dir, p)
        return c if os.path.exists(c) else p

    _build_textured_canvas(1920, 1080, bg_rgb, seed=11).save(os.path.join(out_dir, "bg_base.png"), "PNG")
    _build_textured_canvas(1920, 1080, hex_to_rgb(pal["surface_alt"]), seed=22).save(os.path.join(out_dir, "bg_alt.png"), "PNG")
    make_contact_shadow(os.path.join(out_dir, "shadow_hero.png"), 640, 140, 175)

    draw_glass_or_swiss_card(1060, 148, surf_rgb, acc1, border_rgb, is_light, 22, "left").save(os.path.join(out_dir, "card_step.png"), "PNG")
    for ci, ac in enumerate([acc1, acc2, acc3], start=1):
        draw_glass_or_swiss_card(730, 192, surf_rgb, ac, border_rgb, is_light, 22, "left").save(os.path.join(out_dir, f"card_spec_{ci}.png"), "PNG")
        draw_glass_or_swiss_card(552, 420, surf_rgb, ac, border_rgb, is_light, 24, "top").save(os.path.join(out_dir, f"card_bento_tile_{ci}.png"), "PNG")
    draw_glass_or_swiss_card(672, 768, surf_rgb, acc1, border_rgb, is_light, 26, "top").save(os.path.join(out_dir, "card_bento_hero.png"), "PNG")
    draw_glass_or_swiss_card(1732, 296, surf_rgb, acc1, border_rgb, is_light, 26, "left").save(os.path.join(out_dir, "card_bento_strip.png"), "PNG")
    draw_glass_or_swiss_card(1024, 520, surf_rgb, acc1, border_rgb, is_light, 26, "left").save(os.path.join(out_dir, "card_panel.png"), "PNG")
    draw_glass_or_swiss_card(1760, 256, surf_rgb, acc1, border_rgb, is_light, 34, "none").save(os.path.join(out_dir, "dock_bg.png"), "PNG")

    gw, gh = 416, 216
    glider = Image.new("RGBA", (gw, gh), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glider)
    gd.rounded_rectangle([4, 4, gw - 4, gh - 4], radius=26, fill=(*acc1, 52), outline=(*acc1, 235), width=3)
    gd.line([(34, 6), (gw - 34, 6)], fill=(255, 255, 255, 175), width=2)
    glider.save(os.path.join(out_dir, "dock_glider.png"), "PNG")

    # Reusable documentary & cinematic FX overlays
    render_organic_ink_drips(os.path.join(out_dir, "ep_drips.png"), color_hex=pal["accent1"])
    render_soft_paper_cloud(os.path.join(out_dir, "ep_paper_cloud.png"), paper_hex=pal["surface"])
    render_cloud_sweep(os.path.join(out_dir, "flyby_cloud.png"))
    render_silhouette_horizon(os.path.join(out_dir, "odo_silhouette.png"), color_hex=pal["accent1"])

    resolved_glb_out = _resolve_glb_entry(resolved.get("hero_3d", {}), os.path.join(out_dir, "hero_real.glb"), rel)
    if resolved_glb_out and os.path.exists(resolved_glb_out):
        render_glb_to_image(resolved_glb_out, 720, 720, 12.0, -22.0, 2.0, theme=universe).save(
            os.path.join(out_dir, "hero_cutout_auto.png"), "PNG"
        )

    news_done = False
    dock_thumbs_done = False
    portal_done = False

    for idx, sl in enumerate(resolved["slides"], start=1):
        st = sl.get("type", "")

        # Resolve slide-level Wikimedia / URL photography fields cleanly
        for p_key, q_key, u_key in (
            ("bg_image", "bg_query", "bg_url"),
            ("map_image", "map_query", "map_url"),
            ("wall_image", "wall_query", "wall_url"),
            ("framed_art_image", "framed_art_query", "framed_art_url"),
            ("poster_bg_image", "poster_bg_query", "poster_bg_url"),
            ("lower_image", "lower_query", "lower_url"),
        ):
            res_p = _resolve_photo_field(sl, p_key, q_key, u_key, out_dir, f"s{idx}_{p_key}", rel)
            if res_p and os.path.exists(res_p):
                dst_crop = os.path.join(out_dir, f"s{idx}_{p_key}.png")
                cover_crop(Image.open(res_p).convert("RGB"), 1920, 1080).save(dst_crop, "PNG")
                sl[f"_{p_key}"] = dst_crop

        if st in ("newspaper_dive", "editorial_print_overlay") and not news_done:
            custom_ns = rel(sl.get("newspaper_image", ""))
            ns_out = os.path.join(out_dir, "news_sheet.png")
            if custom_ns and os.path.exists(custom_ns):
                cover_crop(Image.open(custom_ns).convert("RGB"), 1920, 1080).save(ns_out, "PNG")
            else:
                photos = [rel(p) for p in sl.get("archival_photos", [])]
                for qi, aq in enumerate(sl.get("archival_queries", [])[:2]):
                    try:
                        photos.append(fetch_wikimedia_photo(aq, os.path.join(out_dir, f"_wm_news_{qi}.png")))
                    except Exception:
                        pass
                generate_newspaper_sheet(
                    headline=sl.get("newspaper_headline") or sl.get("headline") or resolved["brand"].get("name", "SPECIAL REPORT"),
                    subhead=sl.get("newspaper_subhead") or sl.get("subhead") or "COMPREHENSIVE BRIEFING & ARCHIVAL DISPATCH",
                    dateline=sl.get("dateline") or "ARCHIVE EDITION  //  VOL. IV",
                    columns=sl.get("columns", []),
                    photo_paths=photos,
                    out_path=ns_out,
                    paper_hex=pal["bg"] if is_light else "#EDE6D6",
                    accent_hex=pal["accent1"],
                )
            news_done = True

        # Cutout helpers on slides: supports studio photo OR real GLB-rendered cutout (`*_glb` spec)!
        for field_key, out_suffix in (
            ("left_cutout", f"s{idx}_left_cutout.png"),
            ("right_cutout", f"s{idx}_right_cutout.png"),
            ("prop_cutout", f"s{idx}_prop_cutout.png"),
            ("portrait_cutout", f"s{idx}_portrait_cutout.png"),
            ("flyby_cutout", f"s{idx}_flyby_cutout.png"),
            ("hero_cutout", f"s{idx}_hero_cutout.png"),
            ("fg_left_cutout", f"s{idx}_fg_left.png"),
            ("fg_right_cutout", f"s{idx}_fg_right.png"),
            ("silhouette_cutout", f"s{idx}_silhouette.png"),
            ("room_subject_cutout", f"s{idx}_room_subject.png"),
        ):
            src_p = rel(sl.get(field_key, ""))
            dst_p = os.path.join(out_dir, out_suffix)
            if src_p and os.path.exists(src_p):
                _resolve_cutout_if_needed(src_p, dst_p, already_cutout=bool(sl.get("already_cutout", False)))
                sl[f"_{field_key}"] = dst_p
            elif isinstance(sl.get(f"{field_key}_glb"), dict):
                gcfg = sl[f"{field_key}_glb"]
                gpath = _resolve_glb_entry(gcfg, os.path.join(out_dir, f"_tmp_{field_key}_{idx}.glb"), rel)
                if gpath and os.path.exists(gpath):
                    render_glb_to_image(
                        gpath, 720, 720,
                        float(gcfg.get("rx", 12.0)), float(gcfg.get("ry", -24.0)), float(gcfg.get("rz", 2.0)),
                        theme=universe,
                    ).save(dst_p, "PNG")
                    sl[f"_{field_key}"] = dst_p

        if st == "real_portal_flythrough" and not portal_done:
            out_src = _resolve_photo_field(sl, "portal_outer_image", "portal_outer_query", "portal_outer_url", out_dir, f"s{idx}_pout", rel)
            in_src = _resolve_photo_field(sl, "portal_inner_image", "portal_inner_query", "portal_inner_url", out_dir, f"s{idx}_pin", rel)
            bg_src = _resolve_photo_field(sl, "portal_bg_image", "portal_bg_query", "portal_bg_url", out_dir, f"s{idx}_pbg", rel) or sl.get("_bg_image", "")

            if out_src and os.path.exists(out_src):
                box = tuple(sl.get("outer_box", [0.27, 0.18, 0.73, 0.82]))
                punch_portal_mask(out_src, os.path.join(out_dir, "portal_outer.png"), box_frac=box, radius=18)
            else:
                punch_portal_mask(os.path.join(out_dir, "bg_alt.png"), os.path.join(out_dir, "portal_outer.png"), box_frac=(0.27, 0.18, 0.73, 0.82), radius=24)

            if in_src and os.path.exists(in_src):
                box_in = tuple(sl.get("inner_box", [0.18, 0.12, 0.82, 0.88]))
                punch_portal_mask(in_src, os.path.join(out_dir, "portal_inner.png"), box_frac=box_in, radius=16)
            else:
                shutil.copyfile(os.path.join(out_dir, "portal_outer.png"), os.path.join(out_dir, "portal_inner.png"))

            if bg_src and os.path.exists(bg_src):
                cover_crop(Image.open(bg_src).convert("RGB"), 1920, 1080).save(os.path.join(out_dir, "portal_bg.png"), "PNG")
            else:
                shutil.copyfile(os.path.join(out_dir, "bg_base.png"), os.path.join(out_dir, "portal_bg.png"))
            portal_done = True

        if st == "vertical_world_elevator":
            bg_src = _resolve_photo_field(sl, "panorama_image", "panorama_query", "panorama_url", out_dir, f"s{idx}_pano", rel) or sl.get("_bg_image", "")
            if bg_src and os.path.exists(bg_src):
                cover_crop(Image.open(bg_src).convert("RGB"), 1920, 2160).save(os.path.join(out_dir, "elev_bg.png"), "PNG")
                split_horizon_cutout(os.path.join(out_dir, "elev_bg.png"), os.path.join(out_dir, "elev_fg.png"), size=(1920, 2160))

        if st == "parchment_3d_split":
            low_src = sl.get("_lower_image") or sl.get("_bg_image") or ""
            if low_src and os.path.exists(low_src):
                cover_crop(Image.open(low_src).convert("RGB"), 1920, 560).save(os.path.join(out_dir, f"s{idx}_split_lower.png"), "PNG")

        if st == "refracted_spotlight_dock":
            bg_src = sl.get("_bg_image", "")
            bg_dst = os.path.join(out_dir, f"s{idx}_bg.png")
            if bg_src and os.path.exists(bg_src):
                cover_crop(Image.open(bg_src).convert("RGB"), 1920, 1080).save(bg_dst, "PNG")
            else:
                shutil.copyfile(os.path.join(out_dir, "bg_base.png"), bg_dst)
            render_refracted_glass_dock(bg_dst, os.path.join(out_dir, f"s{idx}_dock_refract.png"), tint_rgba=(*surf_rgb, 146))
            if not dock_thumbs_done:
                for di, it in enumerate(sl.get("dock", [])[:4], start=1):
                    tsrc = _resolve_photo_field(it, "image", "query", "url", out_dir, f"thumb_{di}", rel)
                    tdst = os.path.join(out_dir, f"thumb_{di}.png")
                    if tsrc and os.path.exists(tsrc):
                        cover_crop(Image.open(tsrc).convert("RGB"), 180, 140).save(tdst, "PNG")
                    else:
                        _build_textured_canvas(180, 140, acc1 if di % 2 else acc2, seed=100 + di).save(tdst, "PNG")
                dock_thumbs_done = True

        if st == "picture_frame_pullback":
            hframe_dst = os.path.join(out_dir, f"s{idx}_hollow_frame.png")
            render_hollow_gold_frame(hframe_dst, frame_style=sl.get("frame_style", "ornate_gold"), border_px=84)
            sl["_hollow_frame"] = hframe_dst
            if sl.get("_framed_art_image"):
                sl["_framed_art"] = sl["_framed_art_image"]

        if st == "chart_infographic":
            _render_chart_png(resolved, sl, os.path.join(out_dir, f"s{idx}_chart.png"))
        if st == "architecture_blueprint":
            _render_arch_png(resolved, sl, os.path.join(out_dir, f"s{idx}_arch.png"))

    resolved["resolved_glb"] = resolved_glb_out
    with open(os.path.join(out_dir, "deck_resolved.json"), "w", encoding="utf-8") as f:
        json.dump(resolved, f, indent=2, ensure_ascii=False)
    print(f"Prepared cinema assets ({universe}, {len(resolved['slides'])} slides, glb={bool(resolved_glb_out)}) -> {out_dir}")
    return resolved


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare real 3D .glb + photorealistic studio assets for morphdeck-3d v3.")
    parser.add_argument("--spec", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    prepare_all(args.spec, args.out)


if __name__ == "__main__":
    main()
