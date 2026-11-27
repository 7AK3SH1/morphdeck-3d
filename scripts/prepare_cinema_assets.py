#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Cinema-Morph v4.0.0 Asset Orchestrator for morphdeck-3d.

Resolves real high-poly CC0 .glb models, juried Wikimedia Featured/Quality photography,
infinity-pool architectural portal pavilions, cloud-proof 1D horizon sky splits,
tourist-free museum spotlight walls, mitered gilt frames, optical ND text scrims,
backdrop-refracted glass docks, and 2x supersampled editorial charts/blueprints.
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
    apply_tonal_text_scrim,
    cover_crop,
    generate_newspaper_sheet,
    hex_to_rgb,
    load_font,
    make_contact_shadow,
    make_studio_cutout,
    punch_portal_mask,
    render_architectural_portal_exterior,
    render_architectural_portal_interior,
    render_cloud_sweep,
    render_gallery_wall_room,
    render_hollow_gold_frame,
    render_odometer_masks,
    render_organic_ink_drips,
    render_refracted_card,
    render_refracted_glass_dock,
    render_silhouette_horizon,
    render_soft_dock_glider,
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
        "border": "#C5BBA8",
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
        "surface_alt": "#181A20",
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
        "surface_alt": "#1A1E2B",
        "ink": "#F8FAFC",
        "ink_muted": "#94A3B8",
        "border": "#2E384D",
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
        "surface": "#121620",
        "surface_alt": "#191F2E",
        "ink": "#FFFFFF",
        "ink_muted": "#94A3B8",
        "border": "#283244",
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
    w: int, h: int, surf_rgb: tuple, acc_rgb: tuple, border_rgb: tuple, is_light: bool, radius: int = 20, bar_mode: str = "left"
) -> Image.Image:
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sh_alpha = 34 if is_light else 120
    ImageDraw.Draw(sh).rounded_rectangle([10, 12, w - 10, h - 6], radius=radius + 2, fill=(0, 0, 0, sh_alpha))
    canvas = Image.alpha_composite(canvas, sh.filter(ImageFilter.GaussianBlur(8)))

    d = ImageDraw.Draw(canvas)
    fill_a = 246 if is_light else 214
    out_col = (*border_rgb, 185) if is_light else (255, 255, 255, 44)
    d.rounded_rectangle([6, 6, w - 6, h - 10], radius=radius, fill=(*surf_rgb, fill_a), outline=out_col, width=1)
    if not is_light:
        d.line([(6 + radius, 7), (w - 6 - radius, 7)], fill=(255, 255, 255, 110), width=1)

    if bar_mode == "left":
        d.rounded_rectangle([6, 22, 12, h - 26], radius=3, fill=(*acc_rgb, 240))
    elif bar_mode == "top":
        d.rounded_rectangle([6 + radius, 6, w - 6 - radius, 11], radius=2, fill=(*acc_rgb, 240))
    return canvas


def _build_textured_canvas(w: int, h: int, bg_rgb: tuple, seed: int = 7) -> Image.Image:
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    vig = ((((xx - w * 0.5) / (w * 0.65)) ** 2 + ((yy - h * 0.5) / (h * 0.65)) ** 2) * 14).clip(0, 24).astype(np.int16)[..., None]
    noise = rng.integers(-4, 5, size=(h, w, 1), dtype=np.int16)
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
    img = draw_glass_or_swiss_card(W, H, surf, accents[0], border, is_light, radius=38, bar_mode="none")
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
                d.rounded_rectangle([W - 100 - pw, ry - 4, W - 100, ry + 50], radius=14, fill=(*col, 38), outline=col, width=2)
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
    img = draw_glass_or_swiss_card(W, H, surf, accents[0], border, is_light, radius=38, bar_mode="none")
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
        d.rounded_rectangle([margin_x, ly, W - margin_x, ly + layer_h], radius=24, fill=(*surf_alt, 220), outline=(*col, 175), width=2)
        d.text((margin_x + 34, ly + 22), str(layer.get("name", f"LAYER {li + 1}")).upper()[:34], font=f_layer, fill=col)

        nodes = layer.get("nodes", [])[:4]
        n_nodes = max(1, len(nodes))
        inner_top = ly + 74
        inner_h = layer_h - 96
        inner_gap = 36
        node_w = ((W - margin_x * 2 - 68) - inner_gap * (n_nodes - 1)) // n_nodes
        for ni, nd in enumerate(nodes):
            nx = margin_x + 34 + ni * (node_w + inner_gap)
            d.rounded_rectangle([nx, inner_top, nx + node_w, inner_top + inner_h], radius=18, fill=(*surf, 245), outline=(*border, 255), width=2)
            d.rounded_rectangle([nx, inner_top, nx + 8, inner_top + inner_h], radius=4, fill=col)
            d.text((nx + 30, inner_top + 22), str(nd.get("title", ""))[:26], font=f_node_t, fill=ink)
            for wi, line in enumerate(textwrap.wrap(str(nd.get("subtitle", "")), width=max(18, int((node_w - 60) / 17)))[:2]):
                d.text((nx + 30, inner_top + 78 + wi * 38), line, font=f_node_s, fill=muted)

        if li < n_layers - 1:
            ay1, ay2 = ly + layer_h + 6, ly + layer_h + gap_y - 6
            mid_y = (ay1 + ay2) // 2
            for ax in (W // 4, W // 2, (W * 3) // 4):
                d.line([(ax, ay1), (ax, ay2 - 12)], fill=col, width=4)
                d.polygon([(ax, ay2), (ax - 11, ay2 - 17), (ax + 11, ay2 - 17)], fill=col)
            if li < len(flows) and flows[li]:
                fl = flows[li]
                fbb = d.textbbox((0, 0), fl, font=f_tag)
                fw = (fbb[2] - fbb[0]) + 44
                d.rounded_rectangle([W // 2 - fw // 2, mid_y - 20, W // 2 + fw // 2, mid_y + 20], radius=16, fill=(*surf, 255), outline=col, width=2)
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
        "schema_version": int(spec.get("schema_version", 4)),
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
    make_contact_shadow(os.path.join(out_dir, "shadow_hero.png"), 640, 140, 150)

    draw_glass_or_swiss_card(1060, 148, surf_rgb, acc1, border_rgb, is_light, 20, "left").save(os.path.join(out_dir, "card_step.png"), "PNG")
    for ci, ac in enumerate([acc1, acc2, acc3], start=1):
        draw_glass_or_swiss_card(730, 192, surf_rgb, ac, border_rgb, is_light, 20, "left").save(os.path.join(out_dir, f"card_spec_{ci}.png"), "PNG")
        draw_glass_or_swiss_card(552, 420, surf_rgb, ac, border_rgb, is_light, 22, "top").save(os.path.join(out_dir, f"card_bento_tile_{ci}.png"), "PNG")
    draw_glass_or_swiss_card(672, 768, surf_rgb, acc1, border_rgb, is_light, 24, "top").save(os.path.join(out_dir, "card_bento_hero.png"), "PNG")
    draw_glass_or_swiss_card(1732, 296, surf_rgb, acc1, border_rgb, is_light, 24, "left").save(os.path.join(out_dir, "card_bento_strip.png"), "PNG")
    draw_glass_or_swiss_card(1024, 520, surf_rgb, acc1, border_rgb, is_light, 24, "left").save(os.path.join(out_dir, "card_panel.png"), "PNG")
    draw_glass_or_swiss_card(1760, 256, surf_rgb, acc1, border_rgb, is_light, 32, "none").save(os.path.join(out_dir, "dock_bg.png"), "PNG")

    render_soft_dock_glider(os.path.join(out_dir, "dock_glider.png"), width=540, height=264, radius=28)

    render_organic_ink_drips(os.path.join(out_dir, "ep_drips.png"), color_hex=pal["accent1"])
    render_soft_paper_cloud(os.path.join(out_dir, "ep_paper_cloud.png"), paper_hex=pal["surface"])
    render_cloud_sweep(os.path.join(out_dir, "flyby_cloud.png"))
    render_silhouette_horizon(os.path.join(out_dir, "odo_silhouette.png"), color_hex=pal["accent1"])
    render_odometer_masks(os.path.join(out_dir, "odo_mask_top.png"), os.path.join(out_dir, "odo_mask_bot.png"), bg_hex="#08090C")

    resolved_glb_out = _resolve_glb_entry(resolved.get("hero_3d", {}), os.path.join(out_dir, "hero_real.glb"), rel)
    if resolved_glb_out and os.path.exists(resolved_glb_out):
        render_glb_to_image(resolved_glb_out, 640, 640, 12.0, -22.0, 2.0, theme=universe, ssaa=1).save(
            os.path.join(out_dir, "hero_cutout_auto.png"), "PNG"
        )

    news_done = False
    dock_thumbs_done = False
    portal_done = False

    for idx, sl in enumerate(resolved["slides"], start=1):
        st = sl.get("type", "")

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
                cropped = cover_crop(Image.open(res_p).convert("RGB"), 1920, 1080)
                if p_key == "bg_image" and st in ("flyby_wipe", "refracted_spotlight_dock", "closing_finale"):
                    scrim_zone = "left_upper" if st == "refracted_spotlight_dock" else "left"
                    cropped = apply_tonal_text_scrim(cropped, zone=scrim_zone, strength=0.68)
                cropped.save(dst_crop, "PNG")
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
                        gpath,
                        680,
                        680,
                        float(gcfg.get("rx", 12.0)),
                        float(gcfg.get("ry", -24.0)),
                        float(gcfg.get("rz", 2.0)),
                        theme=universe,
                        ssaa=2,
                    ).save(dst_p, "PNG")
                    sl[f"_{field_key}"] = dst_p

        if st == "odometer_counter" and sl.get("bg_hex"):
            render_odometer_masks(
                os.path.join(out_dir, f"s{idx}_odo_mask_top.png"),
                os.path.join(out_dir, f"s{idx}_odo_mask_bot.png"),
                bg_hex=str(sl.get("bg_hex", "#08090C")),
            )

        if st == "real_portal_flythrough" and not portal_done:
            bg_src = _resolve_photo_field(sl, "portal_bg_image", "portal_bg_query", "portal_bg_url", out_dir, f"s{idx}_pbg", rel) or sl.get("_bg_image", "")
            pbg_out = os.path.join(out_dir, "portal_bg.png")
            if bg_src and os.path.exists(bg_src):
                pbg_im = cover_crop(Image.open(bg_src).convert("RGB"), 1920, 1080)
                apply_tonal_text_scrim(pbg_im, pbg_out, zone="left", strength=0.72)
            else:
                shutil.copyfile(os.path.join(out_dir, "bg_base.png"), pbg_out)

            repo_portals = os.path.join(os.path.dirname(SCRIPT_DIR), "docs", "portals")
            cdn_base = "https://raw.githubusercontent.com/7AK3SH1/morphdeck-3d/main/docs/portals"
            local_out = rel(sl.get("portal_outer_image", ""))
            local_in = rel(sl.get("portal_inner_image", ""))

            def _resolve_portal_plate(local_custom: str, plate_filename: str, dst_path: str, fallback_fn) -> None:
                if local_custom and os.path.exists(local_custom):
                    if sl.get("portal_mode") == "photo_punch":
                        punch_portal_mask(local_custom, dst_path, box_frac=(0.29, 0.23, 0.71, 0.59), radius=6)
                    else:
                        shutil.copyfile(local_custom, dst_path)
                    return
                bundled = os.path.join(repo_portals, plate_filename)
                if os.path.exists(bundled):
                    shutil.copyfile(bundled, dst_path)
                    return
                if bg_src:
                    try:
                        from fetch_real_glb import _http_get
                        raw_png = _http_get(f"{cdn_base}/{plate_filename}", timeout=30)
                        if raw_png[:8] == b"\x89PNG\r\n\x1a\n":
                            with open(dst_path, "wb") as fp:
                                fp.write(raw_png)
                            return
                    except Exception:
                        pass
                fallback_fn(dst_path)

            ext_raw_out = os.path.join(out_dir, "portal_outer.png")
            _resolve_portal_plate(
                local_out,
                "portal_exterior.png",
                ext_raw_out,
                lambda dst: render_architectural_portal_exterior(
                    pbg_out, dst, is_light=is_light, stone_hex=pal["surface"] if is_light else "#222630"
                ),
            )
            if os.path.exists(ext_raw_out):
                with Image.open(ext_raw_out) as ext_im:
                    e_rgba = ext_im.convert("RGBA")
                    ea = np.asarray(e_rgba, dtype=np.float32)
                    eh, ew, _ = ea.shape
                    eyy, exx = np.mgrid[0:eh, 0:ew].astype(np.float32)
                    epool = (np.clip((eyy / eh - 0.56) / 0.44, 0.0, 1.0) ** 1.15) * (np.clip((exx / ew - 0.16) / 0.65, 0.0, 1.0) ** 1.1)
                    ea[..., :3] = np.clip(ea[..., :3] * (1.0 - epool[..., None] * 0.54), 0, 255)
                    Image.fromarray(ea.astype(np.uint8), "RGBA").save(ext_raw_out, "PNG")

            _resolve_portal_plate(
                local_in,
                "portal_interior.png",
                os.path.join(out_dir, "portal_inner.png"),
                lambda dst: render_architectural_portal_interior(
                    dst, is_light=is_light, wall_hex=pal["surface"] if is_light else "#1A1E28"
                ),
            )
            portal_done = True

        if st == "vertical_world_elevator":
            bg_src = _resolve_photo_field(sl, "panorama_image", "panorama_query", "panorama_url", out_dir, f"s{idx}_pano", rel) or sl.get("_bg_image", "")
            if bg_src and os.path.exists(bg_src):
                pano_im = cover_crop(Image.open(bg_src).convert("RGB"), 1920, 2160)
                pano_arr = np.asarray(pano_im, dtype=np.float32)
                yy_p, xx_p = np.mgrid[0:2160, 0:1920].astype(np.float32)
                yn_p = yy_p / 2160.0
                xn_p = xx_p / 1920.0
                raw_pano_p = os.path.join(out_dir, "_elev_raw.png")
                pano_im.save(raw_pano_p, "PNG")
                split_horizon_cutout(raw_pano_p, os.path.join(out_dir, "elev_fg.png"), size=(1920, 2160))

                ul_scrim = (np.clip(1.0 - xn_p / 0.58, 0.0, 1.0) ** 1.3) * np.exp(-((yn_p - 0.41) ** 2) / 0.0085) * 0.62
                top_band_scrim = np.clip(1.0 - (yn_p - 0.48) / 0.16, 0.0, 1.0) * np.clip((yn_p - 0.46) / 0.06, 0.0, 1.0) * 0.60
                low_scrim = 0.48 * np.clip((yn_p - 0.52) / 0.48, 0.0, 1.0)
                darken = (1.0 - np.maximum(np.maximum(ul_scrim, low_scrim), top_band_scrim))[..., None]

                Image.fromarray(np.clip(pano_arr * darken, 0, 255).astype(np.uint8), "RGB").save(os.path.join(out_dir, "elev_bg.png"), "PNG")
                with Image.open(os.path.join(out_dir, "elev_fg.png")) as fg_im:
                    fg_arr = np.asarray(fg_im, dtype=np.float32)
                    fg_rgb = np.clip(fg_arr[..., :3] * darken, 0, 255).astype(np.uint8)
                    Image.fromarray(np.dstack([fg_rgb, fg_arr[..., 3].astype(np.uint8)]), "RGBA").save(os.path.join(out_dir, "elev_fg.png"), "PNG")

                lower_view_p = os.path.join(out_dir, "_elev_lower_view.png")
                with Image.open(os.path.join(out_dir, "elev_bg.png")) as ebg_im:
                    ebg_im.crop((0, 1040, 1920, 2120)).save(lower_view_p, "PNG")
                for ci in range(3):
                    # 3-card horizontal frosted glass triptych across lower half of lower_ground
                    cx_pt = 48 + ci * 298
                    box_px = (cx_pt * 2, 348 * 2, (cx_pt + 268) * 2, 488 * 2)
                    render_refracted_card(
                        lower_view_p,
                        os.path.join(out_dir, f"elev_rcard_h_{ci + 1}.png"),
                        crop_box=box_px,
                        radius=22,
                        tint_rgba=(14, 18, 28, 138),
                        accent_hex=pal["accent1"],
                    )
                    # Also render right-stack variant when hero 3D sits on left
                    cy_pt = 142 + ci * 118
                    rbox_px = (888, cy_pt * 2, 1824, (cy_pt + 102) * 2)
                    render_refracted_card(
                        lower_view_p,
                        os.path.join(out_dir, f"elev_rcard_{ci + 1}.png"),
                        crop_box=rbox_px,
                        radius=22,
                        tint_rgba=(14, 18, 28, 138),
                        accent_hex=pal["accent1"],
                    )

        if st == "parchment_3d_split":
            low_src = sl.get("_lower_image") or sl.get("_bg_image") or ""
            if low_src and os.path.exists(low_src):
                cover_crop(Image.open(low_src).convert("RGB"), 1920, 560).save(os.path.join(out_dir, f"s{idx}_split_lower.png"), "PNG")

        if st == "refracted_spotlight_dock":
            bg_src = sl.get("_bg_image", "")
            bg_dst = os.path.join(out_dir, f"s{idx}_bg.png")
            if bg_src and os.path.exists(bg_src):
                shutil.copyfile(bg_src, bg_dst)
            else:
                shutil.copyfile(os.path.join(out_dir, "bg_base.png"), bg_dst)
            render_refracted_glass_dock(
                bg_dst,
                os.path.join(out_dir, f"s{idx}_dock_refract.png"),
                crop_box=(80, 772, 1840, 1044),
                blur_radius=34,
                tint_rgba=(14, 18, 28, 88),
            )
            if not dock_thumbs_done:
                for di, it in enumerate(sl.get("dock", [])[:4], start=1):
                    tsrc = _resolve_photo_field(it, "image", "query", "url", out_dir, f"thumb_{di}", rel)
                    tdst = os.path.join(out_dir, f"thumb_{di}.png")
                    if tsrc and os.path.exists(tsrc):
                        cover_crop(Image.open(tsrc).convert("RGB"), 240, 136).save(tdst, "PNG")
                    else:
                        _build_textured_canvas(240, 136, acc1 if di % 2 else acc2, seed=100 + di).save(tdst, "PNG")
                dock_thumbs_done = True

        if st == "picture_frame_pullback":
            hframe_dst = os.path.join(out_dir, f"s{idx}_hollow_frame.png")
            render_hollow_gold_frame(hframe_dst, frame_style=sl.get("frame_style", "ornate_gold"), border_px=84)
            sl["_hollow_frame"] = hframe_dst
            # Use clean tourist-free parabolic spotlight gallery wall unless user explicitly forces photo_wall
            local_wall = rel(sl.get("wall_image", ""))
            if local_wall and os.path.exists(local_wall):
                sl["_wall_image"] = apply_tonal_text_scrim(local_wall, os.path.join(out_dir, f"s{idx}_wall_clean.png"), zone="left", strength=0.55) and os.path.join(out_dir, f"s{idx}_wall_clean.png")
            else:
                fx = float(sl.get("frame_x", 508))
                fy = float(sl.get("frame_y", 78))
                fw = float(sl.get("frame_w", 340))
                fh = float(sl.get("frame_h", 192))
                wall_dst = os.path.join(out_dir, f"s{idx}_gallery_wall.png")
                render_gallery_wall_room(
                    wall_dst,
                    wall_hex="#1D2028" if is_light else "#141720",
                    spot_cx_frac=(fx + fw * 0.5) / 960.0,
                    spot_cy_frac=(fy + fh * 0.5) / 540.0,
                )
                sl["_wall_image"] = wall_dst
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
    parser = argparse.ArgumentParser(description="Prepare real 3D .glb + museum-grade cinema assets for morphdeck-3d v4.")
    parser.add_argument("--spec", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    prepare_all(args.spec, args.out)


if __name__ == "__main__":
    main()
