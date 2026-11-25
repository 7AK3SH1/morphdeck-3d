#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Cross-Platform OpenXML Presentation Builder (macOS / Linux / Headless CI)
for morphdeck-3d.

Constructs widescreen 16:9 (960 x 540 pt) PowerPoint .pptx presentations using
python-pptx, studio-rendered 3D .glb camera views (via render_glb_preview.py),
Pillow-rendered frosted-glass containers, and ECMA-376 compliant Morph
transitions (<mc:AlternateContent><p159:morph option="byObject"/></mc:AlternateContent>).
"""

import argparse
import io
import json
import os
import re
import sys
import zipfile
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from render_glb_preview import render_glb_to_image  # noqa: E402


def hex_rgb(hex_str: str, default=(37, 99, 235)) -> RGBColor:
    h = (hex_str or "").strip().lstrip("#")
    if len(h) != 6:
        return RGBColor(*default)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_pic(slide, name: str, path: str, x: float, y: float, w: float, h: float):
    if not path or not os.path.exists(path):
        return None
    sh = slide.shapes.add_picture(path, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = name
    return sh


def add_text(slide, name: str, text: str, x: float, y: float, w: float, h: float, font: str, size: float, bold: bool, color: RGBColor, align=PP_ALIGN.LEFT, v_anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.vertical_anchor = v_anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = "" if text is None else str(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bool(bold)
    r.font.color.rgb = color
    return tb


def add_rect(slide, name: str, x: float, y: float, w: float, h: float, fill_col: RGBColor, border_col: RGBColor = None, border_pt: float = 0.0, rounded: bool = False):
    stype = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(stype, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = name
    if rounded and len(sh.adjustments) > 0:
        sh.adjustments[0] = 0.14
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_col
    if border_col is not None and border_pt > 0:
        sh.line.color.rgb = border_col
        sh.line.width = Pt(border_pt)
    else:
        sh.line.fill.background()
    return sh


def add_pill(slide, name: str, text: str, x: float, y: float, w: float, h: float, bg_col: RGBColor, fg_col: RGBColor, size: float, font: str):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = name
    if len(sh.adjustments) > 0:
        sh.adjustments[0] = 0.5
    sh.fill.solid()
    sh.fill.fore_color.rgb = bg_col
    sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = fg_col
    return sh


def add_kpi_box(slide, name: str, val: str, lbl: str, x: float, y: float, w: float, h: float, bg_col: RGBColor, acc_col: RGBColor, fg_col: RGBColor, f_disp: str, f_semi: str):
    sh = add_rect(slide, name, x, y, w, h, bg_col, acc_col, 1.2, rounded=True)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = str(val)
    r1.font.name = f_disp
    r1.font.size = Pt(22)
    r1.font.bold = True
    r1.font.color.rgb = acc_col

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = str(lbl)
    r2.font.name = f_semi
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = fg_col
    return sh


def inject_morph_to_pptx(pptx_path: str, duration_ms: int = 1800) -> None:
    """Inject ECMA-376 compliant mc:AlternateContent + p159:morph on every slide."""
    morph_block = (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main">'
        '<mc:Choice Requires="p159">'
        f'<p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="{duration_ms}">'
        '<p159:morph option="byObject"/>'
        '</p:transition>'
        '</mc:Choice>'
        '<mc:Fallback xmlns="">'
        '<p:transition spd="slow"><p:fade/></p:transition>'
        '</mc:Fallback>'
        '</mc:AlternateContent>'
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r"^ppt/slides/slide\d+\.xml$", item.filename):
                txt = data.decode("utf-8")
                if "p159:morph" not in txt and "</p:sld>" in txt:
                    txt = txt.replace("</p:sld>", f"{morph_block}</p:sld>")
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    with open(pptx_path, "wb") as f:
        f.write(buf.getvalue())


def build_crossplatform_pptx(spec_json: str, assets_dir: str, out_pptx: str) -> str:
    resolved_path = os.path.join(assets_dir, "theme_resolved.json")
    with open(resolved_path if os.path.exists(resolved_path) else spec_json, "r", encoding="utf-8") as f:
        spec = json.load(f)

    pal = spec.get("palette", {})
    fonts = spec.get("fonts", {})
    f_disp = fonts.get("display", "Bahnschrift")
    f_head = fonts.get("heading", f_disp)
    f_body = fonts.get("body", "Segoe UI")
    f_semi = fonts.get("semi", f_body)
    f_mono = fonts.get("mono", "Consolas")

    c_surf = hex_rgb(pal.get("surface", "#141824"))
    c_s_alt = hex_rgb(pal.get("surface_alt", "#1C2233"))
    c_border = hex_rgb(pal.get("border", "#2E384D"))
    c_acc1 = hex_rgb(pal.get("accent1", "#2563EB"))
    c_acc1_d = hex_rgb(pal.get("accent1_dark", "#1D4ED8"))
    c_acc2 = hex_rgb(pal.get("accent2", "#10B981"))
    c_acc3 = hex_rgb(pal.get("accent3", "#F59E0B"))
    c_obs = hex_rgb(pal.get("obsidian", "#090A0D"))
    c_white = RGBColor(255, 255, 255)
    c_light = RGBColor(226, 232, 240)
    c_muted = RGBColor(148, 163, 184)
    c_ink = RGBColor(17, 24, 39)
    c_ink_mid = RGBColor(30, 41, 59)

    hero_arch = spec.get("hero_3d", {}).get("archetype", "ai_quantum_server")
    theme_name = spec.get("theme", "obsidian_kinetic")
    is_phone = hero_arch in ("iphone", "smartphone_pro")
    hero_w, hero_h = (224, 448) if is_phone else (296, 320)

    hero_glb = spec.get("requirements", {}).get("iphone_glb") or os.path.join(assets_dir, "device_iphone.glb")
    laptop_glb = spec.get("requirements", {}).get("macbook_glb") or os.path.join(assets_dir, "device_macbook.glb")

    pose_cache = {}

    def pose_png(glb_path: str, tag: str, rx: float, ry: float, rz: float, res: int = 680) -> str:
        key = (glb_path, round(rx, 1), round(ry, 1), round(rz, 1))
        if key not in pose_cache:
            out_p = os.path.join(assets_dir, f"_pose_{tag}.png")
            render_glb_to_image(glb_path, res, res, rx, ry, rz, theme=theme_name).save(out_p, "PNG")
            pose_cache[key] = out_p
        return pose_cache[key]

    prs = Presentation()
    prs.slide_width = Pt(960)
    prs.slide_height = Pt(540)
    blank = prs.slide_layouts[6]

    slides = spec.get("slides", [])
    brand = spec.get("brand", {})

    for idx, sl in enumerate(slides, start=1):
        s = prs.slides.add_slide(blank)
        st = sl.get("type", "cover_hero")
        prev_st = slides[idx - 2].get("type", "") if idx >= 2 else ""
        next_st = slides[idx].get("type", "") if idx < len(slides) else ""

        if st == "cover_hero":
            style = sl.get("style", "centered_clean")
            if style == "asymmetric_split":
                add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_ivory_grain.png"), 0, 0, 960, 540)
                add_text(s, "c_ovr", sl.get("overline", ""), 52, 40, 460, 20, f_mono, 9.5, True, c_acc1_d)
                add_rect(s, "c_rule", 52, 66, 84, 2.5, c_acc1_d)
                add_pic(s, "c_logo", os.path.join(assets_dir, "logo_badge.png"), 52, 86, 54, 54)
                add_text(s, "c_brand", brand.get("name", ""), 118, 96, 380, 34, f_disp, 20, True, c_ink, v_anchor=MSO_ANCHOR.MIDDLE)
                head = sl.get("headline") or f"{brand.get('title_part1', '')} {brand.get('title_part2', '')}"
                add_text(s, "c_h1", head, 52, 152, 468, 134, f_disp, 40, True, c_ink)
                add_text(s, "c_sub", sl.get("subtitle", ""), 52, 296, 460, 66, f_body, 13.5, False, c_ink_mid)
                for mi, mt in enumerate(sl.get("meta", [])[:2]):
                    add_kpi_box(s, f"c_mk_{mi}", mt.get("value", ""), mt.get("label", ""), 52 + mi * 236, 396, 220, 84, c_obs, c_acc1 if mi == 0 else c_acc2, c_light, f_disp, f_semi)
                add_pic(s, "c_vis", os.path.join(assets_dir, f"s{idx}_cover_visual.png"), 556, 32, 362, 476)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 568, 424, 336, 62)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, -26, 3), 586, 88, 302, 346)
            elif style == "bold_poster":
                add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_finale.png"), 0, 0, 960, 540)
                add_rect(s, "c_band", 0, 0, 960, 10, c_acc1)
                add_text(s, "c_ovr", sl.get("overline", ""), 52, 38, 520, 20, f_mono, 10, True, c_acc1)
                head = sl.get("headline") or f"{brand.get('title_part1', '')}\n{brand.get('title_part2', '')}"
                add_text(s, "c_h1", head, 52, 84, 520, 210, f_head, 54, True, c_white)
                add_rect(s, "c_bar", 52, 312, 136, 5, c_acc2)
                add_text(s, "c_sub", sl.get("subtitle", ""), 52, 332, 500, 76, f_body, 14, False, c_light)
                if sl.get("bottom_cue"):
                    add_pill(s, "c_cue", sl["bottom_cue"], 52, 436, 330, 36, c_acc1, c_white, 10.5, f_mono)
                add_pic(s, "c_vis", os.path.join(assets_dir, f"s{idx}_cover_visual.png"), 588, 36, 330, 468)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 584, 422, 336, 64)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, -24, -4), 602, 74, 304, 358)
            else:
                add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
                add_text(s, "c_ovr", sl.get("overline", ""), 120, 18, 720, 18, f_mono, 9.5, True, c_acc1, PP_ALIGN.CENTER)
                add_pic(s, "c_logo", os.path.join(assets_dir, "logo_badge.png"), 52, 28, 56, 56)
                title = f"{brand.get('title_part1', '')}{brand.get('title_part2', '')}".strip() or brand.get("name", "")
                add_text(s, "c_title", title, 120, 42, 720, 66, f_disp, 46, True, c_white, PP_ALIGN.CENTER)
                add_text(s, "c_sub", sl.get("subtitle", ""), 140, 112, 680, 34, f_body, 12.5, False, c_light, PP_ALIGN.CENTER)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 320, 392, 320, 58)
                mw, mh = (220, 260) if is_phone else (284, 252)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, -22, 3), 480 - mw / 2, 152, mw, mh)
                cols = [c_acc1, c_acc2, c_acc3]
                for mi, mt in enumerate(sl.get("meta", [])[:3]):
                    add_kpi_box(s, f"c_mk_{mi}", mt.get("value", ""), mt.get("label", ""), 132 + mi * 242, 434, 212, 76, c_obs, cols[mi], c_light, f_disp, f_semi)

        elif st == "split_horizon_3d":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_ivory_grain.png"), 0, 0, 960, 540)
            add_pic(s, "!!SplitLower", os.path.join(assets_dir, "split_lower.png"), 0, 275, 960, 265)
            add_text(s, "!!GiantWord", sl.get("giant_word", "PROJECT"), 30, 28, 900, 180, f_disp, 118, True, c_acc1_d, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            add_text(s, "s2_pl", sl.get("left_label", ""), 36, 184, 294, 18, f_disp, 10, True, c_acc1_d)
            add_text(s, "s2_pt", sl.get("left_text", ""), 36, 202, 294, 58, f_semi, 12.5, False, c_ink_mid)
            add_text(s, "s2_sl", sl.get("right_label", ""), 630, 184, 294, 18, f_disp, 10, True, c_acc1_d, PP_ALIGN.RIGHT)
            add_text(s, "s2_st", sl.get("right_text", ""), 630, 202, 294, 58, f_semi, 12.5, False, c_ink_mid, PP_ALIGN.RIGHT)
            kpi_cols = [c_acc1, c_acc3, c_acc2, c_acc2]
            kpi_xs = [36, 192, 624, 780]
            for ki, k in enumerate(sl.get("kpis", [])[:4]):
                add_kpi_box(s, f"s2_k{ki}", k.get("value", ""), k.get("label", ""), kpi_xs[ki], 396, 144, 78, c_obs, kpi_cols[ki], c_light, f_disp, f_semi)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 324, 472, 312, 62)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 8, 12, 0), 480 - hero_w / 2, 116 if not is_phone else 60, hero_w, 330)

        elif st == "portal_flythrough_3d":
            stage = sl.get("stage", "exterior")
            portal_bg = os.path.join(assets_dir, "portal_bg.png")
            portal_in = os.path.join(assets_dir, "portal_inner.png")
            portal_out = os.path.join(assets_dir, "portal_outer.png")
            if stage == "interior":
                add_pic(s, "!!portal_bg", portal_bg, -64, -36, 1088, 612)
                add_pic(s, "!!portal_outer", portal_out, -1728, -972, 4416, 2484)
                add_pic(s, "!!portal_inner", portal_in, 0, 0, 960, 540)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 526, 408, 348, 64)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, -20, 2, res=760), 544, 72, 316, 352)
                add_pic(s, "!!pt_card_next", os.path.join(assets_dir, "card_bento_wide.png"), 64, 108, 430, 324)
                add_text(s, "pt_ovr", sl.get("overline", ""), 88, 126, 382, 20, f_mono, 9.5, True, c_acc1)
                add_text(s, "pt_h1", sl.get("headline", ""), 88, 150, 382, 68, f_disp, 24, True, c_white)
                add_text(s, "pt_body", sl.get("body", sl.get("subtitle", "")), 88, 224, 382, 92, f_body, 11.5, False, c_light)
                for ki, k in enumerate(sl.get("kpis", [])[:2]):
                    add_kpi_box(s, f"pt_k_{ki}", k.get("value", ""), k.get("label", ""), 88 + ki * 194, 330, 182, 76, c_obs, c_acc1 if ki == 0 else c_acc2, c_light, f_disp, f_semi)
            elif stage == "horizon":
                add_pic(s, "!!portal_bg", portal_bg, -120, -68, 1200, 675)
                add_pic(s, "!!portal_outer", portal_out, -1728, -972, 4416, 2484)
                add_pic(s, "!!portal_inner", portal_in, -1536, -864, 4032, 2268)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 522, 424, 388, 68)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, 24, -3, res=760), 540, 58, 356, 384)
                add_text(s, "pt_ovr", sl.get("overline", ""), 52, 46, 460, 20, f_mono, 9.5, True, c_acc1)
                add_text(s, "pt_h1", sl.get("headline", ""), 52, 72, 460, 96, f_disp, 32, True, c_white)
                add_text(s, "pt_body", sl.get("body", sl.get("subtitle", "")), 52, 180, 440, 110, f_body, 12.5, False, c_light)
                kcols = [c_acc1, c_acc2, c_acc3]
                for ki, k in enumerate(sl.get("kpis", [])[:3]):
                    add_kpi_box(s, f"pt_k_{ki}", k.get("value", ""), k.get("label", ""), 52 + ki * 158, 396, 146, 78, c_obs, kcols[ki], c_light, f_disp, f_semi)
            else:
                add_pic(s, "!!portal_bg", portal_bg, 0, 0, 960, 540)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 380, 344, 200, 44)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 16, -28, -4), 392, 146, 176, 214)
                add_pic(s, "!!portal_inner", portal_in, 272, 102, 416, 336)
                add_pic(s, "!!portal_outer", portal_out, 0, 0, 960, 540)
                add_text(s, "pt_ovr", sl.get("overline", ""), 48, 20, 864, 18, f_mono, 9.5, True, c_acc1, PP_ALIGN.CENTER)
                add_text(s, "pt_h1", sl.get("headline", ""), 48, 40, 864, 42, f_disp, 25, True, c_white, PP_ALIGN.CENTER)
                for mi, mt in enumerate(sl.get("meta", [])[:2]):
                    mx = 34 if mi == 0 else 720
                    add_kpi_box(s, f"pt_m_{mi}", mt.get("value", ""), mt.get("label", ""), mx, 204, 206, 112, c_obs, c_acc1 if mi == 0 else c_acc2, c_light, f_disp, f_semi)
                if sl.get("bottom_cue"):
                    add_pill(s, "pt_cue", sl["bottom_cue"], 290, 480, 380, 34, c_obs, c_acc1, 10, f_mono)
                if next_st == "portal_flythrough_3d":
                    add_pic(s, "!!pt_card_next", os.path.join(assets_dir, "card_bento_wide.png"), -520, 108, 430, 324)

        elif st == "vertical_elevator_3d":
            stop = sl.get("elevator_stop", "upper_sky")
            elev_bg = os.path.join(assets_dir, "elev_bg.png")
            elev_fg = os.path.join(assets_dir, "elev_fg.png")
            gw = sl.get("giant_word", "ORBITAL")
            if stop == "lower_ground":
                add_pic(s, "!!elev_bg", elev_bg, 0, -520, 960, 1080)
                add_text(s, "!!elev_word", gw, 40, 290, 880, 170, f_head, 126, True, c_white, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                add_pic(s, "!!elev_fg", elev_fg, 0, -520, 960, 1080)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 46, 418, 356, 62)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 16, 28, -3, res=760), 62, 74, 328, 362)
                add_text(s, "ev_ovr2", sl.get("overline", ""), 440, 34, 474, 20, f_mono, 9.5, True, c_acc1)
                add_text(s, "ev_h2", sl.get("headline", ""), 440, 56, 474, 68, f_disp, 25, True, c_white)
                cards = sl.get("cards") or sl.get("kpis") or []
                cols = [c_acc1, c_acc2, c_acc3]
                for ci, crd in enumerate(cards[:3]):
                    cy = 142 + ci * 120
                    add_pic(s, f"!!ev_rcard_{ci}", os.path.join(assets_dir, f"card_spec_{ci + 1}.png"), 440, cy, 474, 104)
                    add_text(s, f"ev_rmet_{ci}", crd.get("value", crd.get("metric", "")), 462, cy + 16, 128, 68, f_disp, 28, True, cols[ci], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                    add_text(s, f"ev_rttl_{ci}", crd.get("title", crd.get("label", "")), 598, cy + 16, 294, 26, f_disp, 14, True, c_white)
                    add_text(s, f"ev_rdsc_{ci}", crd.get("desc", ""), 598, cy + 46, 294, 44, f_body, 11, False, c_light)
            else:
                add_pic(s, "!!elev_bg", elev_bg, 0, 0, 960, 1080)
                add_text(s, "!!elev_word", gw, 40, 94, 880, 156, f_head, 126, True, c_white, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                add_pic(s, "!!elev_fg", elev_fg, 0, 0, 960, 1080)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 548, 412, 356, 62)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 10, -22, 3, res=760), 564, 68, 326, 360)
                add_pic(s, "ev_panel", os.path.join(assets_dir, "card_bento_wide.png"), 46, 250, 468, 242)
                add_text(s, "ev_ovr", sl.get("overline", ""), 72, 268, 416, 18, f_mono, 9.5, True, c_acc1)
                add_text(s, "ev_h1", sl.get("headline", ""), 72, 292, 416, 62, f_disp, 22, True, c_white)
                add_text(s, "ev_body", sl.get("body", ""), 72, 358, 416, 76, f_body, 11.5, False, c_light)
                if sl.get("badge"):
                    add_pill(s, "ev_badge", sl["badge"], 72, 444, 198, 28, c_acc1, c_white, 9.5, f_mono)
                if next_st == "vertical_elevator_3d":
                    for ci, ox in enumerate([1080, 1240, 1400]):
                        add_pic(s, f"!!ev_rcard_{ci}", os.path.join(assets_dir, f"card_spec_{ci + 1}.png"), ox, 142 + ci * 120, 474, 104)

        elif st == "depth_sandwich_pan":
            pan_state = sl.get("pan_state", "wide")
            if pan_state == "zoom_right":
                bx, by, bw, bh = -240, -96, 1344, 756
                mx, my, mw, mh, rx, ry, rz = 528, 46, 366, 396, 14, -28, 3
                wx, wy, ww, wh, wsz, wal = 44, 96, 660, 190, 110, PP_ALIGN.LEFT
                cap_x, cap_y, cap_w = 46, 346, 448
            elif pan_state == "zoom_left":
                bx, by, bw, bh = -120, -108, 1344, 756
                mx, my, mw, mh, rx, ry, rz = 58, 46, 366, 396, 14, 28, -3
                wx, wy, ww, wh, wsz, wal = 260, 96, 650, 190, 110, PP_ALIGN.RIGHT
                cap_x, cap_y, cap_w = 466, 346, 448
            else:
                bx, by, bw, bh = 0, 0, 960, 540
                mx, my, mw, mh, rx, ry, rz = 318, 46, 324, 356, 10, -14, 2
                wx, wy, ww, wh, wsz, wal = 30, 88, 900, 200, 136, PP_ALIGN.CENTER
                cap_x, cap_y, cap_w = 46, 384, 416

            add_pic(s, "!!pan_bg", os.path.join(assets_dir, f"s{idx}_pan_bg.png"), bx, by, bw, bh)
            add_text(s, "!!pan_word", sl.get("giant_word", "CORE"), wx, wy, ww, wh, f_head, wsz, True, c_white, wal, MSO_ANCHOR.MIDDLE)
            add_pic(s, "!!pan_fg", os.path.join(assets_dir, f"s{idx}_pan_fg.png"), bx, by, bw, bh)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), mx - 12, my + mh - 12, mw + 24, 64)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", rx, ry, rz), mx, my, mw, mh)
            add_pill(s, "pan_ovr", sl.get("overline", ""), 46, 28, 336, 28, c_obs, c_acc1, 9.5, f_mono)
            add_pic(s, "pan_cap_bg", os.path.join(assets_dir, "card_bento_wide.png"), cap_x, cap_y, cap_w, 124)
            add_text(s, "pan_h", sl.get("headline", ""), cap_x + 24, cap_y + 14, cap_w - 40, 34, f_disp, 18, True, c_white)
            add_text(s, "pan_b", sl.get("caption", ""), cap_x + 24, cap_y + 52, cap_w - 40, 56, f_body, 11.5, False, c_light)

        elif st == "macro_dolly_3d":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            add_text(s, "md_ovr", sl.get("overline", ""), 46, 22, 520, 18, f_mono, 10, True, c_acc1)
            add_text(s, "md_h1", sl.get("headline", ""), 46, 46, 470, 76, f_disp, 25, True, c_white)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_laptop.png"), 486, 436, 440, 76)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 22, -38, -4, res=760), 504, 42, 416, 428)
            cols = [c_acc1, c_acc2, c_acc3]
            for ci, co in enumerate(sl.get("callouts", [])[:3]):
                cy = 142 + ci * 116
                ac = cols[ci]
                add_pic(s, f"md_card_{ci}", os.path.join(assets_dir, f"card_spec_{ci + 1}.png"), 46, cy, 384, 96)
                add_text(s, f"md_ttl_{ci}", co.get("title", ""), 68, cy + 14, 344, 22, f_disp, 13, True, ac)
                add_text(s, f"md_dsc_{ci}", co.get("desc", ""), 68, cy + 40, 344, 46, f_body, 11, False, c_light)
                add_rect(s, f"md_line_{ci}", 430, cy + 46, 96, 2, ac)
                dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Pt(522), Pt(cy + 42), Pt(10), Pt(10))
                dot.name = f"md_dot_{ci}"
                dot.fill.solid()
                dot.fill.fore_color.rgb = ac
                dot.line.fill.background()

        elif st == "timeline_morph_step":
            stages = sl.get("stages", [])[:4]
            n_st = max(1, len(stages))
            act = max(0, min(n_st - 1, int(sl.get("active_index", 0))))
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
            add_text(s, "tl_ovr", sl.get("overline", ""), 52, 28, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "tl_h1", sl.get("headline", ""), 52, 50, 760, 48, f_disp, 27, True, c_white)

            rail_x0, rail_span = 76, 808
            add_rect(s, "!!tl_rail", rail_x0, 144, rail_span, 4, c_border)
            centers = [480.0] if n_st == 1 else [rail_x0 + 70 + (rail_span - 140) * (ti / (n_st - 1)) for ti in range(n_st)]
            act_cx = centers[act]
            add_rect(s, "!!tl_progress", rail_x0, 144, max(24.0, act_cx - rail_x0), 4, c_acc1)

            for ti, stg in enumerate(stages):
                cx = centers[ti]
                nw, nh, ny = (172, 36, 128) if ti == act else (148, 30, 131)
                bg = c_acc1 if ti == act else c_surf
                fg = c_white if ti == act else (c_acc1 if ti < act else c_muted)
                nx = max(48.0, min(912.0 - nw, cx - nw / 2))
                add_pill(s, f"!!tl_node_{ti}", f"{stg.get('code', '')}  {stg.get('label', '')}", nx, ny, nw, nh, bg, fg, 9.5, f_mono)

            add_rect(s, "!!tl_pointer", max(326.0, min(836.0, act_cx - 36)), 192, 72, 4, c_acc1)
            ry_opts = [22, -24, 38, -16]
            rx_opts = [8, 16, 12, 6]
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 32, 448, 276, 54)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", rx_opts[act % 4], ry_opts[act % 4], 2), 44, 192, 252, 276)

            cur = stages[act]
            add_pic(s, "!!tl_drawer", os.path.join(assets_dir, "card_timeline_drawer.png"), 316, 204, 596, 292)
            add_text(s, "tl_code", f"STAGE  {cur.get('code', '')}", 344, 226, 176, 20, f_mono, 10, True, c_acc1)
            add_text(s, "tl_metric", cur.get("metric", ""), 344, 252, 176, 76, f_disp, 42, True, c_white)
            add_text(s, "tl_mlbl", cur.get("metric_label", ""), 344, 334, 176, 36, f_semi, 11, False, c_light)
            if cur.get("tag"):
                add_pill(s, "tl_tag", cur["tag"], 344, 422, 164, 30, c_s_alt, c_acc1, 9, f_mono)
            add_rect(s, "tl_vdiv", 526, 228, 1, 244, c_border)
            add_text(s, "tl_sh1", cur.get("title", ""), 544, 226, 346, 52, f_disp, 18.5, True, c_white)
            add_text(s, "tl_sum", cur.get("summary", ""), 544, 284, 346, 58, f_body, 11.5, False, c_light)
            for bi, b in enumerate(cur.get("bullets", [])[:3]):
                by = 354 + bi * 40
                add_rect(s, f"tl_bdot_{bi}", 546, by + 9, 7, 7, c_acc1)
                add_text(s, f"tl_btxt_{bi}", b, 564, by, 328, 32, f_semi, 11, False, c_white)

        elif st == "kpi_bento":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
            add_text(s, "kb_ovr", sl.get("overline", ""), 48, 26, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "kb_h1", sl.get("headline", ""), 48, 48, 700, 46, f_disp, 27, True, c_white)
            hero = sl.get("hero_tile", {})
            if sl.get("layout") == "top_hero_strip":
                add_pic(s, "kb_hero", os.path.join(assets_dir, "card_bento_strip.png"), 46, 106, 588, 154)
                add_text(s, "kb_htag", hero.get("label", ""), 74, 120, 180, 18, f_mono, 9.5, True, c_acc1)
                add_text(s, "kb_hval", hero.get("value", ""), 74, 144, 180, 84, f_disp, 44, True, c_white)
                add_rect(s, "kb_vdiv", 262, 128, 1, 108, c_border)
                add_text(s, "kb_hsub", hero.get("title", ""), 280, 124, 334, 30, f_disp, 16.5, True, c_white)
                add_text(s, "kb_hdsc", hero.get("desc", ""), 280, 158, 334, 62, f_body, 11.5, False, c_light)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 654, 234, 260, 44)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, -28, 3), 668, 42, 232, 214)
                cols = [c_acc1, c_acc2, c_acc3]
                for ti, t in enumerate(sl.get("tiles", [])[:3]):
                    tx = 46 + ti * 296
                    add_pic(s, f"kb_tbg_{ti}", os.path.join(assets_dir, f"card_bento_tile_{ti + 1}.png"), tx, 276, 276, 222)
                    add_text(s, f"kb_tlbl_{ti}", t.get("label", ""), tx + 22, 298, 232, 20, f_mono, 9.5, True, cols[ti])
                    add_text(s, f"kb_tval_{ti}", t.get("value", ""), tx + 22, 324, 232, 62, f_disp, 38, True, c_white)
                    add_text(s, f"kb_tdsc_{ti}", t.get("desc", ""), tx + 22, 396, 232, 82, f_body, 11.5, False, c_light)
            else:
                add_pic(s, "kb_hero", os.path.join(assets_dir, "card_bento_hero.png"), 46, 110, 284, 390)
                add_text(s, "kb_htag", hero.get("label", ""), 70, 134, 236, 20, f_mono, 9.5, True, c_acc1)
                add_text(s, "kb_hval", hero.get("value", ""), 70, 162, 236, 82, f_disp, 48, True, c_white)
                add_text(s, "kb_hsub", hero.get("title", ""), 70, 250, 236, 30, f_disp, 15, True, c_white)
                add_text(s, "kb_hdsc", hero.get("desc", ""), 70, 286, 236, 116, f_body, 12, False, c_light)
                if hero.get("badge"):
                    add_pill(s, "kb_hcap", hero["badge"], 70, 436, 210, 32, c_acc1, c_white, 9.5, f_mono)
                cols = [c_acc1, c_acc2]
                for ti, t in enumerate(sl.get("tiles", [])[:2]):
                    tx = 344 + ti * 198
                    add_pic(s, f"kb_tbg_{ti}", os.path.join(assets_dir, f"card_bento_tile_{ti + 1}.png"), tx, 110, 186, 184)
                    add_text(s, f"kb_tlbl_{ti}", t.get("label", ""), tx + 18, 128, 150, 18, f_mono, 9, True, cols[ti])
                    add_text(s, f"kb_tval_{ti}", t.get("value", ""), tx + 18, 150, 150, 54, f_disp, 32, True, c_white)
                    add_text(s, f"kb_tdsc_{ti}", t.get("desc", ""), tx + 18, 210, 150, 68, f_body, 10.5, False, c_light)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 738, 264, 184, 42)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, -26, 3), 744, 64, 176, 218)
                wide = sl.get("wide_tile", {})
                add_pic(s, "kb_wide", os.path.join(assets_dir, "card_bento_wide.png"), 344, 310, 570, 190)
                add_text(s, "kb_wovr", wide.get("overline", ""), 368, 326, 520, 18, f_mono, 9, True, c_acc2)
                add_text(s, "kb_wh1", wide.get("title", ""), 368, 348, 520, 30, f_disp, 16.5, True, c_white)
                for ci, col in enumerate(wide.get("columns", [])[:2]):
                    cx = 368 + ci * 264
                    add_text(s, f"kb_wcl_{ci}", col.get("label", ""), cx, 390, 246, 20, f_semi, 11.5, True, c_acc1)
                    add_text(s, f"kb_wcd_{ci}", col.get("desc", ""), cx, 414, 246, 66, f_body, 11, False, c_light)

        elif st == "chart_infographic":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
            add_text(s, "ch_ovr", sl.get("overline", ""), 46, 26, 640, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "ch_h1", sl.get("headline", ""), 46, 48, 680, 46, f_disp, 27, True, c_white)
            add_pic(s, "ch_img", os.path.join(assets_dir, f"s{idx}_chart.png"), 46, 110, 584, 384)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 656, 244, 248, 44)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, -24, 2), 672, 38, 220, 222)
            hl = sl.get("highlight", {})
            add_pic(s, "ch_hlbg", os.path.join(assets_dir, "card_spec_1.png"), 648, 278, 268, 104)
            add_text(s, "ch_hlv", f"{hl.get('value', '')}  {hl.get('label', '')}", 670, 290, 232, 30, f_disp, 18, True, c_acc1)
            add_text(s, "ch_hld", hl.get("desc", ""), 670, 324, 232, 48, f_body, 10.5, False, c_light)
            if sl.get("insights"):
                ins = sl["insights"][0]
                add_pic(s, "ch_inbg_0", os.path.join(assets_dir, "card_spec_2.png"), 648, 394, 268, 100)
                add_text(s, "ch_int_0", ins.get("title", ""), 670, 406, 232, 24, f_disp, 13, True, c_acc2)
                add_text(s, "ch_ind_0", ins.get("desc", ""), 670, 434, 232, 50, f_body, 10.5, False, c_light)

        elif st == "architecture_diagram":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            add_text(s, "ar_ovr", sl.get("overline", ""), 42, 26, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "ar_h1", sl.get("headline", ""), 42, 48, 640, 46, f_disp, 26, True, c_white)
            add_pic(s, "ar_img", os.path.join(assets_dir, f"s{idx}_arch.png"), 42, 104, 624, 398)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 686, 244, 236, 42)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 22, -34, -3), 698, 36, 212, 222)
            cols = [c_acc1, c_acc2]
            for pi, pl in enumerate(sl.get("pillars", [])[:2]):
                sy = 276 + pi * 116
                add_pic(s, f"ar_cbg_{pi}", os.path.join(assets_dir, f"card_spec_{pi + 1}.png"), 682, sy, 236, 106)
                add_text(s, f"ar_ctag_{pi}", f"{pl.get('tag', '')}  //  {pl.get('title', '')}", 702, sy + 12, 204, 22, f_disp, 12, True, cols[pi])
                add_text(s, f"ar_cdsc_{pi}", pl.get("desc", ""), 702, sy + 38, 204, 58, f_body, 10.5, False, c_light)

        elif st == "editorial_split":
            side = sl.get("image_side", "right")
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
            img_x, txt_x, mx, ry = (44, 504, 228, 24) if side == "left" else (508, 48, 476, -24)
            add_pic(s, "es_vis", os.path.join(assets_dir, f"s{idx}_split_visual.png"), img_x, 34, 406, 468)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), mx - 10, 432, 280, 54)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, ry, 2), mx, 118, 260, 326)
            add_text(s, "es_ovr", sl.get("overline", ""), txt_x, 44, 416, 20, f_mono, 9.5, True, c_acc1)
            add_rect(s, "es_rule", txt_x, 70, 72, 2.0, c_acc1)
            add_text(s, "es_h1", sl.get("headline", ""), txt_x, 88, 416, 108, f_disp, 30, True, c_white)
            add_text(s, "es_body", sl.get("body", ""), txt_x, 206, 416, 104, f_body, 13.5, False, c_light)
            fcols = [c_acc1, c_acc2]
            for fi, fc in enumerate(sl.get("facts", [])[:2]):
                fx = txt_x + fi * 212
                add_pic(s, f"es_fcard_{fi}", os.path.join(assets_dir, f"card_spec_{fi + 1}.png"), fx, 342, 198, 136)
                add_text(s, f"es_fv_{fi}", fc.get("value", ""), fx + 18, 354, 166, 40, f_disp, 26, True, fcols[fi])
                add_text(s, f"es_fl_{fi}", fc.get("label", ""), fx + 18, 396, 166, 18, f_mono, 8.5, True, c_white)
                add_text(s, f"es_fd_{fi}", fc.get("desc", ""), fx + 18, 416, 166, 52, f_body, 10, False, c_light)

        elif st == "orbit_steps_3d":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
            p_side = sl.get("side", "left")
            if p_side == "right":
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 628, 468, 296, 60)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 6, -21, 3), 636, 72, hero_w, 370)
                col_x = 46
            else:
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 36, 468, 296, 60)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 6, 21, -3), 54, 72, hero_w, 370)
                col_x = 366
            add_text(s, "s4_ovr", sl.get("overline", ""), col_x, 36, 540, 20, f_mono, 10.5, True, c_acc1)
            add_text(s, "s4_h1", sl.get("headline", ""), col_x, 56, 550, 72, f_disp, 29, True, c_white)
            for si, stp in enumerate(sl.get("steps", [])[:4]):
                sy = 132 + si * 90
                add_pic(s, f"s4_card_{si}", os.path.join(assets_dir, "card_step.png"), col_x, sy, 546, 76)
                add_text(s, f"s4_num_{si}", stp.get("num", ""), col_x + 20, sy + 14, 56, 48, f_disp, 26, True, c_acc1, PP_ALIGN.CENTER)
                add_text(s, f"s4_ttl_{si}", stp.get("title", ""), col_x + 86, sy + 13, 440, 24, f_disp, 13.5, True, c_white)
                add_text(s, f"s4_dsc_{si}", stp.get("desc", ""), col_x + 86, sy + 40, 445, 24, f_body, 11.5, False, c_light)

        elif st == "spotlight_dock":
            d_idx = int(sl.get("dock_index", 0))
            dock_xs = [66, 282, 498, 714]
            glider_xs = [54, 270, 486, 702]
            if next_st == "spotlight_dock":
                nbg = add_pic(s, f"!!sp_bg_{idx + 1}", os.path.join(assets_dir, f"s{idx + 1}_bg.png"), 0, 0, 960, 540)
                if nbg is not None:
                    nbg.crop_top = -0.95
                    nbg.crop_bottom = 0.95
            add_pic(s, f"!!sp_bg_{idx}", os.path.join(assets_dir, f"s{idx}_bg.png"), 0, 0, 960, 540)
            if prev_st == "spotlight_dock":
                prev_sp = slides[idx - 2].get("spotlight", {})
                add_text(s, f"!!sp_tag_{idx - 1}", prev_sp.get("tag", ""), 50, -280, 540, 20, f_mono, 10, True, c_acc1)
                add_text(s, f"!!sp_title_{idx - 1}", prev_sp.get("title", ""), 50, -240, 550, 64, f_disp, 42, True, c_white)
                add_text(s, f"!!sp_sub_{idx - 1}", prev_sp.get("subtitle", ""), 50, -180, 560, 26, f_semi, 13, False, c_acc3)
                add_text(s, f"!!sp_body_{idx - 1}", prev_sp.get("body", ""), 50, -140, 520, 112, f_body, 13.5, False, c_light)
            sp = sl.get("spotlight", {})
            add_text(s, f"!!sp_tag_{idx}", sp.get("tag", ""), 50, 48, 540, 20, f_mono, 10, True, c_acc1)
            add_text(s, f"!!sp_title_{idx}", sp.get("title", ""), 50, 74, 550, 64, f_disp, 42, True, c_white)
            add_text(s, f"!!sp_sub_{idx}", sp.get("subtitle", ""), 50, 142, 560, 26, f_semi, 13, False, c_acc3)
            add_text(s, f"!!sp_body_{idx}", sp.get("body", ""), 50, 180, 520, 112, f_body, 13.5, False, c_light)
            if next_st == "spotlight_dock":
                nx_sp = slides[idx].get("spotlight", {})
                add_text(s, f"!!sp_tag_{idx + 1}", nx_sp.get("tag", ""), -460, 48, 540, 20, f_mono, 10, True, c_acc1)
                add_text(s, f"!!sp_title_{idx + 1}", nx_sp.get("title", ""), -560, 74, 550, 64, f_disp, 42, True, c_white)
                add_text(s, f"!!sp_sub_{idx + 1}", nx_sp.get("subtitle", ""), -680, 142, 560, 26, f_semi, 13, False, c_acc3)
                add_text(s, f"!!sp_body_{idx + 1}", nx_sp.get("body", ""), -800, 180, 520, 112, f_body, 13.5, False, c_light)
            if not is_phone or sl.get("show_3d"):
                ry_sp = -16 if d_idx % 2 == 0 else 22
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 644, 304, 268, 46)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 10, ry_sp, 2), 652, 38, 252, 276)
                if sp.get("badge"):
                    add_pill(s, "sp_bdg", sp["badge"], 686, 328, 184, 30, c_obs, c_acc3, 9.5, f_mono)
            else:
                add_pic(s, "sp_port", os.path.join(assets_dir, f"s{idx}_portrait.png"), 656, 36, 244, 282)
            refract_dock = os.path.join(assets_dir, f"s{idx}_dock_refract.png")
            dock_img = refract_dock if os.path.exists(refract_dock) else os.path.join(assets_dir, "dock_bg.png")
            add_pic(s, "!!GlassDock", dock_img, 40, 386, 880, 128)
            add_pic(s, "!!DockGlider", os.path.join(assets_dir, "dock_glider.png"), glider_xs[min(3, max(0, d_idx))], 396, 208, 108)
            dock_items = sl.get("dock") or spec.get("dock", [])
            for di in range(4):
                it = dock_items[di] if di < len(dock_items) else {"title": f"Item {di + 1}", "subtitle": ""}
                dx = dock_xs[di]
                add_pic(s, f"!!DockThumb{di + 1}", os.path.join(assets_dir, f"thumb_{di + 1}.png"), dx, 416, 76, 60)
                add_text(s, f"!!DockLbl{di + 1}", it.get("title", ""), dx + 84, 418, 110, 24, f_disp, 13, True, c_white)
                add_text(s, f"!!DockSub{di + 1}", it.get("subtitle", ""), dx + 84, 444, 110, 22, f_body, 11, False, c_muted)
            for pi in range(2):
                is_act = (pi == (d_idx % 2))
                dw = 14 if is_act else 7
                dxp = 930 if is_act else 933.5
                dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Pt(dxp), Pt(222 + pi * 24), Pt(dw), Pt(dw))
                dot.name = f"!!sp_dot_{pi}"
                dot.fill.solid()
                dot.fill.fore_color.rgb = c_acc1 if is_act else c_muted
                dot.line.fill.background()

        elif st == "macbook_studio_3d":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            m_side = sl.get("side", "right")
            m_col_x, lap_sh_x, lap_x, lap_ry = (548, 8, 28, 22) if m_side == "left" else (44, 408, 432, -22)
            add_text(s, "m_ovr", sl.get("overline", ""), m_col_x, 40, 366, 20, f_mono, 10, True, c_acc2)
            add_text(s, "m_h1", sl.get("headline", ""), m_col_x, 64, 366, 78, f_disp, 28, True, c_white)
            cols = [c_acc1, c_acc2, c_acc3]
            for pi, pl in enumerate(sl.get("pillars", [])[:3]):
                sy = 152 + pi * 112
                add_pic(s, f"m_card_{pi + 1}", os.path.join(assets_dir, f"card_spec_{pi + 1}.png"), m_col_x, sy, 366, 96)
                add_text(s, f"m_ttl_{pi + 1}", pl.get("title", ""), m_col_x + 20, sy + 14, 334, 22, f_disp, 12.5, True, cols[pi])
                add_text(s, f"m_dsc_{pi + 1}", pl.get("desc", ""), m_col_x + 20, sy + 40, 334, 48, f_body, 11, False, c_light)
            if laptop_glb and os.path.exists(laptop_glb):
                add_pic(s, "!!LaptopShadow", os.path.join(assets_dir, "shadow_laptop.png"), lap_sh_x, 398, 536, 84)
                add_pic(s, "!!Laptop3D", pose_png(laptop_glb, f"s{idx}_lap", 10, lap_ry, 0, res=780), lap_x, 76, 498, 368)

        elif st == "problem_solution_split":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_main.png"), 0, 0, 960, 540)
            add_text(s, "ps_ovr", sl.get("overline", ""), 42, 24, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "ps_h1", sl.get("headline", ""), 42, 46, 720, 44, f_disp, 27, True, c_white)
            left = sl.get("left", {})
            add_pic(s, "ps_lbg", os.path.join(assets_dir, "card_split_left.png"), 38, 102, 302, 404)
            add_text(s, "ps_ltag", left.get("tag", ""), 60, 120, 258, 18, f_mono, 9, True, c_muted)
            add_text(s, "ps_ltl", left.get("title", ""), 60, 140, 258, 32, f_disp, 17, True, c_white)
            for li, it in enumerate(left.get("items", [])[:3]):
                ry = 186 + li * 100
                add_rect(s, f"ps_lb_{li}", 58, ry + 8, 3, 64, c_muted)
                add_text(s, f"ps_lit_{li}", it.get("title", ""), 72, ry + 6, 250, 24, f_semi, 12.5, True, c_white)
                add_text(s, f"ps_lid_{li}", it.get("desc", ""), 72, ry + 32, 250, 46, f_body, 10.5, False, c_muted)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 346, 420, 268, 54)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, 18, -3), 354, 114, 252, 318)
            add_pill(s, "ps_bridge", sl.get("bridge_label", "SHIFT"), 432, 456, 96, 32, c_acc1, c_white, 10, f_mono)
            right = sl.get("right", {})
            add_pic(s, "ps_rbg", os.path.join(assets_dir, "card_split_right.png"), 620, 102, 302, 404)
            add_text(s, "ps_rtag", right.get("tag", ""), 642, 120, 258, 18, f_mono, 9, True, c_acc1)
            add_text(s, "ps_rtl", right.get("title", ""), 642, 140, 258, 32, f_disp, 17, True, c_white)
            for ri, it in enumerate(right.get("items", [])[:3]):
                ry = 186 + ri * 100
                add_rect(s, f"ps_rb_{ri}", 640, ry + 8, 4, 64, c_acc1)
                add_text(s, f"ps_rit_{ri}", it.get("title", ""), 654, ry + 6, 250, 24, f_disp, 13, True, c_acc1)
                add_text(s, f"ps_rid_{ri}", it.get("desc", ""), 654, ry + 32, 250, 46, f_body, 10.5, False, c_light)

        elif st == "comparison_matrix":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            add_text(s, "cm_ovr", sl.get("overline", ""), 42, 24, 640, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "cm_h1", sl.get("headline", ""), 42, 46, 700, 44, f_disp, 27, True, c_white)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 28, 438, 228, 50)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, 24, -3), 36, 122, 212, 326)
            cols = sl.get("columns", ["Criterion", "Legacy", "Alternative", brand.get("name", "Solution")])
            xs_ws = [(266, 206, c_s_alt, c_muted, PP_ALIGN.LEFT), (480, 134, c_s_alt, c_muted, PP_ALIGN.CENTER), (622, 134, c_s_alt, c_muted, PP_ALIGN.CENTER), (764, 154, c_acc1, c_white, PP_ALIGN.CENTER)]
            for ci, (cx, cw, bgc, fgc, aln) in enumerate(xs_ws):
                add_rect(s, f"cm_hc{ci}", cx, 106, cw, 40, bgc, c_border if ci < 3 else c_acc1, 1.0, rounded=True)
                add_text(s, f"cm_ht{ci}", cols[ci] if ci < len(cols) else "", cx + 6, 106, cw - 12, 40, f_disp, 11, True, fgc, aln, MSO_ANCHOR.MIDDLE)
            for ri, row in enumerate(sl.get("rows", [])[:5]):
                ry = 156 + ri * 66
                add_pic(s, f"cm_rbg_{ri}", os.path.join(assets_dir, "card_step.png"), 266, ry, 490, 58)
                add_text(s, f"cm_rl_{ri}", row.get("criterion", ""), 284, ry, 186, 58, f_semi, 11.5, True, c_white, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
                add_text(s, f"cm_rv1_{ri}", row.get("col1", ""), 486, ry, 122, 58, f_body, 10.5, False, c_muted, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                add_text(s, f"cm_rv2_{ri}", row.get("col2", ""), 628, ry, 122, 58, f_body, 10.5, False, c_muted, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                add_rect(s, f"cm_rhi_{ri}", 764, ry, 154, 58, c_surf, c_acc1, 1.6, rounded=True)
                add_text(s, f"cm_rv3_{ri}", row.get("featured", ""), 772, ry, 138, 58, f_disp, 12, True, c_acc1, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

        elif st == "chapter_divider_morph":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_finale.png"), 0, 0, 960, 540)
            add_text(s, "cd_ovr", sl.get("overline", ""), 58, 34, 520, 20, f_mono, 10, True, c_acc1)
            add_text(s, "cd_num", sl.get("chapter_num", "01"), 54, 66, 320, 140, f_head, 100, True, c_acc1)
            add_rect(s, "cd_bar", 58, 226, 124, 5, c_acc1)
            add_text(s, "cd_h1", sl.get("headline", ""), 58, 248, 520, 112, f_disp, 34, True, c_white)
            add_text(s, "cd_sub", sl.get("subtitle", ""), 58, 368, 510, 56, f_body, 13.5, False, c_light)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 596, 418, 320, 60)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 16, -32, 4), 612, 64, 290, 364)
            add_rect(s, "cd_line", 58, 448, 844, 1.0, c_border)
            for ti, tc in enumerate(sl.get("toc", [])[:4]):
                is_act = bool(tc.get("active"))
                add_text(s, f"cd_toc_{ti}", tc.get("label", ""), 58 + ti * 214, 464, 200, 28, f_mono, 10, is_act, c_acc1 if is_act else c_light)

        elif st == "full_bleed_visual":
            add_pic(s, "bg_slide", os.path.join(assets_dir, f"s{idx}_visual.png"), 0, 0, 960, 540)
            add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 610, 384, 280, 54)
            add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, -24, 2), 624, 54, 256, 340)
            if sl.get("overline"):
                add_pill(s, "fb_ovr", sl["overline"], 46, 32, 320, 28, c_obs, c_acc1, 9.5, f_mono)
            add_pic(s, "fb_bar", os.path.join(assets_dir, "card_bento_strip.png"), 46, 406, 868, 102)
            add_text(s, "fb_h1", sl.get("headline", ""), 74, 418, 590, 34, f_disp, 21, True, c_white)
            add_text(s, "fb_cap", sl.get("caption", ""), 74, 454, 590, 38, f_body, 12, False, c_light)
            if sl.get("badge"):
                add_pill(s, "fb_badge", sl["badge"], 696, 440, 194, 34, c_acc1, c_white, 10, f_mono)

        else:  # closing_finale
            style = sl.get("style", "editorial_statement")
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_finale.png"), 0, 0, 960, 540)
            if style == "dual_3d" and laptop_glb and os.path.exists(laptop_glb):
                add_pic(s, "f_logo", os.path.join(assets_dir, "logo_badge.png"), 48, 74, 62, 62)
                add_text(s, "f_title", brand.get("name", ""), 48, 148, 312, 54, f_disp, 38, True, c_white)
                add_text(s, "f_tag", sl.get("tagline", ""), 48, 206, 310, 28, f_disp, 15.5, True, c_acc1)
                add_rect(s, "f_div", 48, 246, 56, 2.5, c_acc1)
                add_text(s, "f_sum", sl.get("summary", ""), 48, 264, 310, 112, f_body, 12.5, False, c_light)
                add_pill(s, "f_cta", sl.get("cta", "GET STARTED"), 48, 396, 294, 38, c_acc1, c_white, 11, f_disp)
                add_pic(s, "!!LaptopShadow", os.path.join(assets_dir, "shadow_laptop.png"), 468, 396, 480, 78)
                add_pic(s, "!!Laptop3D", pose_png(laptop_glb, f"s{idx}_lap", 8, -14, 0, res=760), 486, 114, 452, 316)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 356, 442, 200, 46)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}_hero", 6, 20, -3), 364, 194, 210, 260)
            elif style == "executive_cta":
                add_text(s, "fe_ovr", sl.get("overline", ""), 52, 34, 560, 20, f_mono, 10, True, c_acc1)
                add_text(s, "fe_h1", sl.get("headline", ""), 52, 58, 560, 72, f_disp, 34, True, c_white)
                add_text(s, "fe_sum", sl.get("summary", ""), 52, 134, 540, 44, f_body, 13, False, c_light)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_phone.png"), 648, 202, 256, 44)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 14, -24, 3), 664, 24, 226, 196)
                cols = [c_acc1, c_acc2, c_acc3]
                for ai, act_it in enumerate(sl.get("actions", [])[:3]):
                    ax = 52 + ai * 292
                    add_pic(s, f"fe_bg_{ai}", os.path.join(assets_dir, f"card_bento_tile_{ai + 1}.png"), ax, 224, 272, 178)
                    add_text(s, f"fe_num_{ai}", act_it.get("num", ""), ax + 22, 244, 80, 24, f_mono, 13, True, cols[ai])
                    add_text(s, f"fe_ttl_{ai}", act_it.get("title", ""), ax + 22, 272, 228, 28, f_disp, 15, True, c_white)
                    add_text(s, f"fe_dsc_{ai}", act_it.get("desc", ""), ax + 22, 306, 228, 76, f_body, 11.5, False, c_light)
                add_pic(s, "fe_fbar", os.path.join(assets_dir, "card_step.png"), 52, 428, 856, 64)
                add_text(s, "fe_fbrand", f"{brand.get('name', '')}  |  {sl.get('tagline', '')}", 82, 444, 480, 32, f_semi, 13, True, c_white)
                add_pill(s, "fe_cta", sl.get("cta", "APPROVE"), 644, 442, 240, 36, c_acc1, c_white, 10.5, f_mono)
            else:
                add_pic(s, "fs_logo", os.path.join(assets_dir, "logo_badge.png"), 52, 48, 62, 62)
                add_text(s, "fs_ovr", sl.get("overline", ""), 128, 66, 440, 22, f_mono, 9.5, True, c_acc1)
                add_text(s, "fs_h1", sl.get("headline", ""), 52, 128, 500, 112, f_disp, 38, True, c_white)
                add_rect(s, "fs_rule", 52, 252, 120, 2.5, c_acc1)
                add_text(s, "fs_sub", sl.get("summary", ""), 52, 270, 490, 66, f_body, 13.5, False, c_light)
                cols = [c_acc1, c_acc2, c_acc3]
                for mi, mt in enumerate(sl.get("metrics", [])[:3]):
                    add_kpi_box(s, f"fs_mk_{mi}", mt.get("value", ""), mt.get("label", ""), 52 + mi * 166, 352, 152, 72, c_obs, cols[mi], c_light, f_disp, f_semi)
                add_pill(s, "fs_cta", sl.get("cta", "GET STARTED"), 52, 446, 286, 40, c_acc1, c_white, 10.5, f_disp)
                add_pic(s, "!!PhoneShadow", os.path.join(assets_dir, "shadow_laptop.png"), 546, 432, 376, 68)
                add_pic(s, "!!Phone3D", pose_png(hero_glb, f"s{idx}", 12, -24, 2, res=760), 568, 62, 336, 382)

    os.makedirs(os.path.dirname(os.path.abspath(out_pptx)), exist_ok=True)
    prs.save(out_pptx)
    inject_morph_to_pptx(out_pptx)
    print(f"Saved cross-platform Morph deck cleanly -> {out_pptx}")
    return out_pptx


def main():
    parser = argparse.ArgumentParser(description="Build a cross-platform Morph .pptx presentation from rendered studio assets.")
    parser.add_argument("--spec", required=True, help="Path to deck_spec.json")
    parser.add_argument("--assets", required=True, help="Directory generated by render_studio_assets.py")
    parser.add_argument("--out", required=True, help="Output .pptx file path")
    args = parser.parse_args()
    build_crossplatform_pptx(args.spec, args.assets, args.out)


if __name__ == "__main__":
    main()
