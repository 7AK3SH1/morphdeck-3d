#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Cross-Platform True-Morph OpenXML Builder (macOS / Linux / Windows / Headless CI)
for morphdeck-3d v4.0.0.

Implements full pre-staging, post-staging, staggered card kinematics, masked odometer
digit reels, <a:alphaModFix amt="0"/> portal fly-through dissolves, <a:srcRect> curtain
crops, and Z-buffered 2x-SSAA PBR .glb pose rendering so every adjacent slide pair
animates with genuine multi-object PowerPoint Morph continuity.
"""

import argparse
import io
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
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


def add_pic(slide, name: str, path: str, x: float, y: float, w: float, h: float, rot: float = 0.0):
    if not path or not os.path.exists(path):
        return None
    sh = slide.shapes.add_picture(path, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = name
    if rot:
        sh.rotation = float(rot)
    return sh


def add_rect(
    slide,
    name: str,
    x: float,
    y: float,
    w: float,
    h: float,
    col: RGBColor,
    border_col: RGBColor | None = None,
    rounded: bool = False,
    oval: bool = False,
    no_fill: bool = False,
):
    stype = MSO_SHAPE.OVAL if oval else (MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE)
    sh = slide.shapes.add_shape(stype, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = name
    if rounded and not oval and len(sh.adjustments) > 0:
        sh.adjustments[0] = 0.18
    if no_fill:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = col
    if border_col is not None:
        sh.line.color.rgb = border_col
        sh.line.width = Pt(1.2)
    else:
        sh.line.fill.background()
    return sh


def add_text(
    slide,
    name: str,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    font: str,
    size: float,
    bold: bool,
    col: RGBColor,
    align=PP_ALIGN.LEFT,
    v_anchor=MSO_ANCHOR.TOP,
):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(2)
    tf.vertical_anchor = v_anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = "" if text is None else str(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bool(bold)
    r.font.color.rgb = col
    return tb


def add_odometer_strip_xplat(slide, name: str, x: float, y: float, w: float, col: RGBColor, font: str):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(1180))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(0)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for d_i in range(10):
        p = tf.paragraphs[0] if d_i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = Pt(112)
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = str(d_i)
        r.font.name = font
        r.font.size = Pt(90)
        r.font.bold = True
        r.font.color.rgb = col
    return tb


def inject_morph_and_alpha_to_pptx(
    pptx_path: str,
    alpha_zero_targets: set[tuple[int, str]],
    curtain_wipe_targets: set[tuple[int, str]] | None = None,
    duration_ms: int = 1900,
) -> None:
    """
    Inject ECMA-376 <p159:morph option="byObject"/> transitions, <a:alphaModFix amt="0"/>
    optical dissolves on zoomed outgoing portal shapes, and <a:srcRect t="-100000" b="100000"/>
    curtain wipe seeds on incoming dock backgrounds.
    """
    curtain_wipe_targets = curtain_wipe_targets or set()
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
    pic_re = re.compile(r"<p:pic\b.*?</p:pic>", re.DOTALL)

    buf = io.BytesIO()
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r"^ppt/slides/slide(\d+)\.xml$", item.filename)
            if m:
                s_num = int(m.group(1))
                txt = data.decode("utf-8")
                names_for_slide = {nm for (si, nm) in alpha_zero_targets if si == s_num}
                if names_for_slide:
                    def _inject_amf(match: re.Match) -> str:
                        blk = match.group(0)
                        if not any(f'name="{nm}"' in blk for nm in names_for_slide) or "alphaModFix" in blk:
                            return blk
                        blk = re.sub(r"<a:blip\b([^>]*?)/>", r'<a:blip\1><a:alphaModFix amt="0"/></a:blip>', blk, count=1)
                        if "alphaModFix" not in blk:
                            blk = re.sub(r"(<a:blip\b[^>]*>)", r'\1<a:alphaModFix amt="0"/>', blk, count=1)
                        return blk

                    txt = pic_re.sub(_inject_amf, txt)
                wipe_for_slide = {nm for (si, nm) in curtain_wipe_targets if si == s_num}
                if wipe_for_slide:
                    def _inject_wipe(match: re.Match) -> str:
                        blk = match.group(0)
                        if not any(f'name="{nm}"' in blk for nm in wipe_for_slide):
                            return blk
                        if "<a:srcRect" in blk:
                            return re.sub(r"<a:srcRect\b[^>]*/>", '<a:srcRect t="-100000" b="100000"/>', blk, count=1)
                        return blk.replace("<a:stretch>", '<a:srcRect t="-100000" b="100000"/><a:stretch>', 1)

                    txt = pic_re.sub(_inject_wipe, txt)
                if "p159:morph" not in txt and "</p:sld>" in txt:
                    txt = txt.replace("</p:sld>", f"{morph_block}</p:sld>")
                ET.fromstring(txt)
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    with open(pptx_path, "wb") as f:
        f.write(buf.getvalue())


def build_cinema_crossplatform(spec_json: str, assets_dir: str, out_pptx: str) -> str:
    resolved_path = os.path.join(assets_dir, "deck_resolved.json")
    with open(resolved_path if os.path.exists(resolved_path) else spec_json, "r", encoding="utf-8") as f:
        spec = json.load(f)

    pal = spec.get("palette", {})
    fonts = spec.get("fonts", {})
    universe = spec.get("universe", "commercial_studio")
    f_disp = fonts.get("display", "Bahnschrift")
    f_head = fonts.get("heading", f_disp)
    f_body = fonts.get("body", "Segoe UI")
    f_semi = fonts.get("semi", "Segoe UI Semibold")
    f_mono = fonts.get("mono", "Consolas")

    c_surf = hex_rgb(pal.get("surface", "#141824"))
    c_ink = hex_rgb(pal.get("ink", "#FFFFFF"))
    c_muted = hex_rgb(pal.get("ink_muted", "#94A3B8"))
    c_acc1 = hex_rgb(pal.get("accent1", "#E11D48"))
    c_acc2 = hex_rgb(pal.get("accent2", "#38BDF8"))
    c_acc3 = hex_rgb(pal.get("accent3", "#F59E0B"))
    c_bdr = hex_rgb(pal.get("border", "#334155"))
    c_white = RGBColor(255, 255, 255)
    c_slate = RGBColor(203, 213, 225)
    c_gold = RGBColor(251, 191, 36)
    c_obs = RGBColor(10, 12, 18)

    hero_glb = spec.get("resolved_glb", "")
    pose_cache: dict[tuple[float, float, float], str] = {}

    def pose_png(tag: str, rx: float, ry: float, rz: float) -> str:
        if not hero_glb or not os.path.exists(hero_glb):
            return ""
        key = (round(rx, 1), round(ry, 1), round(rz, 1))
        if key not in pose_cache:
            out_p = os.path.join(assets_dir, f"_pose_{tag}.png")
            render_glb_to_image(hero_glb, 680, 680, rx, ry, rz, theme=universe, ssaa=2).save(out_p, "PNG")
            pose_cache[key] = out_p
        return pose_cache[key]

    prs = Presentation()
    prs.slide_width = Pt(960)
    prs.slide_height = Pt(540)
    blank = prs.slide_layouts[6]

    slides = spec.get("slides", [])
    alpha_zero_targets: set[tuple[int, str]] = set()
    curtain_wipe_targets: set[tuple[int, str]] = set()

    for si, sl in enumerate(slides):
        idx = si + 1
        s = prs.slides.add_slide(blank)
        st = sl.get("type", "swiss_bento_morph")
        prev_sl = slides[si - 1] if si > 0 else None
        prev_type = prev_sl.get("type", "") if prev_sl else ""
        next_sl = slides[si + 1] if (si + 1) < len(slides) else None
        next_type = next_sl.get("type", "") if next_sl else ""

        # --- Post-stage outgoing portals/elevators/sheets from prev_sl for smooth zoom/exit dissolve ---
        if prev_type == "real_portal_flythrough" and st != "real_portal_flythrough":
            add_pic(s, "!!portal_bg", os.path.join(assets_dir, "portal_bg.png"), -96, -54, 1152, 648)
            add_pic(s, "!!portal_inner", os.path.join(assets_dir, "portal_inner.png"), -1820, -820, 4400, 2480)
            alpha_zero_targets.add((idx, "!!portal_inner"))
            alpha_zero_targets.add((idx, "!!portal_bg"))
        if prev_type == "vertical_world_elevator" and st != "vertical_world_elevator":
            ebg_p = os.path.join(assets_dir, "elev_bg.png")
            if os.path.exists(ebg_p):
                add_pic(s, "!!elev_bg", ebg_p, 0, -980, 960, 1080)
                alpha_zero_targets.add((idx, "!!elev_bg"))
            efg_p = os.path.join(assets_dir, "elev_fg.png")
            if os.path.exists(efg_p):
                add_pic(s, "!!elev_fg", efg_p, 0, -980, 960, 1080)
                alpha_zero_targets.add((idx, "!!elev_fg"))

        # --- Active slide composition ---
        if st == "newspaper_dive":
            stg = sl.get("stage", "frontpage")
            if stg == "collage":
                add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), -1180, -680, 3456, 1944)
                add_pic(s, "!!col_cutout_L", sl.get("_left_cutout", ""), 32, 94, 336, 446)
                add_pic(s, "!!col_cutout_R", sl.get("_right_cutout", ""), 468, 156, 410, 384)
                add_pic(s, "!!ns_cloud", os.path.join(assets_dir, "ep_paper_cloud.png"), 476, 30, 448, 224)
                add_text(s, "!!ns_h1", sl.get("headline", ""), 514, 68, 376, 66, f_disp, 23, True, c_acc1)
                add_text(s, "!!ns_body", sl.get("body", ""), 514, 140, 376, 82, f_body, 11.5, False, c_ink)
            elif stg == "masthead_spread":
                add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), -680, -320, 2400, 1350)
                add_pic(s, "!!ns_cloud", os.path.join(assets_dir, "ep_paper_cloud.png"), 160, 148, 640, 260)
                add_text(s, "!!ns_ovr", sl.get("overline", ""), 210, 182, 540, 22, f_mono, 9.5, True, c_acc1, PP_ALIGN.CENTER)
                add_text(s, "!!ns_h1", sl.get("headline", ""), 200, 210, 560, 84, f_disp, 30, True, c_ink, PP_ALIGN.CENTER)
                add_text(s, "!!ns_body", sl.get("body", ""), 220, 300, 520, 76, f_body, 12.5, False, c_muted, PP_ALIGN.CENTER)
            else:
                add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), 0, 0, 960, 540)
                add_pic(s, "!!ns_cloud", os.path.join(assets_dir, "ep_paper_cloud.png"), 160, 620, 640, 260)
                add_text(s, "!!ns_h1", sl.get("headline", ""), 200, 660, 560, 84, f_disp, 30, True, c_ink, PP_ALIGN.CENTER)

        elif st == "editorial_print_overlay":
            add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), -280, -180, 1440, 810)
            add_pic(s, "!!ep_drips", os.path.join(assets_dir, "ep_drips.png"), 0, 0, 960, 120)
            cut_p = sl.get("_portrait_cutout") or (pose_png(f"s{idx}_ep", 12, -26, 2) if hero_glb else "")
            add_pic(s, "!!ep_port", cut_p, 48, 86, 344, 444)
            add_pic(s, "!!ep_cloud", os.path.join(assets_dir, "ep_paper_cloud.png"), 396, 118, 528, 356)
            add_text(s, "!!ep_ovr", sl.get("overline", ""), 446, 162, 430, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!ep_h1", sl.get("headline", ""), 446, 190, 430, 74, f_disp, 26, True, c_acc1)
            add_text(s, "!!ep_body", sl.get("body", ""), 446, 272, 430, 138, f_body, 12.5, False, c_ink)

        elif st == "odometer_counter":
            bg_c = hex_rgb(sl.get("bg_hex", "#08090C"))
            add_rect(s, "!!odo_curtain", 0, 0, 960, 540, bg_c)
            target_str = str(sl.get("target_value", "1 000"))
            win_y, step_h = 204.0, 112.0
            tot_w = sum(28 if ch == " " else (64 if ch.isdigit() else 32) for ch in target_str)
            cur_x = 480.0 - (tot_w / 2.0)
            for ci, ch in enumerate(target_str):
                if ch == " ":
                    cur_x += 28
                    continue
                if ch.isdigit():
                    d_val = int(ch)
                    add_odometer_strip_xplat(s, f"!!odo_col_{ci}", cur_x, win_y - d_val * step_h, 64, c_white, f_head)
                    cur_x += 64
                else:
                    add_text(s, f"odo_sep_{ci}", ch, cur_x, win_y, 32, 112, f_head, 90, True, c_white, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                    cur_x += 32
            m_top = next((p for p in (os.path.join(assets_dir, f"s{idx}_odo_mask_top.png"), os.path.join(assets_dir, "odo_mask_top.png")) if os.path.exists(p)), "")
            m_bot = next((p for p in (os.path.join(assets_dir, f"s{idx}_odo_mask_bot.png"), os.path.join(assets_dir, "odo_mask_bot.png")) if os.path.exists(p)), "")
            if m_top and m_bot:
                add_pic(s, "!!odo_mask_top", m_top, -4, -4, 968, 224)
                add_pic(s, "!!odo_mask_bot", m_bot, -4, 304, 968, 240)
            else:
                add_rect(s, "!!odo_mask_top", -4, -4, 968, 224, bg_c)
                add_rect(s, "!!odo_mask_bot", -4, 304, 968, 240, bg_c)
            add_pic(s, "!!odo_sil", sl.get("_silhouette_cutout") or os.path.join(assets_dir, "odo_silhouette.png"), 0, 378, 960, 162)
            add_text(s, "!!odo_ovr", sl.get("overline", ""), 120, 128, 720, 32, f_disp, 20, True, c_acc1, PP_ALIGN.CENTER)
            add_text(s, "!!odo_sub", sl.get("sublabel", ""), 120, 326, 720, 34, f_disp, 22, True, c_white, PP_ALIGN.CENTER)

        elif st == "flyby_wipe":
            bg_p = sl.get("_bg_image") or os.path.join(assets_dir, "bg_base.png")
            add_pic(s, "!!StageBg", bg_p, 0, 0, 960, 540)
            add_pic(s, "!!flyby_cloud", os.path.join(assets_dir, "flyby_cloud.png"), 980, -220, 760, 440)
            if sl.get("_flyby_cutout"):
                add_pic(s, "!!flyby_obj", sl["_flyby_cutout"], 1140, -300, 440, 280, rot=-14)
            txt_c = c_white if sl.get("_bg_image") else c_ink
            ovr_c = c_white if sl.get("_bg_image") else c_acc1
            add_text(s, "!!fb_ovr", sl.get("overline", ""), 56, 56, 480, 20, f_mono, 9.5, True, ovr_c)
            add_text(s, "!!fb_h1", sl.get("headline", ""), 56, 84, 490, 92, f_disp, 34, True, txt_c)
            add_text(s, "!!fb_body", sl.get("body", ""), 56, 188, 460, 108, f_body, 13, False, txt_c)
            if hero_glb and sl.get("use_hero_3d") is True:
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, -24, 2), 534, 114, 388, 336)

        elif st == "cutout_map_stage":
            bg_p = sl.get("_bg_image") or os.path.join(assets_dir, "bg_base.png")
            add_pic(s, "!!StageBg", bg_p, 0, 0, 960, 540)
            if sl.get("_map_image"):
                add_pic(s, "!!cm_map", sl["_map_image"], 42, 54, 520, 432)
            # Open-canvas editorial split (no cards inside cards!)
            add_rect(s, "!!cm_div", 588, 68, 1.5, 404, c_bdr)
            add_text(s, "!!cm_ovr", sl.get("overline", ""), 614, 72, 294, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!cm_h1", sl.get("headline", ""), 614, 100, 294, 84, f_disp, 30, True, c_acc1)
            add_text(s, "!!cm_body", sl.get("body", ""), 614, 196, 294, 150, f_body, 12.5, False, c_ink)
            for ki, k in enumerate((sl.get("kpis") or [])[:2]):
                kx = 614 + ki * 150
                add_rect(s, f"!!cm_kline_{ki}", kx, 368, 132, 2, c_acc1)
                add_text(s, f"!!cm_kv_{ki}", k.get("value", ""), kx, 378, 132, 40, f_disp, 26, True, c_acc1)
                add_text(s, f"!!cm_kl_{ki}", k.get("label", ""), kx, 420, 132, 22, f_mono, 9, False, c_muted)

        elif st == "real_portal_flythrough":
            stg = sl.get("stage", "exterior")
            if stg == "interior":
                add_pic(s, "!!portal_bg", os.path.join(assets_dir, "portal_bg.png"), -48, -27, 1056, 594)
                add_pic(s, "!!portal_outer", os.path.join(assets_dir, "portal_outer.png"), -1900, -720, 4620, 2600)
                alpha_zero_targets.add((idx, "!!portal_outer"))
                add_pic(s, "!!portal_inner", os.path.join(assets_dir, "portal_inner.png"), 0, 0, 960, 540)
                add_text(s, "!!pt_ext_title", prev_sl.get("headline", "") if prev_sl else "", 320, 640, 588, 104, f_disp, 38, True, c_white, PP_ALIGN.RIGHT)
                add_text(s, "!!pt_int_title", sl.get("headline", ""), 116, 174, 440, 132, f_disp, 36, True, c_white)
                add_text(s, "!!pt_int_body", sl.get("body", ""), 116, 314, 420, 90, f_body, 13, False, c_white)
                if next_type == "real_portal_flythrough":
                    add_text(s, "!!pt_hor_title", next_sl.get("headline", ""), -620, 78, 480, 128, f_disp, 40, True, c_white)
                    add_text(s, "!!pt_hor_body", next_sl.get("body", ""), -720, 212, 440, 96, f_body, 13, False, c_white)
                if hero_glb and sl.get("use_hero_3d", False):
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 10, -22, 1), 572, 196, 326, 296)
            elif stg == "horizon":
                add_pic(s, "!!portal_bg", os.path.join(assets_dir, "portal_bg.png"), 0, 0, 960, 540)
                add_pic(s, "!!portal_inner", os.path.join(assets_dir, "portal_inner.png"), -1820, -820, 4400, 2480)
                alpha_zero_targets.add((idx, "!!portal_inner"))
                add_text(s, "!!pt_int_title", prev_sl.get("headline", "") if prev_sl else "", 116, -280, 440, 132, f_disp, 36, True, c_white)
                add_text(s, "!!pt_int_body", prev_sl.get("body", "") if prev_sl else "", 116, -240, 420, 90, f_body, 13, False, c_white)
                add_text(s, "!!pt_hor_title", sl.get("headline", ""), 58, 78, 480, 128, f_disp, 40, True, c_white)
                add_text(s, "!!pt_hor_body", sl.get("body", ""), 58, 212, 440, 96, f_body, 13, False, c_white)
            else:
                add_pic(s, "!!portal_bg", os.path.join(assets_dir, "portal_bg.png"), 0, 0, 960, 540)
                add_pic(s, "!!portal_inner", os.path.join(assets_dir, "portal_inner.png"), 270, 124, 420, 196)
                add_pic(s, "!!portal_outer", os.path.join(assets_dir, "portal_outer.png"), 0, 0, 960, 540)
                add_text(s, "!!pt_ext_title", sl.get("headline", ""), 290, 388, 614, 118, f_disp, 38, True, c_white, PP_ALIGN.RIGHT)
                if next_type == "real_portal_flythrough":
                    add_text(s, "!!pt_int_title", next_sl.get("headline", ""), -560, 174, 440, 132, f_disp, 36, True, c_white)
                    add_text(s, "!!pt_int_body", next_sl.get("body", ""), -680, 314, 420, 90, f_body, 13, False, c_white)

        elif st == "parchment_3d_split":
            add_pic(s, "!!StageBg", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            low_p = os.path.join(assets_dir, f"s{idx}_split_lower.png")
            if os.path.exists(low_p):
                add_pic(s, "!!SplitLower", low_p, 0, 272, 960, 268)
            add_text(s, "!!GiantWord", sl.get("giant_word", "STUDIO"), 30, 18, 900, 168, f_disp, 118, True, c_acc1, PP_ALIGN.CENTER)
            add_text(s, "!!s2_pl", sl.get("left_label", ""), 42, 186, 248, 18, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!s2_pt", sl.get("left_text", ""), 42, 206, 248, 62, f_semi, 11.5, False, c_ink)
            add_text(s, "!!s2_sl", sl.get("right_label", ""), 670, 186, 248, 18, f_mono, 9.5, True, c_acc1, PP_ALIGN.RIGHT)
            add_text(s, "!!s2_st", sl.get("right_text", ""), 670, 206, 248, 62, f_semi, 11.5, False, c_ink, PP_ALIGN.RIGHT)
            if hero_glb:
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 334, 468, 292, 54)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 8, -16, 0), 318, 96, 324, 396)

        elif st == "vertical_world_elevator":
            stop = sl.get("elevator_stop", "upper_sky")
            ebg_p = os.path.join(assets_dir, "elev_bg.png")
            if not os.path.exists(ebg_p):
                ebg_p = os.path.join(assets_dir, "bg_base.png")
            efg_p = os.path.join(assets_dir, "elev_fg.png")
            gw_txt = sl.get("giant_word", "HORIZON")
            show_hero = bool(hero_glb and sl.get("use_hero_3d") is True)
            if stop == "lower_ground":
                add_pic(s, "!!elev_bg", ebg_p, 0, -520, 960, 1080)
                word_y = 286 if show_hero else 118
                add_text(s, "!!elev_word", gw_txt, 40, word_y, 880, 170, f_head, 128, True, c_white, PP_ALIGN.CENTER)
                if os.path.exists(efg_p):
                    add_pic(s, "!!elev_fg", efg_p, 0, -520, 960, 1080)
                add_text(s, "!!ev_up_h1", prev_sl.get("headline", "") if prev_sl else "", 56, -240, 440, 72, f_disp, 30, True, c_white)
                cards = (sl.get("cards") or [])[:3]
                if show_hero:
                    add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 48, 456, 348, 56)
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 14, 26, -2), 44, 140, 356, 328)
                    add_text(s, "!!ev_h2", sl.get("headline", ""), 444, 52, 468, 64, f_disp, 26, True, c_white)
                    for ci, crd in enumerate(cards):
                        cy = 142 + ci * 118
                        rc_p = os.path.join(assets_dir, f"elev_rcard_{ci + 1}.png")
                        if not os.path.exists(rc_p):
                            rc_p = os.path.join(assets_dir, f"card_spec_{ci + 1}.png")
                        add_pic(s, f"!!ev_rcard_{ci}", rc_p, 444, cy, 468, 102)
                        add_text(s, f"ev_rv_{ci}", crd.get("value", ""), 464, cy + 16, 124, 66, f_disp, 28, True, c_gold, PP_ALIGN.CENTER)
                        add_text(s, f"ev_rt_{ci}", crd.get("title", ""), 596, cy + 16, 294, 26, f_disp, 14, True, c_white)
                        add_text(s, f"ev_rd_{ci}", crd.get("desc", ""), 596, cy + 44, 294, 44, f_body, 11, False, c_slate)
                else:
                    add_text(s, "!!ev_h2", sl.get("headline", ""), 56, 36, 520, 64, f_disp, 26, True, c_white)
                    for ci, crd in enumerate(cards):
                        cx_c, cy_c = 48 + ci * 298, 348
                        rc_p = os.path.join(assets_dir, f"elev_rcard_h_{ci + 1}.png")
                        if not os.path.exists(rc_p):
                            rc_p = os.path.join(assets_dir, f"card_spec_{ci + 1}.png")
                        add_pic(s, f"!!ev_rcard_{ci}", rc_p, cx_c, cy_c, 268, 140)
                        add_text(s, f"ev_rv_{ci}", crd.get("value", ""), cx_c + 18, cy_c + 14, 232, 42, f_disp, 28, True, c_gold)
                        add_text(s, f"ev_rt_{ci}", crd.get("title", ""), cx_c + 18, cy_c + 58, 232, 24, f_disp, 13.5, True, c_white)
                        add_text(s, f"ev_rd_{ci}", crd.get("desc", ""), cx_c + 18, cy_c + 86, 232, 42, f_body, 10.5, False, c_slate)
            else:
                add_pic(s, "!!elev_bg", ebg_p, 0, 0, 960, 1080)
                add_text(s, "!!elev_word", gw_txt, 40, 68, 880, 168, f_head, 132, True, c_white, PP_ALIGN.CENTER)
                if os.path.exists(efg_p):
                    add_pic(s, "!!elev_fg", efg_p, 0, 0, 960, 1080)
                add_text(s, "!!ev_up_ovr", sl.get("overline", ""), 56, 334, 420, 20, f_mono, 9.5, True, c_gold)
                add_text(s, "!!ev_up_h1", sl.get("headline", ""), 56, 358, 460, 72, f_disp, 30, True, c_white)
                add_text(s, "!!ev_up_body", sl.get("body", ""), 56, 434, 440, 68, f_body, 12.5, False, c_white)
                if next_type == "vertical_world_elevator":
                    next_show_hero = bool(hero_glb and next_sl and next_sl.get("use_hero_3d") is True)
                    for ci in range(3):
                        if next_show_hero:
                            rc_n = os.path.join(assets_dir, f"elev_rcard_{ci + 1}.png")
                            if not os.path.exists(rc_n):
                                rc_n = os.path.join(assets_dir, f"card_spec_{ci + 1}.png")
                            add_pic(s, f"!!ev_rcard_{ci}", rc_n, 1060 + ci * 160, 142 + ci * 118, 468, 102)
                        else:
                            rc_n = os.path.join(assets_dir, f"elev_rcard_h_{ci + 1}.png")
                            if not os.path.exists(rc_n):
                                rc_n = os.path.join(assets_dir, f"card_spec_{ci + 1}.png")
                            add_pic(s, f"!!ev_rcard_{ci}", rc_n, 48 + ci * 298, 640 + ci * 80, 268, 140)
                    if next_show_hero:
                        add_pic(s, "!!Hero3D", pose_png(f"s{idx}_pre", 10, -18, 2), -380, 140, 356, 328)

        elif st == "refracted_spotlight_dock":
            dock_idx = int(sl.get("dock_index", 0))
            dock_items = sl.get("dock") or []
            n_dock = 3 if len(dock_items) <= 3 else 4
            if n_dock == 3:
                dock_xs, glider_xs, glider_w = [62, 342, 622], [46, 326, 606], 286
                thumb_w, thumb_h, txt_off, txt_w, line_w = 92, 52, 104, 154, 254
            else:
                dock_xs, glider_xs, glider_w = [64, 280, 496, 712], [50, 266, 482, 698], 212
                thumb_w, thumb_h, txt_off, txt_w, line_w = 76, 44, 86, 114, 196
            bg_p = os.path.join(assets_dir, f"s{idx}_bg.png")
            add_pic(s, f"!!sp_bg_{idx}", bg_p if os.path.exists(bg_p) else os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            if prev_type == "refracted_spotlight_dock" and prev_sl:
                p_sp = prev_sl.get("spotlight", {})
                add_text(s, f"!!sp_tag_{idx - 1}", p_sp.get("tag", ""), 56, -280, 520, 20, f_mono, 9.5, True, c_gold)
                add_text(s, f"!!sp_title_{idx - 1}", p_sp.get("title", ""), 56, -240, 540, 88, f_disp, 36, True, c_white)
                add_text(s, f"!!sp_sub_{idx - 1}", p_sp.get("subtitle", ""), 56, -180, 520, 26, f_semi, 13, False, c_slate)
                add_text(s, f"!!sp_body_{idx - 1}", p_sp.get("body", ""), 56, -140, 490, 104, f_body, 13, False, c_white)
            if next_type == "refracted_spotlight_dock" and next_sl:
                nbg_p = os.path.join(assets_dir, f"s{idx + 1}_bg.png")
                add_pic(s, f"!!sp_bg_{idx + 1}", nbg_p if os.path.exists(nbg_p) else os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
                curtain_wipe_targets.add((idx, f"!!sp_bg_{idx + 1}"))
                nx_sp = next_sl.get("spotlight", {})
                add_text(s, f"!!sp_tag_{idx + 1}", nx_sp.get("tag", ""), -460, 48, 520, 20, f_mono, 9.5, True, c_gold)
                add_text(s, f"!!sp_title_{idx + 1}", nx_sp.get("title", ""), -580, 74, 540, 88, f_disp, 36, True, c_white)
                add_text(s, f"!!sp_sub_{idx + 1}", nx_sp.get("subtitle", ""), -700, 168, 520, 26, f_semi, 13, False, c_slate)
                add_text(s, f"!!sp_body_{idx + 1}", nx_sp.get("body", ""), -820, 202, 490, 104, f_body, 13, False, c_white)
            sp = sl.get("spotlight", {})
            add_text(s, f"!!sp_tag_{idx}", sp.get("tag", ""), 56, 48, 520, 20, f_mono, 9.5, True, c_gold)
            add_text(s, f"!!sp_title_{idx}", sp.get("title", ""), 56, 74, 540, 88, f_disp, 36, True, c_white)
            add_text(s, f"!!sp_sub_{idx}", sp.get("subtitle", ""), 56, 168, 520, 26, f_semi, 13, False, c_slate)
            add_text(s, f"!!sp_body_{idx}", sp.get("body", ""), 56, 202, 490, 104, f_body, 13, False, c_white)
            if hero_glb and sl.get("use_hero_3d", False):
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 10, -20, 2), 624, 58, 296, 254)

            dock_p = os.path.join(assets_dir, f"s{idx}_dock_refract.png")
            add_pic(s, "!!GlassDock", dock_p if os.path.exists(dock_p) else os.path.join(assets_dir, "dock_bg.png"), 40, 386, 880, 136)
            add_pic(s, "!!DockGlider", os.path.join(assets_dir, "dock_glider.png"), glider_xs[max(0, min(n_dock - 1, dock_idx))], 388, glider_w, 132)
            for di in range(n_dock):
                it = dock_items[di] if di < len(dock_items) else {"title": f"Item {di + 1}", "subtitle": ""}
                dn, dx = di + 1, dock_xs[di]
                add_pic(s, f"!!DockThumb{dn}", os.path.join(assets_dir, f"thumb_{dn}.png"), dx, 416, thumb_w, thumb_h)
                add_text(s, f"!!DockLbl{dn}", it.get("title", ""), dx + txt_off, 416, txt_w, 24, f_disp, 13.5, True, c_white)
                add_text(s, f"!!DockSub{dn}", it.get("subtitle", ""), dx + txt_off, 442, txt_w, 26, f_body, 10.5, False, c_slate)
                add_rect(s, f"!!DockLine{dn}", dx, 484, line_w, 1.2, c_white if di == (dock_idx % n_dock) else c_bdr)
            for pi in range(3):
                is_act = pi == (dock_idx % 3)
                sz = 10.5 if is_act else 7.0
                px = 928.0 - (sz / 2.0)
                add_rect(s, f"!!sp_ring_{pi}", px, 210 + pi * 20, sz, sz, c_white, border_col=c_white, oval=True, no_fill=not is_act)

        elif st == "poster_depth_reveal":
            if next_type == "picture_frame_pullback":
                wall_n = next_sl.get("_wall_image") or os.path.join(assets_dir, "bg_alt.png")
                add_pic(s, "!!wall_room", wall_n, -1632, -918, 4224, 2376)
            add_pic(s, "!!frame_art", sl.get("_poster_bg_image") or os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            add_rect(s, "!!poster_band", 348, 0, 264, 540, c_acc1)
            gw_txt = str(sl.get("giant_word", "2026"))
            gw_sz = 184 if len(gw_txt) <= 5 else 152
            add_text(s, "!!poster_word", gw_txt, 30, 88, 900, 220, f_head, gw_sz, True, c_white, PP_ALIGN.CENTER)
            if sl.get("_hero_cutout"):
                add_pic(s, "!!Hero3D", sl["_hero_cutout"], 262, 166, 436, 352)
            elif hero_glb and sl.get("use_hero_3d", True):
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, -22, 2), 268, 170, 424, 344)
            if next_type == "picture_frame_pullback" and next_sl.get("_hollow_frame"):
                add_pic(s, "!!hollow_frame", next_sl["_hollow_frame"], -240, -180, 1440, 900)

        elif st == "picture_frame_pullback":
            add_pic(s, "!!wall_room", sl.get("_wall_image") or os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            fx = float(sl.get("frame_x", 508))
            fy = float(sl.get("frame_y", 78))
            fw = float(sl.get("frame_w", 340))
            fh = float(sl.get("frame_h", 192))
            prev_poster_bg = (prev_sl.get("_poster_bg_image") if prev_sl else "") or sl.get("_framed_art") or os.path.join(assets_dir, "bg_alt.png")
            add_pic(s, "!!frame_art", prev_poster_bg, fx, fy, fw, fh)
            add_rect(s, "!!poster_band", fx + fw * 0.362, fy, fw * 0.276, fh, c_acc1)
            gw_prev = prev_sl.get("giant_word", "1945") if prev_sl else "1945"
            add_text(s, "!!poster_word", gw_prev, fx + fw * 0.06, fy + fh * 0.18, fw * 0.88, fh * 0.42, f_head, 40, True, c_white, PP_ALIGN.CENTER)
            if hero_glb:
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, -22, 2), fx + fw * 0.23, fy + fh * 0.28, fw * 0.54, fh * 0.64)
            if sl.get("_hollow_frame"):
                add_pic(s, "!!hollow_frame", sl["_hollow_frame"], fx - fw * 0.0755, fy - fh * 0.152, fw / 0.86875, fh / 0.76667)
            add_text(s, "!!pf_ovr", sl.get("overline", ""), 62, 148, 380, 20, f_mono, 9.5, True, c_gold)
            add_text(s, "!!pf_h1", sl.get("headline", ""), 62, 178, 396, 102, f_disp, 34, True, c_white)
            add_text(s, "!!pf_sub", sl.get("subtitle", ""), 62, 290, 380, 68, f_body, 13, False, c_slate)

        elif st == "rotational_drift_3d":
            add_pic(s, "!!StageBg", sl.get("_bg_image") or os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "!!md_ovr", sl.get("overline", ""), 48, 26, 520, 18, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!md_h1", sl.get("headline", ""), 48, 48, 470, 74, f_disp, 26, True, c_ink)
            if hero_glb:
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 488, 426, 418, 58)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 18, -34, -4), 484, 72, 426, 360)
            c_cols = [c_acc1, c_acc2, c_acc3]
            for ci, co in enumerate((sl.get("callouts") or [])[:3]):
                cy = 144 + ci * 116
                ac = c_cols[ci % 3]
                add_pic(s, f"!!md_card_{ci}", os.path.join(assets_dir, f"card_spec_{ci + 1}.png"), 48, cy, 384, 96)
                add_text(s, f"md_t_{ci}", co.get("title", ""), 68, cy + 14, 344, 22, f_disp, 13.5, True, ac)
                add_text(s, f"md_d_{ci}", co.get("desc", ""), 68, cy + 40, 344, 44, f_body, 11, False, c_ink)
                add_rect(s, f"!!md_line_{ci}", 432, cy + 47, 54, 1.5, ac)

        elif st == "swiss_bento_morph":
            add_pic(s, "!!StageBg", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "!!kb_ovr", sl.get("overline", ""), 48, 24, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!kb_h1", sl.get("headline", ""), 48, 46, 660, 54, f_disp, 25, True, c_ink)
            hero_t = sl.get("hero_tile") or {}
            add_pic(s, "!!kb_hero", os.path.join(assets_dir, "card_bento_strip.png"), 46, 106, 588, 154)
            add_text(s, "kb_htag", hero_t.get("label", ""), 74, 120, 180, 18, f_mono, 9.5, True, c_acc1)
            add_text(s, "kb_hval", hero_t.get("value", ""), 74, 144, 180, 84, f_disp, 44, True, c_ink)
            add_text(s, "kb_hsub", hero_t.get("title", ""), 280, 124, 334, 30, f_disp, 16.5, True, c_ink)
            add_text(s, "kb_hdsc", hero_t.get("desc", ""), 280, 158, 334, 62, f_body, 11.5, False, c_muted)
            if hero_glb:
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 656, 240, 254, 42)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 14, -28, 3), 650, 44, 266, 218)
            b_cols = [c_acc1, c_acc2, c_acc3]
            for ti, t in enumerate((sl.get("tiles") or [])[:3]):
                tx = 46 + ti * 296
                add_pic(s, f"!!kb_tbg_{ti}", os.path.join(assets_dir, f"card_bento_tile_{ti + 1}.png"), tx, 276, 276, 222)
                add_text(s, f"kb_tl_{ti}", t.get("label", ""), tx + 22, 298, 232, 20, f_mono, 9.5, True, b_cols[ti % 3])
                add_text(s, f"kb_tv_{ti}", t.get("value", ""), tx + 22, 324, 232, 62, f_disp, 38, True, c_ink)
                add_text(s, f"kb_td_{ti}", t.get("desc", ""), tx + 22, 396, 232, 82, f_body, 11.5, False, c_muted)

        elif st == "timeline_rail_morph":
            add_pic(s, "!!StageBg", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            stages = (sl.get("stages") or [])[:4]
            n_st = max(1, len(stages))
            act_i = max(0, min(n_st - 1, int(sl.get("active_index", 0))))
            add_text(s, "!!tl_ovr", sl.get("overline", ""), 52, 24, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!tl_h1", sl.get("headline", ""), 52, 46, 760, 54, f_disp, 25, True, c_ink)
            add_rect(s, "!!tl_rail", 76, 144, 808, 3, c_bdr)
            centers = [480.0] if n_st == 1 else [146.0 + 668.0 * (ti / (n_st - 1)) for ti in range(n_st)]
            add_rect(s, "!!tl_progress", 76, 144, max(24.0, centers[act_i] - 76.0), 3, c_acc1)
            for ti, stg in enumerate(stages):
                cx = centers[ti]
                nw = 172.0 if ti == act_i else 148.0
                add_rect(s, f"!!tl_node_{ti}", cx - nw / 2.0, 128, nw, 34, c_acc1 if ti == act_i else c_surf, border_col=None if ti == act_i else c_bdr, rounded=True)
                add_text(s, f"tl_nt_{ti}", f"{stg.get('code', '')}  {stg.get('label', '')}", cx - nw / 2.0 + 8, 134, nw - 16, 22, f_mono, 9.5, True, c_white if ti == act_i else c_muted, PP_ALIGN.CENTER)
            if hero_glb:
                rot_ys = [22, -24, 38, -16]
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 32, 434, 264, 48)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, rot_ys[act_i % 4], 2), 28, 208, 272, 240)
            if stages:
                cur = stages[act_i]
                add_pic(s, "!!tl_drawer", os.path.join(assets_dir, "card_panel.png"), 316, 204, 596, 292)
                add_text(s, "tl_m", cur.get("metric", ""), 344, 240, 176, 76, f_disp, 42, True, c_acc1)
                add_text(s, "tl_ml", cur.get("metric_label", ""), 344, 322, 176, 36, f_semi, 11, False, c_ink)
                add_text(s, "tl_sh", cur.get("title", ""), 544, 226, 346, 52, f_disp, 18.5, True, c_ink)
                add_text(s, "tl_su", cur.get("summary", ""), 544, 284, 346, 110, f_body, 12, False, c_muted)

        elif st == "chart_infographic":
            add_pic(s, "!!StageBg", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "!!ch_ovr", sl.get("overline", ""), 46, 24, 600, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!ch_h1", sl.get("headline", ""), 46, 46, 600, 54, f_disp, 25, True, c_ink)
            add_pic(s, "!!ch_img", os.path.join(assets_dir, f"s{idx}_chart.png"), 46, 110, 584, 384)
            if hero_glb:
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 656, 242, 250, 44)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 16, 28, -2), 650, 38, 262, 222)
            hl = sl.get("highlight") or {}
            add_pic(s, "!!ch_hlbg", os.path.join(assets_dir, "card_spec_1.png"), 648, 278, 268, 104)
            add_text(s, "ch_hlv", f"{hl.get('value', '')}  {hl.get('label', '')}", 670, 290, 232, 30, f_disp, 18, True, c_acc1)
            add_text(s, "ch_hld", hl.get("desc", ""), 670, 324, 232, 48, f_body, 10.5, False, c_ink)

        elif st == "architecture_blueprint":
            add_pic(s, "!!StageBg", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "!!ar_ovr", sl.get("overline", ""), 42, 22, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!ar_h1", sl.get("headline", ""), 42, 46, 630, 52, f_disp, 24, True, c_ink)
            add_pic(s, "!!ar_img", os.path.join(assets_dir, f"s{idx}_arch.png"), 42, 104, 624, 398)
            if hero_glb:
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 690, 238, 224, 40)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 22, -34, -3), 686, 36, 232, 216)
            for pi, pl in enumerate((sl.get("pillars") or [])[:2]):
                sy = 276 + pi * 116
                add_pic(s, f"!!ar_cbg_{pi}", os.path.join(assets_dir, f"card_spec_{pi + 1}.png"), 682, sy, 236, 106)
                add_text(s, f"ar_ct_{pi}", f"{pl.get('tag', '')}  //  {pl.get('title', '')}", 702, sy + 12, 204, 22, f_disp, 12, True, c_acc1)
                add_text(s, f"ar_cd_{pi}", pl.get("desc", ""), 702, sy + 38, 204, 58, f_body, 10.5, False, c_ink)

        else:  # closing_finale
            bg_p = sl.get("_bg_image") or os.path.join(assets_dir, "bg_base.png")
            add_pic(s, "!!StageBg", bg_p, 0, 0, 960, 540)
            txt_c = c_white if sl.get("_bg_image") else c_ink
            mut_c = c_white if sl.get("_bg_image") else c_muted
            add_text(s, "!!fs_ovr", sl.get("overline", ""), 54, 58, 460, 22, f_mono, 9.5, True, c_acc1)
            add_text(s, "!!fs_h1", sl.get("headline", ""), 54, 92, 496, 118, f_disp, 36, True, txt_c)
            add_rect(s, "!!fs_rule", 54, 222, 96, 2.5, c_acc1)
            add_text(s, "!!fs_sub", sl.get("summary", ""), 54, 242, 480, 76, f_body, 13.5, False, mut_c)
            for mi, m in enumerate((sl.get("metrics") or [])[:3]):
                mx = 54 + mi * 164
                add_rect(s, f"!!fs_mline_{mi}", mx, 344, 142, 2, c_acc1)
                add_text(s, f"!!fs_mv_{mi}", m.get("value", ""), mx, 354, 142, 42, f_disp, 28, True, c_acc1)
                add_text(s, f"!!fs_ml_{mi}", m.get("label", ""), mx, 398, 142, 22, f_mono, 9, False, mut_c)
            if hero_glb:
                add_pic(s, "!!HeroShadow", os.path.join(assets_dir, "shadow_hero.png"), 510, 432, 388, 56)
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 10, -20, 1), 498, 92, 412, 352)

        # --- Universal Pre-Staging Bridge for the NEXT slide (guarantees >= 2 moving !! actors on every pair!) ---
        if next_sl is not None:
            if next_type == "editorial_print_overlay" and st != "editorial_print_overlay":
                add_pic(s, "!!ep_drips", os.path.join(assets_dir, "ep_drips.png"), 0, -180, 960, 120)
                cut_n = next_sl.get("_portrait_cutout") or (pose_png(f"s{idx + 1}_ep", 12, -26, 2) if hero_glb else "")
                add_pic(s, "!!ep_port", cut_n, -390, 86, 344, 444)
                add_pic(s, "!!ep_cloud", os.path.join(assets_dir, "ep_paper_cloud.png"), 1020, 118, 528, 356)
            elif next_type == "odometer_counter":
                next_bg = hex_rgb(next_sl.get("bg_hex", "#08090C"))
                add_rect(s, "!!odo_curtain", -980, 0, 960, 540, next_bg)
                next_t = str(next_sl.get("target_value", "1 000"))
                tw_n = sum(28 if ch == " " else (64 if ch.isdigit() else 32) for ch in next_t)
                cx_n = -980.0 + (480.0 - tw_n / 2.0)
                for ci, ch in enumerate(next_t):
                    if ch == " ":
                        cx_n += 28
                    elif ch.isdigit():
                        add_odometer_strip_xplat(s, f"!!odo_col_{ci}", cx_n, 204.0 + ((ci % 3) - 9) * 112.0, 64, c_white, f_head)
                        cx_n += 64
                    else:
                        cx_n += 32
                m_top_n = next((p for p in (os.path.join(assets_dir, f"s{idx + 1}_odo_mask_top.png"), os.path.join(assets_dir, "odo_mask_top.png")) if os.path.exists(p)), "")
                m_bot_n = next((p for p in (os.path.join(assets_dir, f"s{idx + 1}_odo_mask_bot.png"), os.path.join(assets_dir, "odo_mask_bot.png")) if os.path.exists(p)), "")
                if m_top_n and m_bot_n:
                    add_pic(s, "!!odo_mask_top", m_top_n, -984, -4, 968, 224)
                    add_pic(s, "!!odo_mask_bot", m_bot_n, -984, 304, 968, 240)
                else:
                    add_rect(s, "!!odo_mask_top", -984, -4, 968, 224, next_bg)
                    add_rect(s, "!!odo_mask_bot", -984, 304, 968, 240, next_bg)
                add_pic(s, "!!odo_sil", os.path.join(assets_dir, "odo_silhouette.png"), -980, 378, 960, 162)
            elif next_type == "flyby_wipe":
                add_pic(s, "!!flyby_cloud", os.path.join(assets_dir, "flyby_cloud.png"), -860, 320, 760, 440)
                if next_sl.get("_flyby_cutout"):
                    add_pic(s, "!!flyby_obj", next_sl["_flyby_cutout"], -480, 400, 440, 280, rot=-8)
                if hero_glb and next_sl.get("use_hero_3d") is True and not any(sh.name == "!!Hero3D" for sh in s.shapes):
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx + 1}_pre", 8, 18, -4), 1040, 114, 388, 336)
            elif next_type == "cutout_map_stage":
                if next_sl.get("_map_image"):
                    add_pic(s, "!!cm_map", next_sl["_map_image"], -560, 54, 520, 432)
                add_rect(s, "!!cm_div", 1020, 68, 1.5, 404, c_bdr)
                add_text(s, "!!cm_h1", next_sl.get("headline", ""), 1040, 100, 294, 84, f_disp, 30, True, c_acc1)
            elif next_type == "poster_depth_reveal":
                add_rect(s, "!!poster_band", 348, -560, 264, 540, c_acc1)
                ngw = str(next_sl.get("giant_word", "2026"))
                add_text(s, "!!poster_word", ngw, 30, -240, 900, 220, f_head, 184 if len(ngw) <= 5 else 152, True, c_white, PP_ALIGN.CENTER)
                if hero_glb and not any(sh.name == "!!Hero3D" for sh in s.shapes):
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx + 1}_pre", 16, 26, -2), 268, 580, 424, 344)
            elif next_type == "parchment_3d_split" and st != "parchment_3d_split":
                low_np = os.path.join(assets_dir, f"s{idx + 1}_split_lower.png")
                if os.path.exists(low_np):
                    add_pic(s, "!!SplitLower", low_np, 0, 580, 960, 268)
                add_text(s, "!!GiantWord", next_sl.get("giant_word", "STUDIO"), 30, -210, 900, 168, f_disp, 118, True, c_acc1, PP_ALIGN.CENTER)
                if hero_glb and not any(sh.name == "!!Hero3D" for sh in s.shapes):
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx + 1}_pre", 14, -28, 2), 318, 580, 324, 396)
            elif next_type == "vertical_world_elevator" and st != "vertical_world_elevator":
                ebg_np = os.path.join(assets_dir, "elev_bg.png")
                if os.path.exists(ebg_np):
                    add_pic(s, "!!elev_bg", ebg_np, 0, 580, 960, 1080)
                add_text(s, "!!elev_word", next_sl.get("giant_word", "HORIZON"), 40, 760, 880, 168, f_head, 132, True, c_white, PP_ALIGN.CENTER)
                efg_np = os.path.join(assets_dir, "elev_fg.png")
                if os.path.exists(efg_np):
                    add_pic(s, "!!elev_fg", efg_np, 0, 580, 960, 1080)
            elif next_type == "refracted_spotlight_dock" and st != "refracted_spotlight_dock":
                dock_np = os.path.join(assets_dir, f"s{idx + 1}_dock_refract.png")
                n_cnt = 3 if len(next_sl.get("dock") or []) <= 3 else 4
                add_pic(s, "!!GlassDock", dock_np if os.path.exists(dock_np) else os.path.join(assets_dir, "dock_bg.png"), 40, 620, 880, 136)
                if n_cnt == 3:
                    add_pic(s, "!!DockGlider", os.path.join(assets_dir, "dock_glider.png"), 46, 680, 286, 132)
                    for dn in range(1, 4):
                        tp = os.path.join(assets_dir, f"thumb_{dn}.png")
                        if os.path.exists(tp):
                            add_pic(s, f"!!DockThumb{dn}", tp, 62 + (dn - 1) * 280, 650 + dn * 55, 92, 52)
                        add_rect(s, f"!!DockLine{dn}", 62 + (dn - 1) * 280, 760, 254, 1.2, c_bdr)
                else:
                    add_pic(s, "!!DockGlider", os.path.join(assets_dir, "dock_glider.png"), 50, 680, 212, 132)
                    for dn in range(1, 5):
                        tp = os.path.join(assets_dir, f"thumb_{dn}.png")
                        if os.path.exists(tp):
                            add_pic(s, f"!!DockThumb{dn}", tp, 64 + (dn - 1) * 216, 650 + dn * 55, 76, 44)
                        add_rect(s, f"!!DockLine{dn}", 64 + (dn - 1) * 216, 760, 196, 1.2, c_bdr)
            elif next_type == "rotational_drift_3d" and st != "rotational_drift_3d":
                for ci in range(3):
                    add_pic(s, f"!!md_card_{ci}", os.path.join(assets_dir, f"card_spec_{ci + 1}.png"), -460 - ci * 140, 144 + ci * 116, 384, 96)
            elif next_type == "swiss_bento_morph" and st != "swiss_bento_morph":
                add_pic(s, "!!kb_hero", os.path.join(assets_dir, "card_bento_strip.png"), -640, 106, 588, 154)
                for ti in range(3):
                    add_pic(s, f"!!kb_tbg_{ti}", os.path.join(assets_dir, f"card_bento_tile_{ti + 1}.png"), 46 + ti * 296, 580 + ti * 70, 276, 222)
            elif next_type == "chart_infographic" and st != "chart_infographic":
                add_pic(s, "!!ch_img", os.path.join(assets_dir, f"s{idx + 1}_chart.png"), -640, 110, 584, 384)
                add_pic(s, "!!ch_hlbg", os.path.join(assets_dir, "card_spec_1.png"), 1020, 278, 268, 104)
            elif next_type == "timeline_rail_morph" and st != "timeline_rail_morph":
                add_rect(s, "!!tl_rail", -860, 144, 808, 3, c_bdr)
                add_pic(s, "!!tl_drawer", os.path.join(assets_dir, "card_panel.png"), 316, 590, 596, 292)
            elif next_type == "architecture_blueprint" and st != "architecture_blueprint":
                add_pic(s, "!!ar_img", os.path.join(assets_dir, f"s{idx + 1}_arch.png"), -680, 104, 624, 398)
            elif next_type == "closing_finale" and st != "closing_finale":
                add_text(s, "!!fs_h1", next_sl.get("headline", ""), -560, 92, 496, 118, f_disp, 36, True, c_ink)
                add_rect(s, "!!fs_rule", -560, 222, 96, 2.5, c_acc1)
                if hero_glb and not any(sh.name == "!!Hero3D" for sh in s.shapes):
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx + 1}_pre", 14, 24, -2), 1040, 92, 412, 352)

    os.makedirs(os.path.dirname(os.path.abspath(out_pptx)), exist_ok=True)
    prs.save(out_pptx)
    inject_morph_and_alpha_to_pptx(out_pptx, alpha_zero_targets, curtain_wipe_targets=curtain_wipe_targets)
    print(f"Saved cross-platform True-Morph v5 deck -> {out_pptx}")
    return out_pptx


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--spec", required=True)
    parser.add_argument("--assets", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    build_cinema_crossplatform(args.spec, args.assets, args.out)


if __name__ == "__main__":
    main()
