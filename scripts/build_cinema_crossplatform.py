#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Cross-Platform Cinema-Morph OpenXML Builder (macOS / Linux / Headless CI)
for morphdeck-3d v3.0.0.
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


def add_rect(slide, name: str, x: float, y: float, w: float, h: float, col: RGBColor, border_col: RGBColor | None = None, rounded: bool = False):
    stype = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(stype, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = name
    if rounded and len(sh.adjustments) > 0:
        sh.adjustments[0] = 0.14
    sh.fill.solid()
    sh.fill.fore_color.rgb = col
    if border_col is not None:
        sh.line.color.rgb = border_col
        sh.line.width = Pt(1.2)
    else:
        sh.line.fill.background()
    return sh


def add_text(
    slide, name: str, text: str, x: float, y: float, w: float, h: float,
    font: str, size: float, bold: bool, col: RGBColor, align=PP_ALIGN.LEFT, v_anchor=MSO_ANCHOR.TOP,
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


def inject_morph_to_pptx(pptx_path: str, duration_ms: int = 1800) -> None:
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
    f_mono = fonts.get("mono", "Consolas")

    c_surf = hex_rgb(pal.get("surface", "#141824"))
    c_ink = hex_rgb(pal.get("ink", "#FFFFFF"))
    c_muted = hex_rgb(pal.get("ink_muted", "#94A3B8"))
    c_acc1 = hex_rgb(pal.get("accent1", "#E11D48"))
    c_bdr = hex_rgb(pal.get("border", "#334155"))
    c_white = RGBColor(255, 255, 255)

    hero_glb = spec.get("resolved_glb", "")
    pose_cache = {}

    def pose_png(tag: str, rx: float, ry: float, rz: float) -> str:
        if not hero_glb or not os.path.exists(hero_glb):
            return ""
        key = (round(rx, 1), round(ry, 1), round(rz, 1))
        if key not in pose_cache:
            out_p = os.path.join(assets_dir, f"_pose_{tag}.png")
            render_glb_to_image(hero_glb, 720, 720, rx, ry, rz, theme=universe).save(out_p, "PNG")
            pose_cache[key] = out_p
        return pose_cache[key]

    prs = Presentation()
    prs.slide_width = Pt(960)
    prs.slide_height = Pt(540)
    blank = prs.slide_layouts[6]

    for idx, sl in enumerate(spec.get("slides", []), start=1):
        s = prs.slides.add_slide(blank)
        st = sl.get("type", "swiss_bento_morph")

        if st == "newspaper_dive":
            stg = sl.get("stage", "frontpage")
            if stg == "collage":
                add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), -1180, -680, 3456, 1944)
                add_pic(s, "!!col_cutout_L", sl.get("_left_cutout", ""), 32, 94, 336, 446)
                add_pic(s, "!!col_cutout_R", sl.get("_right_cutout", ""), 468, 156, 410, 384)
                add_pic(s, "ns_panel", os.path.join(assets_dir, "card_panel.png"), 496, 36, 420, 202)
                add_text(s, "ns_h1", sl.get("headline", ""), 520, 76, 372, 64, f_disp, 22, True, c_acc1)
                add_text(s, "ns_body", sl.get("body", ""), 520, 144, 372, 78, f_body, 11.5, False, c_ink)
            else:
                add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), 0, 0, 960, 540)
        elif st == "editorial_print_overlay":
            add_pic(s, "!!news_sheet", os.path.join(assets_dir, "news_sheet.png"), -280, -180, 1440, 810)
            add_pic(s, "ep_drips", os.path.join(assets_dir, "ep_drips.png"), 0, 0, 960, 190)
            add_pic(s, "!!ep_port", sl.get("_portrait_cutout", ""), 46, 66, 352, 474)
            add_pic(s, "ep_cloud", os.path.join(assets_dir, "ep_paper_cloud.png"), 384, 108, 552, 372)
            add_text(s, "ep_h1", sl.get("headline", ""), 438, 176, 446, 72, f_disp, 26, True, c_acc1)
            add_text(s, "ep_body", sl.get("body", ""), 438, 256, 446, 140, f_body, 12.5, False, c_ink)
        elif st == "odometer_counter":
            bg_c = hex_rgb(sl.get("bg_hex", "#08090C"))
            add_rect(s, "!!odo_curtain", 0, 0, 960, 540, bg_c)
            add_pic(s, "odo_sil", os.path.join(assets_dir, "odo_silhouette.png"), 0, 378, 960, 162)
            add_text(s, "odo_ovr", sl.get("overline", ""), 120, 144, 720, 34, f_disp, 22, True, c_acc1, PP_ALIGN.CENTER)
            add_text(s, "odo_val", sl.get("target_value", ""), 120, 218, 720, 88, f_head, 66, True, c_white, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            add_text(s, "odo_sub", sl.get("sublabel", ""), 120, 322, 720, 34, f_disp, 24, True, c_white, PP_ALIGN.CENTER)
        elif st == "real_portal_flythrough":
            stg = sl.get("stage", "exterior")
            add_pic(s, "!!portal_bg", os.path.join(assets_dir, "portal_bg.png"), 0, 0, 960, 540)
            if stg == "interior":
                add_pic(s, "!!portal_outer", os.path.join(assets_dir, "portal_outer.png"), -1824, -1026, 4608, 2592)
                add_pic(s, "!!portal_inner", os.path.join(assets_dir, "portal_inner.png"), 0, 0, 960, 540)
                add_pic(s, "!!pt_card_next", os.path.join(assets_dir, "card_panel.png"), 56, 102, 426, 336)
                add_text(s, "pt_h1", sl.get("headline", ""), 80, 146, 378, 68, f_disp, 24, True, c_ink)
                if hero_glb:
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, -22, 2), 542, 78, 308, 348)
            else:
                if hero_glb:
                    add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 16, -28, -4), 392, 152, 176, 210)
                add_pic(s, "!!portal_inner", os.path.join(assets_dir, "portal_inner.png"), 272, 102, 416, 336)
                add_pic(s, "!!portal_outer", os.path.join(assets_dir, "portal_outer.png"), 0, 0, 960, 540)
                add_text(s, "pt_h1", sl.get("headline", ""), 48, 38, 864, 44, f_disp, 26, True, c_white, PP_ALIGN.CENTER)
        elif st == "poster_depth_reveal":
            add_pic(s, "!!frame_art", sl.get("_poster_bg_image") or os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            add_rect(s, "!!poster_band", 348, 0, 264, 540, c_acc1)
            add_text(s, "!!poster_word", sl.get("giant_word", "2026"), 60, 52, 840, 168, f_head, 120, True, c_white, PP_ALIGN.CENTER)
            if sl.get("_hero_cutout"):
                add_pic(s, "!!poster_hero", sl["_hero_cutout"], 170, 76, 620, 464)
            elif hero_glb and sl.get("use_hero_3d", True):
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, -22, 2), 288, 76, 384, 374)
        elif st == "picture_frame_pullback":
            add_pic(s, "!!wall_room", sl.get("_wall_image") or os.path.join(assets_dir, "bg_alt.png"), 0, 0, 960, 540)
            fx, fy, fw, fh = sl.get("frame_x", 348), sl.get("frame_y", 42), sl.get("frame_w", 264), sl.get("frame_h", 148)
            add_pic(s, "!!frame_art", sl.get("_framed_art") or os.path.join(assets_dir, "bg_alt.png"), fx, fy, fw, fh)
            if sl.get("_hollow_frame"):
                add_pic(s, "!!hollow_frame", sl["_hollow_frame"], fx - fw * 0.0755, fy - fh * 0.152, fw / 0.86875, fh / 0.76667)
            add_pic(s, "pf_room_subj", sl.get("_room_subject_cutout", ""), 470, 156, 468, 384)
            add_text(s, "pf_h1", sl.get("headline", ""), 62, 272, 360, 92, f_disp, 30, True, c_white)
        elif st == "chart_infographic":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "ch_h1", sl.get("headline", ""), 46, 46, 600, 54, f_disp, 25, True, c_ink)
            add_pic(s, "ch_img", os.path.join(assets_dir, f"s{idx}_chart.png"), 46, 110, 584, 384)
            if hero_glb:
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 14, -24, 2), 672, 38, 220, 222)
        elif st == "architecture_blueprint":
            add_pic(s, "bg_slide", os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "ar_h1", sl.get("headline", ""), 42, 46, 630, 52, f_disp, 24, True, c_ink)
            add_pic(s, "ar_img", os.path.join(assets_dir, f"s{idx}_arch.png"), 42, 104, 624, 398)
            if hero_glb:
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 22, -34, -3), 698, 36, 212, 222)
        else:
            add_pic(s, "bg_slide", sl.get("_bg_image") or os.path.join(assets_dir, "bg_base.png"), 0, 0, 960, 540)
            add_text(s, "sl_ovr", sl.get("overline", ""), 52, 34, 620, 20, f_mono, 9.5, True, c_acc1)
            add_text(s, "sl_h1", sl.get("headline", sl.get("giant_word", "")), 52, 60, 560, 78, f_disp, 28, True, c_ink)
            add_text(s, "sl_body", sl.get("body", sl.get("summary", "")), 52, 152, 460, 120, f_body, 12.5, False, c_muted)
            if sl.get("_hero_cutout"):
                add_pic(s, "!!HeroCut", sl["_hero_cutout"], 520, 78, 370, 390)
            elif hero_glb:
                add_pic(s, "!!Hero3D", pose_png(f"s{idx}", 12, -22, 2), 540, 82, 340, 360)
            kpis = sl.get("kpis") or sl.get("metrics") or []
            for ki, k in enumerate(kpis[:3]):
                add_rect(s, f"k_{ki}", 52 + ki * 158, 360, 146, 74, c_surf, c_bdr, rounded=True)
                add_text(s, f"kv_{ki}", k.get("value", ""), 60 + ki * 158, 370, 130, 32, f_disp, 20, True, c_acc1, PP_ALIGN.CENTER)

    os.makedirs(os.path.dirname(os.path.abspath(out_pptx)), exist_ok=True)
    prs.save(out_pptx)
    inject_morph_to_pptx(out_pptx)
    print(f"Saved cross-platform Cinema-Morph deck -> {out_pptx}")
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
