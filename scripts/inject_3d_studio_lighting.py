#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
3-Point Colored Studio Lighting Injector (<am3d:model3d>) for morphdeck-3d.

Upgrades PowerPoint's default flat white <am3d:model3d> illumination with
universe- and theme-calibrated <am3d:ambientLight> + 3 colored <am3d:ptLight>
emitters (warm key light, accent fill light, and rim backlight) while strictly
preserving OpenXML namespace prefixes (am3d:, a:, p:, mc:, p159:).
"""

import argparse
import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET

LIGHT_RIGS = {
    # v3 Visual Universes
    "documentary_press": {
        "ambient": (52000, 49000, 44000, 500000),
        "lights": [
            (100000, 82000, 60000, 9800000, 22000000, 70000000, 18500000),
            (58000, 68000, 82000, 10800000, -36000000, 50000000, 54000000),
            (96000, 64000, 52000, 3800000, -35000000, 56000000, -34000000),
        ],
    },
    "architectural_portal": {
        "ambient": (54000, 52000, 49000, 530000),
        "lights": [
            (100000, 88000, 72000, 9900000, 24000000, 70000000, 20000000),
            (52000, 76000, 92000, 11000000, -36000000, 51000000, 56000000),
            (100000, 90000, 76000, 3600000, -34000000, 57000000, -34000000),
        ],
    },
    "sculptural_museum": {
        "ambient": (50000, 48000, 45000, 490000),
        "lights": [
            (100000, 78000, 52000, 9800000, 22000000, 71000000, 16500000),
            (44000, 64000, 95000, 12000000, -38000000, 51000000, 57500000),
            (92000, 82000, 100000, 3600000, -37500000, 58000000, -35000000),
        ],
    },
    "commercial_studio": {
        "ambient": (46000, 48000, 54000, 480000),
        "lights": [
            (100000, 82000, 66000, 10200000, 22500000, 70500000, 18000000),
            (38000, 68000, 98000, 11800000, -38000000, 51000000, 56500000),
            (94000, 70000, 100000, 4000000, -37000000, 58000000, -35000000),
        ],
    },
    "executive_swiss": {
        "ambient": (52000, 54000, 58000, 520000),
        "lights": [
            (100000, 95000, 88000, 9600000, 22000000, 70500000, 17500000),
            (44000, 70000, 96000, 11200000, -38000000, 51000000, 56000000),
            (84000, 90000, 100000, 3500000, -36500000, 57500000, -35000000),
        ],
    },
    "editorial_pop": {
        "ambient": (50000, 48000, 52000, 500000),
        "lights": [
            (100000, 80000, 60000, 10200000, 22500000, 70000000, 18000000),
            (36000, 72000, 98000, 11600000, -37500000, 51000000, 57000000),
            (98000, 52000, 74000, 3900000, -37000000, 58000000, -35000000),
        ],
    },
    # v2 Backward-Compatible Studio Themes
    "obsidian_kinetic": {
        "ambient": (45000, 48000, 55000, 480000),
        "lights": [
            (100000, 80000, 58000, 10000000, 22500000, 70500000, 17500000),
            (35000, 72000, 100000, 12000000, -38000000, 51000000, 57000000),
            (95000, 65000, 100000, 4000000, -37000000, 58000000, -35000000),
        ],
    },
    "editorial_ivory": {
        "ambient": (55000, 53000, 50000, 550000),
        "lights": [
            (100000, 85000, 65000, 9500000, 23000000, 70000000, 19000000),
            (55000, 75000, 95000, 11000000, -36000000, 50000000, 55000000),
            (100000, 90000, 80000, 3500000, -35000000, 56000000, -34000000),
        ],
    },
    "cyber_blueprint": {
        "ambient": (40000, 48000, 62000, 470000),
        "lights": [
            (92000, 96000, 100000, 9500000, 21000000, 71000000, 16000000),
            (25000, 82000, 100000, 13000000, -39000000, 51000000, 58000000),
            (60000, 100000, 92000, 4200000, -37000000, 58000000, -35000000),
        ],
    },
    "kinetic_poster": {
        "ambient": (48000, 46000, 50000, 500000),
        "lights": [
            (100000, 76000, 52000, 10500000, 22500000, 70500000, 18000000),
            (40000, 78000, 100000, 12500000, -37500000, 51000000, 57000000),
            (100000, 55000, 70000, 4000000, -37000000, 58000000, -35000000),
        ],
    },
    "executive_clean": {
        "ambient": (52000, 54000, 58000, 520000),
        "lights": [
            (100000, 96000, 90000, 9500000, 22000000, 70500000, 17500000),
            (42000, 68000, 96000, 11000000, -38000000, 51000000, 56000000),
            (85000, 90000, 100000, 3500000, -36500000, 57500000, -35000000),
        ],
    },
    "biotech_luminance": {
        "ambient": (48000, 55000, 54000, 510000),
        "lights": [
            (95000, 100000, 92000, 9800000, 22500000, 70500000, 17500000),
            (32000, 86000, 90000, 11500000, -38000000, 51000000, 57000000),
            (65000, 92000, 100000, 3800000, -37000000, 58000000, -35000000),
        ],
    },
    "monochrome_luxury": {
        "ambient": (48000, 48000, 50000, 470000),
        "lights": [
            (100000, 88000, 62000, 9800000, 22500000, 70500000, 17500000),
            (65000, 68000, 75000, 10500000, -38000000, 51000000, 56000000),
            (100000, 92000, 78000, 3600000, -37000000, 58000000, -35000000),
        ],
    },
}

THEME_LIGHT_RIGS = LIGHT_RIGS
UNIVERSE_LIGHT_RIGS = LIGHT_RIGS

MODEL3D_BLOCK_RE = re.compile(r"(<am3d:model3d\b[^>]*>)(.*?)(</am3d:model3d>)", re.DOTALL)
EXISTING_LIGHTS_RE = re.compile(
    r"<am3d:(?:ambientLight|ptLight|dirLight|spotLight|unknownLight)\b[^>]*(?:/>|>.*?</am3d:(?:ambientLight|ptLight|dirLight|spotLight|unknownLight)>)",
    re.DOTALL,
)


def _build_lighting_xml(theme_or_universe: str) -> str:
    rig = LIGHT_RIGS.get(theme_or_universe, LIGHT_RIGS["commercial_studio"])
    ar, ag, ab, a_ill = rig["ambient"]
    parts = [
        f'<am3d:ambientLight><am3d:clr><a:scrgbClr r="{ar}" g="{ag}" b="{ab}"/></am3d:clr>'
        f'<am3d:illuminance n="{a_ill}" d="1000000"/></am3d:ambientLight>'
    ]
    for r, g, b, inten, px, py, pz in rig["lights"]:
        parts.append(
            f'<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="{r}" g="{g}" b="{b}"/></am3d:clr>'
            f'<am3d:intensity n="{inten}" d="1000000"/>'
            f'<am3d:pos x="{px}" y="{py}" z="{pz}"/></am3d:ptLight>'
        )
    return "".join(parts)


def inject_3d_studio_lighting(pptx_path: str, theme: str = "commercial_studio", universe: str = "") -> int:
    """Upgrade every <am3d:model3d> node in pptx_path with a 3-point colored studio lighting rig."""
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

    profile = universe or theme
    lighting_xml = _build_lighting_xml(profile)
    upgraded_count = 0

    def _replace_model_lights(match: re.Match) -> str:
        nonlocal upgraded_count
        open_tag, inner, close_tag = match.group(1), match.group(2), match.group(3)
        cleaned = EXISTING_LIGHTS_RE.sub("", inner)
        ext_idx = cleaned.find("<am3d:extLst")
        if ext_idx != -1:
            updated_inner = cleaned[:ext_idx] + lighting_xml + cleaned[ext_idx:]
        else:
            updated_inner = cleaned + lighting_xml
        upgraded_count += 1
        return f"{open_tag}{updated_inner}{close_tag}"

    buf = io.BytesIO()
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r"^ppt/slides/slide\d+\.xml$", item.filename):
                txt = data.decode("utf-8")
                if "<am3d:model3d" in txt:
                    new_txt = MODEL3D_BLOCK_RE.sub(_replace_model_lights, txt)
                    ET.fromstring(new_txt)
                    data = new_txt.encode("utf-8")
            zout.writestr(item, data)

    with open(pptx_path, "wb") as f:
        f.write(buf.getvalue())

    print(f"Studio 3D lighting rig injected ({profile}): upgraded {upgraded_count} <am3d:model3d> node(s) -> {pptx_path}")
    return upgraded_count


def main() -> None:
    parser = argparse.ArgumentParser(description="Inject 3-point colored studio lighting into <am3d:model3d> PPTX nodes.")
    parser.add_argument("--pptx", required=True, help="Target .pptx file path")
    parser.add_argument("--theme", default="commercial_studio", choices=sorted(LIGHT_RIGS.keys()))
    args = parser.parse_args()
    inject_3d_studio_lighting(args.pptx, theme=args.theme)


if __name__ == "__main__":
    main()
