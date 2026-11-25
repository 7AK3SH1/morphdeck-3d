#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
3-Point Colored Studio Lighting Injector (<am3d:model3d>) for morphdeck-3d.

Upgrades Microsoft PowerPoint's default white <am3d:model3d> illumination with
theme-calibrated <am3d:ambientLight> + 3 colored <am3d:ptLight> nodes
(warm key light, theme accent fill light, and crisp cool/tinted rim backlight)
while strictly preserving XML namespace prefixes (am3d:, a:, p:, mc:, p159:).
"""

import argparse
import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET

# OpenXML scrgbClr uses 0..100000 linear percentage units.
# Coordinates & intensities match PowerPoint <am3d:ptLight> schema (~1e7 EMU scale).
THEME_LIGHT_RIGS = {
    "obsidian_kinetic": {
        "ambient": (46000, 48000, 54000, 480000),
        "lights": [
            # Warm amber-coral key light (upper right front)
            (100000, 76000, 56000, 9800000, 22000000, 70500000, 18500000),
            # Cool cinema sky-blue fill light (left front)
            (40000, 65000, 98000, 11800000, -38000000, 51000000, 56000000),
            # Crisp violet-white rim backlight (left rear)
            (88000, 74000, 100000, 3800000, -37000000, 58000000, -35000000),
        ],
    },
    "editorial_ivory": {
        "ambient": (54000, 51000, 47000, 520000),
        "lights": [
            # Terracotta-gold key light
            (100000, 74000, 48000, 9800000, 24000000, 70000000, 20000000),
            # Soft warm ivory-teal fill
            (46000, 72000, 78000, 10600000, -36000000, 50000000, 55000000),
            # Warm champagne rim
            (100000, 86000, 70000, 3600000, -34000000, 56000000, -34000000),
        ],
    },
    "cyber_blueprint": {
        "ambient": (40000, 48000, 60000, 470000),
        "lights": [
            # Clean neutral-warm key for high-contrast specular read
            (92000, 95000, 100000, 9600000, 22000000, 71000000, 18000000),
            # Electric cyan fill
            (24000, 78000, 96000, 12200000, -38000000, 51000000, 58000000),
            # Indigo-violet rim backlight
            (62000, 58000, 100000, 4200000, -38000000, 58000000, -36000000),
        ],
    },
    "kinetic_poster": {
        "ambient": (46000, 44000, 50000, 460000),
        "lights": [
            # Crisp warm-amber key
            (100000, 80000, 58000, 10200000, 22500000, 70000000, 18000000),
            # Electric cobalt fill
            (36000, 60000, 98000, 11600000, -37500000, 51000000, 57000000),
            # Crimson-magenta rim kick
            (96000, 42000, 62000, 4000000, -37000000, 58000000, -35000000),
        ],
    },
    "executive_clean": {
        "ambient": (48000, 50000, 56000, 500000),
        "lights": [
            # Daylight warm-white key
            (100000, 92000, 82000, 9600000, 22000000, 70500000, 17500000),
            # Boardroom azure-teal fill
            (38000, 68000, 94000, 11400000, -38000000, 51000000, 56000000),
            # Platinum-blue rim
            (82000, 88000, 100000, 3500000, -36500000, 57500000, -35000000),
        ],
    },
    "biotech_luminance": {
        "ambient": (42000, 54000, 52000, 480000),
        "lights": [
            # Clean optical white-warm key
            (96000, 98000, 88000, 9600000, 22000000, 70500000, 18000000),
            # Bio-emerald cyan fill
            (28000, 86000, 76000, 12000000, -38000000, 51000000, 57500000),
            # Lime-cyan rim backlight
            (64000, 96000, 88000, 4000000, -37500000, 58000000, -35000000),
        ],
    },
    "monochrome_luxury": {
        "ambient": (48000, 46000, 44000, 490000),
        "lights": [
            # Champagne gold key
            (100000, 84000, 52000, 9900000, 22000000, 70500000, 17000000),
            # Cool steel-silver fill
            (58000, 68000, 86000, 11200000, -37500000, 51000000, 56500000),
            # Pearl violet-white rim
            (92000, 86000, 100000, 3600000, -37000000, 58000000, -35000000),
        ],
    },
}

MODEL3D_BLOCK_RE = re.compile(r"(<am3d:model3d\b[^>]*>)(.*?)(</am3d:model3d>)", re.DOTALL)
EXISTING_LIGHTS_RE = re.compile(
    r"<am3d:(?:ambientLight|ptLight|dirLight|spotLight|unknownLight)\b[^>]*(?:/>|>.*?</am3d:(?:ambientLight|ptLight|dirLight|spotLight|unknownLight)>)",
    re.DOTALL,
)


def _build_lighting_xml(theme: str) -> str:
    rig = THEME_LIGHT_RIGS.get(theme, THEME_LIGHT_RIGS["obsidian_kinetic"])
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


def inject_3d_studio_lighting(pptx_path: str, theme: str = "obsidian_kinetic") -> int:
    """
    Upgrade every <am3d:model3d> node in pptx_path with a 3-point colored studio lighting rig.
    Preserves all XML namespace prefixes (am3d:, a:, p:, mc:, p159:) and validates XML before saving.
    Returns the total number of <am3d:model3d> blocks upgraded.
    """
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

    lighting_xml = _build_lighting_xml(theme)
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

    print(f"Studio 3D lighting rig injected ({theme}): upgraded {upgraded_count} <am3d:model3d> node(s) -> {pptx_path}")
    return upgraded_count


def main() -> None:
    parser = argparse.ArgumentParser(description="Inject 3-point colored studio lighting into <am3d:model3d> PPTX nodes.")
    parser.add_argument("--pptx", required=True, help="Target .pptx file path")
    parser.add_argument(
        "--theme",
        default="obsidian_kinetic",
        choices=sorted(THEME_LIGHT_RIGS.keys()),
        help="morphdeck-3d visual theme lighting preset",
    )
    args = parser.parse_args()
    inject_3d_studio_lighting(args.pptx, theme=args.theme)


if __name__ == "__main__":
    main()
