#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Automated OpenXML, Morph, GLB, Emoji, and 1080p Render Verification Gate
for morphdeck-3d presentations.

Exits with code 0 only if every check passes.
"""

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from PIL import Image, ImageStat

EMOJI_RE = re.compile(
    "[\U0001F300-\U0001FAFF\u2600-\u27BF\u2300-\u23FF]",
    flags=re.UNICODE,
)


def verify_pptx(
    pptx_path: str,
    renders_dir: str = None,
    require_3d: bool = True,
    allow_emojis: bool = False,
) -> int:
    failures = []
    if not os.path.exists(pptx_path):
        print(f"FAIL: File not found: {pptx_path}")
        return 1

    size_bytes = os.path.getsize(pptx_path)
    with zipfile.ZipFile(pptx_path, "r") as z:
        names = z.namelist()

        # 1. XML & Relationship parse check + external link leak check
        xml_errors = 0
        ext_links = []
        for n in names:
            if n.endswith((".xml", ".rels")):
                raw = z.read(n)
                try:
                    root = ET.fromstring(raw)
                    if n.endswith(".rels"):
                        for rel in root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                            if rel.attrib.get("TargetMode") == "External":
                                ext_links.append((n, rel.attrib.get("Target", "")))
                except ET.ParseError as e:
                    print(f"XML ERROR in {n}: {e}")
                    xml_errors += 1

        if xml_errors:
            failures.append(f"{xml_errors} XML/rels parse error(s)")
        if ext_links:
            failures.append(f"External relationships leaked: {ext_links}")

        # 2. Slide-level Morph & Emoji check
        slides = [n for n in names if re.match(r"^ppt/slides/slide\d+\.xml$", n)]
        slides.sort(key=lambda s: int(re.search(r"slide(\d+)\.xml$", s).group(1)))
        if not slides:
            failures.append("Presentation contains 0 slides")

        emoji_hits = []
        pt_light_count = 0
        model3d_count = 0
        for s in slides:
            txt = z.read(s).decode("utf-8")
            pt_light_count += txt.count("<am3d:ptLight")
            model3d_count += txt.count("<am3d:model3d")
            if "mc:AlternateContent" not in txt or "p159:morph" not in txt:
                failures.append(f"{s} missing compliant <mc:AlternateContent> + <p159:morph>")
            if not allow_emojis:
                root = ET.fromstring(txt)
                for t in root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                    if t.text and EMOJI_RE.search(t.text):
                        emoji_hits.append((s, t.text))

        if emoji_hits:
            failures.append(f"Emoji characters found in text runs: {emoji_hits[:3]}")

        # 3. Embedded 3D GLB verification
        glbs = [n for n in names if n.startswith("ppt/media/") and n.endswith(".glb")]
        glb_sizes = [(g, z.getinfo(g).file_size) for g in glbs]
        if require_3d and not glbs:
            failures.append("No embedded .glb 3D subject model found in ppt/media/")

    # 4. Optional 1080p Render PNG & non-blank variance check
    if renders_dir:
        for i in range(1, len(slides) + 1):
            rp = os.path.join(renders_dir, f"slide_{i}.png")
            if not os.path.exists(rp) or os.path.getsize(rp) == 0:
                failures.append(f"Missing/empty rendered PNG: {rp}")
            else:
                with Image.open(rp) as im:
                    if im.size != (1920, 1080):
                        failures.append(f"{rp} expected (1920, 1080), got {im.size}")
                    max_std = max(ImageStat.Stat(im.convert("RGB")).stddev)
                    if max_std <= 5.0:
                        failures.append(f"{rp} appears blank/flat (max channel stddev={max_std:.2f})")

    print(f"Package: {pptx_path} ({size_bytes} bytes)")
    print(f"Slides verified: {len(slides)} | XML errors: {xml_errors} | External links: {len(ext_links)}")
    if glb_sizes:
        print(f"Embedded GLBs ({len(glbs)}): {glb_sizes}")
        print(f"Studio 3D Lights: {pt_light_count} point light(s) across {model3d_count} <am3d:model3d> node(s)")
    else:
        print("Embedded GLBs (0): cross-platform studio-rendered 3D deck (OK)")
    if renders_dir:
        print(f"Renders checked in: {renders_dir}")

    if failures:
        print("FAILED CHECKS:")
        for f in failures:
            print("  -", f)
        return 1

    print("ALL VERIFICATION GATES PASSED CLEANLY.")
    return 0


def main():
    parser = argparse.ArgumentParser(description="Verify PPTX integrity, Morph compliance, 3D GLB embedding, and 1080p renders.")
    parser.add_argument("pptx", help="Path to generated .pptx presentation")
    parser.add_argument("--renders-dir", help="Optional directory containing exported slide_1..slide_N.png renders")
    parser.add_argument("--crossplatform", action="store_true", help="Allow cross-platform 2D-rendered 3D builds without binary .glb embedding")
    parser.add_argument("--allow-emojis", action="store_true", help="Allow Unicode emojis in slide text")
    args = parser.parse_args()
    sys.exit(verify_pptx(args.pptx, args.renders_dir, require_3d=not args.crossplatform, allow_emojis=args.allow_emojis))


if __name__ == "__main__":
    main()
