#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Unified Turnkey CLI for morphdeck-3d v2.0.0.

Subcommands:
  list       Print all 10 procedural 3D archetypes, 7 themes, and 19 slide methods.
  init       Initialize a starter deck_spec.json from a domain example template.
  validate   Run pre-flight JSON schema & typography checks on a deck_spec.json.
  preview-3d Generate any procedural 3D .glb archetype and render a studio RGBA PNG.
  build      Full pipeline: validate -> render assets -> COM/OpenXML + 3D studio lighting -> verify.
  verify     Run OpenXML, Morph, GLB, studio lights, and 1080p render verification gates on a .pptx.
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
EXAMPLES_DIR = os.path.join(REPO_ROOT, "examples")

if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from generate_subject_glb import ARCHETYPES, generate_subject_glb  # noqa: E402
from inject_3d_studio_lighting import inject_3d_studio_lighting  # noqa: E402
from render_glb_preview import render_glb_to_image  # noqa: E402
from render_studio_assets import THEMES, generate_all  # noqa: E402
from validate_spec import VALID_SLIDE_TYPES, validate_spec  # noqa: E402
from verify_deck import verify_pptx  # noqa: E402
from build_deck_crossplatform import build_crossplatform_pptx  # noqa: E402

TEMPLATE_MAP = {
    "aerospace": "aerospace_portal_cinema.json",
    "architecture": "architecture_pavilion.json",
    "ai": "ai_cloud_cluster.json",
    "product": "product_drop_launch.json",
    "executive": "executive_finops.json",
    "biotech": "biotech_therapeutics.json",
    "mobile": "mobile_saas_platform.json",
}


def cmd_list(_args) -> int:
    print("=== morphdeck-3d v2.0.0 Capabilities (Apache-2.0 | Mohamed S. Sissoko / 7AK3SH1) ===")
    print(f"\n10 Procedural 3D .glb Archetypes (+ aliases: iphone -> smartphone_pro, macbook -> laptop_studio):")
    for a in ARCHETYPES:
        print(f"  - {a}")
    print(f"\n7 Visual Theme Systems (+ 3-Point <am3d:ptLight> Studio Lighting Rigs):")
    for name, cfg in THEMES.items():
        print(f"  - {name:20s} | default 3D: {cfg['default_archetype']:26s} | display: {cfg['display_font']}")
    print(f"\n19 Composable Slide Methods:")
    for st in sorted(VALID_SLIDE_TYPES):
        print(f"  - {st}")
    return 0


def cmd_init(args) -> int:
    src = os.path.join(EXAMPLES_DIR, TEMPLATE_MAP[args.template])
    if not os.path.exists(src):
        print(f"ERROR: Template not found: {src}")
        return 1
    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    shutil.copyfile(src, args.output)
    print(f"Initialized '{args.template}' starter spec -> {args.output}")
    return 0


def cmd_validate(args) -> int:
    errors, warnings = validate_spec(args.spec, allow_emojis=args.allow_emojis)
    for w in warnings:
        print(f"WARN: {w}")
    for e in errors:
        print(f"ERROR: {e}")
    if errors:
        return 1
    print(f"VALID: {args.spec}")
    return 0


def cmd_preview_3d(args) -> int:
    with tempfile.TemporaryDirectory() as tmp:
        glb_path = os.path.join(tmp, f"{args.archetype}.glb")
        generate_subject_glb(
            archetype=args.archetype,
            out_glb=glb_path,
            brand_name=args.brand,
            accent1_hex=args.accent1,
            accent2_hex=args.accent2,
            accent3_hex=args.accent3,
        )
        img = render_glb_to_image(glb_path, args.size, args.size, args.rx, args.ry, args.rz, theme=args.theme)
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        img.save(args.output, "PNG")
        print(f"Preview rendered ({args.archetype}, {args.size}x{args.size}) -> {args.output}")
    return 0


def _can_use_com(engine: str) -> bool:
    if engine == "crossplatform":
        return False
    if engine == "com":
        return sys.platform.startswith("win")
    # auto: use COM on Windows only if PowerPoint executable / ProgID is registered
    if not sys.platform.startswith("win"):
        return False
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "if ([Type]::GetTypeFromProgID('PowerPoint.Application')) { exit 0 } else { exit 1 }"],
            capture_output=True,
            timeout=10,
        )
        return res.returncode == 0
    except Exception:
        return False


def cmd_build(args) -> int:
    errors, warnings = validate_spec(args.spec)
    for w in warnings:
        print(f"WARN: {w}")
    if errors:
        for e in errors:
            print(f"ERROR: {e}")
        return 1

    with open(args.spec, "r", encoding="utf-8") as f:
        spec_data = json.load(f)
    spec_theme = spec_data.get("theme", "obsidian_kinetic")

    work_dir = args.work_dir or tempfile.mkdtemp(prefix="morphdeck_work_")
    assets_dir = os.path.join(work_dir, "assets")
    renders_dir = args.renders or os.path.join(work_dir, "renders")
    use_com = _can_use_com(args.engine)

    try:
        generate_all(args.spec, assets_dir)

        if use_com:
            ps1 = os.path.join(SCRIPT_DIR, "build_deck.ps1")
            out_abs = os.path.abspath(args.out)
            renders_abs = os.path.abspath(renders_dir)
            cmd = [
                "powershell",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                ps1,
                "-SpecJson",
                os.path.abspath(args.spec),
                "-AssetsDir",
                os.path.abspath(assets_dir),
                "-OutputPptx",
                out_abs,
                "-ExportDir",
                renders_abs,
                "-SkipExport",
            ]
            proc = subprocess.run(cmd)
            if proc.returncode != 0:
                return proc.returncode

            inject_3d_studio_lighting(out_abs, theme=spec_theme)

            export_cmd = [
                "powershell",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                ps1,
                "-OutputPptx",
                out_abs,
                "-ExportDir",
                renders_abs,
                "-ExportOnly",
            ]
            if args.copy_to:
                export_cmd.extend(["-CopyTo", os.path.abspath(args.copy_to)])
            proc_exp = subprocess.run(export_cmd)
            if proc_exp.returncode != 0:
                return proc_exp.returncode

            return verify_pptx(out_abs, renders_dir=renders_abs, require_3d=True)
        else:
            build_crossplatform_pptx(args.spec, assets_dir, args.out)
            if args.copy_to:
                shutil.copyfile(args.out, args.copy_to)
            return verify_pptx(args.out, renders_dir=None, require_3d=False)
    finally:
        if not args.work_dir and not args.keep_work and os.path.exists(work_dir):
            shutil.rmtree(work_dir, ignore_errors=True)


def cmd_verify(args) -> int:
    return verify_pptx(
        args.pptx,
        renders_dir=args.renders_dir,
        require_3d=not args.crossplatform,
        allow_emojis=args.allow_emojis,
    )


def main():
    parser = argparse.ArgumentParser(prog="morphdeck", description="morphdeck-3d Turnkey 3D + Morph PowerPoint CLI")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="List 3D archetypes, themes, and slide methods")
    p_list.set_defaults(func=cmd_list)

    p_init = sub.add_parser("init", help="Copy a starter deck_spec.json template")
    p_init.add_argument("--template", choices=sorted(TEMPLATE_MAP.keys()), required=True)
    p_init.add_argument("--output", required=True)
    p_init.set_defaults(func=cmd_init)

    p_val = sub.add_parser("validate", help="Validate deck_spec.json")
    p_val.add_argument("spec")
    p_val.add_argument("--allow-emojis", action="store_true")
    p_val.set_defaults(func=cmd_validate)

    p_prev = sub.add_parser("preview-3d", help="Render a 3D archetype preview PNG")
    p_prev.add_argument("--archetype", choices=ARCHETYPES, required=True)
    p_prev.add_argument("--output", required=True)
    p_prev.add_argument("--brand", default="MORPHDECK")
    p_prev.add_argument("--accent1", default="#2563EB")
    p_prev.add_argument("--accent2", default="#10B981")
    p_prev.add_argument("--accent3", default="#F59E0B")
    p_prev.add_argument("--size", type=int, default=720)
    p_prev.add_argument("--rx", type=float, default=14.0)
    p_prev.add_argument("--ry", type=float, default=-24.0)
    p_prev.add_argument("--rz", type=float, default=2.0)
    p_prev.add_argument("--theme", default="obsidian_kinetic", choices=sorted(THEMES.keys()))
    p_prev.set_defaults(func=cmd_preview_3d)

    p_build = sub.add_parser("build", help="Validate, render assets, build .pptx, and verify")
    p_build.add_argument("spec", help="Path to deck_spec.json")
    p_build.add_argument("--out", required=True, help="Output .pptx path")
    p_build.add_argument("--renders", default="", help="Optional PNG export folder (COM engine)")
    p_build.add_argument("--copy-to", default="", help="Optional secondary copy destination")
    p_build.add_argument("--work-dir", default="", help="Optional explicit temp asset directory")
    p_build.add_argument("--engine", choices=["auto", "com", "crossplatform"], default="auto")
    p_build.add_argument("--keep-work", action="store_true", help="Keep temporary work directory")
    p_build.set_defaults(func=cmd_build)

    p_ver = sub.add_parser("verify", help="Verify generated .pptx package")
    p_ver.add_argument("pptx")
    p_ver.add_argument("--renders-dir", default=None)
    p_ver.add_argument("--crossplatform", action="store_true")
    p_ver.add_argument("--allow-emojis", action="store_true")
    p_ver.set_defaults(func=cmd_verify)

    args = parser.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
