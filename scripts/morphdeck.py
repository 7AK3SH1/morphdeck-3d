#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Unified CLI for morphdeck-3d v3.0.0 ("Cinema-Morph & Photorealistic 3D").

Commands:
    universes   List Cinema-Morph v3 Visual Universes & legacy themes
    list        Alias for catalog listing (subjects, meshes, themes, universes)
    init        Scaffold a ready-to-edit deck JSON spec
    fetch-3d    Download/search real high-poly CC0 .glb models (Khronos / Poly Haven / NASA)
    fetch-photo Download open-license 4K landscape photography (Wikimedia Commons)
    cutout      Segment studio images, punch portals, or split skies
    validate    Validate a deck JSON specification against v2 or v3 schema rules
    preview-3d  Render an isometric PNG preview from any .glb model
    build       Build a native Morph PPTX deck + studio visuals (Windows COM or cross-platform)
    verify      Run OpenXML integrity & visual PNG QA verification on a built PPTX
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
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from build_cinema_crossplatform import build_cinema_crossplatform  # noqa: E402
from build_deck_crossplatform import build_crossplatform_pptx  # noqa: E402
from fetch_real_glb import fetch_khronos_glb, fetch_nasa_glb, fetch_polyhaven_glb, inspect_glb_stats, resolve_auto_glb  # noqa: E402
from fetch_real_photo import fetch_wikimedia_photo  # noqa: E402
from inject_3d_studio_lighting import inject_3d_studio_lighting  # noqa: E402
from prepare_cinema_assets import UNIVERSES, prepare_all  # noqa: E402
from render_glb_preview import render_glb_to_image  # noqa: E402
from render_studio_assets import ARCHETYPES, THEMES, generate_all  # noqa: E402
from studio_cutout import make_studio_cutout, punch_portal_mask, split_horizon_cutout  # noqa: E402
from validate_spec import VALID_ARCHETYPES, VALID_CINEMA_SLIDE_TYPES, VALID_THEMES, validate_spec  # noqa: E402
from verify_deck import verify_pptx  # noqa: E402


def cmd_universes(_args: argparse.Namespace) -> None:
    print("morphdeck-3d v3.0.0 Cinema-Morph Visual Universes:")
    for uname, u in UNIVERSES.items():
        mode = "light" if u["is_light"] else "dark "
        print(f"  - {uname:22s} ({mode} | bg={u['bg']} accent={u['accents'][0]} | {u['display_font']})")
    print(f"\nCinema-Morph v3 Slide Types ({len(VALID_CINEMA_SLIDE_TYPES)}):")
    for st in sorted(VALID_CINEMA_SLIDE_TYPES):
        print(f"  - {st}")


def cmd_list(_args: argparse.Namespace) -> None:
    cmd_universes(_args)
    print("\nLegacy v2 Themes:")
    for tname, t in THEMES.items():
        print(f"  - {tname:20s} ({'light' if t['is_light'] else 'dark'} | accent={t['accents'][0]})")
    print("\nLegacy v2 Archetypes:")
    for a in sorted(VALID_ARCHETYPES):
        print(f"  - {a}")


def cmd_init(args: argparse.Namespace) -> None:
    example_map = {
        "documentary_press": os.path.join(REPO_ROOT, "examples", "documentary_press_cinema.json"),
        "architectural_portal": os.path.join(REPO_ROOT, "examples", "luxury_portal_optics.json"),
        "commercial_studio": os.path.join(REPO_ROOT, "examples", "commercial_sneaker_drop.json"),
        "spatial_3d": os.path.join(REPO_ROOT, "examples", "ev_hypercar_showcase.json"),
        "bento_executive": os.path.join(REPO_ROOT, "examples", "fintech_q4_board.json"),
        "technical_blueprint": os.path.join(REPO_ROOT, "examples", "ai_infra_architecture.json"),
        "chronicle_timeline": os.path.join(REPO_ROOT, "examples", "biotech_clinical_trial.json"),
        "editorial_typography": os.path.join(REPO_ROOT, "examples", "architectural_monograph.json"),
    }
    src = example_map.get(args.template)
    if src and os.path.exists(src):
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        shutil.copyfile(src, args.out)
        print(f"Scaffolded '{args.template}' spec -> {args.out}")
        return

    scaffold = {
        "schema_version": 3,
        "universe": "commercial_studio",
        "brand": {"name": "New Presentation", "monogram": "NP"},
        "hero_3d": {"auto_query": "camera lens"},
        "slides": [
            {
                "type": "real_portal_flythrough",
                "stage": "exterior",
                "overline": "CHAPTER 01  //  OVERVIEW",
                "headline": "Spatial Opening",
                "portal_outer_query": "modern glass villa architecture",
                "portal_bg_query": "Matterhorn Zermatt glacier",
                "bottom_cue": "PRECISION OPTICS IN REAL ENVIRONMENTS",
            },
            {
                "type": "odometer_counter",
                "overline": "MEASURED IMPACT",
                "target_value": "1 250 000",
                "sublabel": "UNITS DELIVERED WORLDWIDE",
            },
            {
                "type": "closing_finale",
                "overline": "CONCLUSION",
                "headline": "Built With Real Depth",
                "summary": "High-poly CC0 3D models, open-license photography, and true PowerPoint Morph choreography.",
                "metrics": [
                    {"value": "4K", "label": "PHOTO ASSETS"},
                    {"value": "PBR", "label": "REAL GLB 3D"},
                ],
                "cta": "THANK YOU",
            },
        ],
    }
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(scaffold, f, indent=2)
    print(f"Scaffolded clean Cinema-Morph v3 spec -> {args.out}")


def cmd_fetch_3d(args: argparse.Namespace) -> None:
    if args.auto:
        src, picked, path = resolve_auto_glb(args.auto, args.out)
        print(f"Resolved ({src}: {picked}) -> {path} | stats={inspect_glb_stats(path)}")
    elif args.source == "khronos":
        fetch_khronos_glb(args.name, args.out)
    elif args.source == "polyhaven":
        fetch_polyhaven_glb(args.name, args.out, res=args.res)
    elif args.source == "nasa":
        fetch_nasa_glb(args.name, args.out)
    else:
        raise ValueError("Provide --auto or --source + --name")


def cmd_fetch_photo(args: argparse.Namespace) -> None:
    fetch_wikimedia_photo(args.query, args.out)
    print(f"Fetched open-license photo -> {args.out}")


def cmd_cutout(args: argparse.Namespace) -> None:
    if args.op == "cutout":
        make_studio_cutout(args.src, args.out, tolerance=args.tolerance)
    elif args.op == "portal":
        box = tuple(float(v) for v in args.box.split(","))
        punch_portal_mask(args.src, args.out, box_frac=box)
    elif args.op == "split-sky":
        split_horizon_cutout(args.src, args.out)
    print(f"Studio cutout ({args.op}) -> {args.out}")


def cmd_validate(args: argparse.Namespace) -> None:
    errs, warns = validate_spec(args.spec)
    for w in warns:
        print(f"WARNING: {w}")
    if errs:
        for e in errs:
            print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    with open(args.spec, "r", encoding="utf-8") as f:
        data = json.load(f)
    ver = "v3" if (int(data.get("schema_version", 2)) >= 3 or "universe" in data) else "v2"
    print(f"Validation OK ({ver}, {len(data['slides'])} slides): {args.spec}")


def cmd_preview_3d(args: argparse.Namespace) -> None:
    img = render_glb_to_image(args.glb, args.size, args.size, args.rx, args.ry, args.rz, theme=args.theme)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    img.save(args.out, "PNG")
    print(f"Rendered 3D GLB preview -> {args.out}")


def _can_use_com(engine: str) -> bool:
    if engine == "crossplatform":
        return False
    if engine == "com":
        if not sys.platform.startswith("win"):
            raise RuntimeError("--engine com is only supported on Windows.")
        return True
    if not sys.platform.startswith("win"):
        return False
    ps = shutil.which("powershell.exe") or shutil.which("powershell")
    if not ps:
        return False
    probe = subprocess.run(
        [ps, "-NoProfile", "-Command", "if (Get-ItemProperty HKLM:\\SOFTWARE\\Classes\\PowerPoint.Application -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"],
        capture_output=True,
    )
    return probe.returncode == 0


def cmd_build(args: argparse.Namespace) -> None:
    spec_path = args.spec or args.spec_pos
    if not spec_path:
        raise ValueError("Provide --spec <path> or a positional spec path.")
    with open(spec_path, "r", encoding="utf-8") as f:
        spec_data = json.load(f)
    errs, _ = validate_spec(spec_data)
    if errs:
        for e in errs:
            print(f"Validation error: {e}", file=sys.stderr)
        sys.exit(1)

    is_v3 = int(spec_data.get("schema_version", 2)) >= 3 or "universe" in spec_data
    out_pptx = os.path.abspath(args.out)
    export_dir_arg = args.export_dir or args.renders
    export_dir = os.path.abspath(export_dir_arg) if export_dir_arg else os.path.join(os.path.dirname(out_pptx), "renders")
    use_temp_assets = not bool(args.assets_dir)
    assets_dir = os.path.abspath(args.assets_dir) if args.assets_dir else tempfile.mkdtemp(prefix="morphdeck_assets_")

    try:
        if is_v3:
            resolved = prepare_all(os.path.abspath(spec_path), assets_dir)
            profile = resolved["universe"]
        else:
            resolved = generate_all(os.path.abspath(spec_path), assets_dir)
            profile = resolved["theme"]

        use_com = _can_use_com(args.engine)
        if use_com:
            ps = shutil.which("powershell.exe") or shutil.which("powershell")
            ps1 = os.path.join(SCRIPT_DIR, "build_cinema_deck.ps1" if is_v3 else "build_deck.ps1")
            build_cmd = [
                ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1,
                "-SpecJson", os.path.abspath(spec_path),
                "-AssetsDir", assets_dir,
                "-OutputPptx", out_pptx,
                "-ExportDir", export_dir,
                "-SkipExport",
            ]
            subprocess.run(build_cmd, check=True)
            inject_3d_studio_lighting(out_pptx, theme=profile)

            if not args.skip_export:
                exp_cmd = [
                    ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1,
                    "-OutputPptx", out_pptx,
                    "-ExportDir", export_dir,
                    "-ExportOnly",
                ]
                subprocess.run(exp_cmd, check=True)
        else:
            if is_v3:
                build_cinema_crossplatform(os.path.abspath(spec_path), assets_dir, out_pptx)
            else:
                build_crossplatform_pptx(os.path.abspath(spec_path), assets_dir, out_pptx)

        if not args.skip_verify:
            rc = verify_pptx(
                out_pptx,
                export_dir if (use_com and not args.skip_export) else None,
                require_3d=bool(resolved.get("resolved_glb") or resolved.get("meshes")),
                require_lighting=use_com and bool(resolved.get("resolved_glb") or resolved.get("meshes")),
            )
            if rc != 0:
                sys.exit(rc)
    finally:
        if use_temp_assets and os.path.isdir(assets_dir):
            shutil.rmtree(assets_dir, ignore_errors=True)


def cmd_verify(args: argparse.Namespace) -> None:
    verify_pptx(args.pptx, args.renders_dir, require_3d=args.require_3d, require_lighting=args.require_lighting)


def main() -> None:
    parser = argparse.ArgumentParser(prog="morphdeck", description="morphdeck-3d v3.0.0 CLI: Cinema-Morph & real 3D PowerPoint builder.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_univ = sub.add_parser("universes", help="List v3 Cinema-Morph universes & slide archetypes")
    p_univ.set_defaults(func=cmd_universes)

    p_list = sub.add_parser("list", help="List universes, themes, and archetypes")
    p_list.set_defaults(func=cmd_list)

    p_init = sub.add_parser("init", help="Create a starter presentation JSON spec")
    p_init.add_argument("--out", required=True, help="Output JSON spec path")
    p_init.add_argument("--template", default="documentary_press", help="Starter template name")
    p_init.set_defaults(func=cmd_init)

    p_f3d = sub.add_parser("fetch-3d", help="Fetch a real CC0 high-poly .glb 3D model")
    p_f3d.add_argument("--auto", default="", help="Topic keyword auto-resolver")
    p_f3d.add_argument("--source", choices=["khronos", "polyhaven", "nasa"], default="khronos")
    p_f3d.add_argument("--name", default="")
    p_f3d.add_argument("--res", default="1k")
    p_f3d.add_argument("--out", required=True)
    p_f3d.set_defaults(func=cmd_fetch_3d)

    p_fph = sub.add_parser("fetch-photo", help="Fetch an open-license landscape photo via Wikimedia Commons")
    p_fph.add_argument("--query", required=True)
    p_fph.add_argument("--out", required=True)
    p_fph.set_defaults(func=cmd_fetch_photo)

    p_cut = sub.add_parser("cutout", help="Create cutouts, portal apertures, or sky splits")
    p_cut.add_argument("--op", required=True, choices=["cutout", "portal", "split-sky"])
    p_cut.add_argument("--src", required=True)
    p_cut.add_argument("--out", required=True)
    p_cut.add_argument("--tolerance", type=float, default=28.0)
    p_cut.add_argument("--box", default="0.26,0.17,0.74,0.83")
    p_cut.set_defaults(func=cmd_cutout)

    p_val = sub.add_parser("validate", help="Validate a JSON presentation spec")
    p_val.add_argument("--spec", required=True)
    p_val.set_defaults(func=cmd_validate)

    p_prv = sub.add_parser("preview-3d", help="Render isometric studio preview PNG from a .glb file")
    p_prv.add_argument("--glb", required=True)
    p_prv.add_argument("--out", required=True)
    p_prv.add_argument("--theme", default="commercial_studio")
    p_prv.add_argument("--size", type=int, default=720)
    p_prv.add_argument("--rx", type=float, default=12.0)
    p_prv.add_argument("--ry", type=float, default=-24.0)
    p_prv.add_argument("--rz", type=float, default=2.0)
    p_prv.set_defaults(func=cmd_preview_3d)

    p_bld = sub.add_parser("build", help="Build PPTX presentation + verify XML & renders")
    p_bld.add_argument("spec_pos", nargs="?", default="", help="Optional positional JSON spec path")
    p_bld.add_argument("--spec", default="", help="Path to JSON spec file")
    p_bld.add_argument("--out", required=True)
    p_bld.add_argument("--engine", choices=["auto", "com", "crossplatform"], default="auto")
    p_bld.add_argument("--assets-dir", default="")
    p_bld.add_argument("--export-dir", default="")
    p_bld.add_argument("--renders", default="", help="Alias for --export-dir")
    p_bld.add_argument("--skip-export", action="store_true")
    p_bld.add_argument("--skip-verify", action="store_true")
    p_bld.set_defaults(func=cmd_build)

    p_ver = sub.add_parser("verify", help="Verify PPTX XML integrity and exported slide PNGs")
    p_ver.add_argument("--pptx", required=True)
    p_ver.add_argument("--renders-dir", default="")
    p_ver.add_argument("--require-3d", action="store_true")
    p_ver.add_argument("--require-lighting", action="store_true")
    p_ver.set_defaults(func=cmd_verify)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
