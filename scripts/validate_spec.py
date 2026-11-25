#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
Pre-Flight Deck Spec Linter for morphdeck-3d.

Validates deck_spec.json BEFORE rendering or invoking PowerPoint/OpenXML so
AI agents and CI catch missing fields, invalid hex colors, invalid 3D
archetypes, unicode emojis, and text overflow risks immediately.
"""

import argparse
import json
import os
import re
import sys

VALID_THEMES = {
    "obsidian_kinetic",
    "editorial_ivory",
    "cyber_blueprint",
    "kinetic_poster",
    "executive_clean",
    "biotech_luminance",
    "monochrome_luxury",
}

VALID_ARCHETYPES = {
    "architectural_pavilion",
    "ai_quantum_server",
    "energy_turbine_core",
    "product_drop_capsule",
    "aerospace_orbital_craft",
    "executive_globe_monolith",
    "biotech_helix_core",
    "fintech_vault_prism",
    "smartphone_pro",
    "laptop_studio",
    "iphone",
    "macbook",
    "custom",
}

VALID_SLIDE_TYPES = {
    "cover_hero",
    "portal_flythrough_3d",
    "vertical_elevator_3d",
    "split_horizon_3d",
    "macro_dolly_3d",
    "depth_sandwich_pan",
    "orbit_steps_3d",
    "timeline_morph_step",
    "kpi_bento",
    "chart_infographic",
    "architecture_diagram",
    "problem_solution_split",
    "comparison_matrix",
    "chapter_divider_morph",
    "editorial_split",
    "spotlight_dock",
    "macbook_studio_3d",
    "full_bleed_visual",
    "closing_finale",
}

VALID_PORTAL_STAGES = {"exterior", "interior", "horizon"}
VALID_ELEVATOR_STOPS = {"upper_sky", "lower_ground"}

HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
EMOJI_RE = re.compile("[\U0001F300-\U0001FAFF\u2600-\u27BF\u2300-\u23FF]", flags=re.UNICODE)


def _walk_strings(obj, path="root"):
    if isinstance(obj, str):
        yield path, obj
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            yield from _walk_strings(item, f"{path}[{i}]")
    elif isinstance(obj, dict):
        for k, v in obj.items():
            yield from _walk_strings(v, f"{path}.{k}")


def validate_spec(spec_path: str, allow_emojis: bool = False) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []

    if not os.path.exists(spec_path):
        return [f"Spec file not found: {spec_path}"], warnings

    try:
        with open(spec_path, "r", encoding="utf-8") as f:
            spec = json.load(f)
    except Exception as exc:
        return [f"Invalid JSON in {spec_path}: {exc}"], warnings

    if not isinstance(spec, dict):
        return ["Top-level JSON value must be an object"], warnings

    theme = spec.get("theme", "obsidian_kinetic")
    if theme not in VALID_THEMES:
        errors.append(f"Invalid theme '{theme}'. Must be one of: {sorted(VALID_THEMES)}")

    brand = spec.get("brand", {})
    if not isinstance(brand, dict) or not brand.get("name"):
        errors.append("Missing required brand.name field")

    pal = spec.get("palette", {})
    if isinstance(pal, dict):
        for key in ("accent1", "accent2", "accent3", "chassis_color"):
            if key in pal and pal[key] and not HEX_RE.match(str(pal[key])):
                errors.append(f"Invalid hex color for palette.{key}: '{pal[key]}' (expected #RRGGBB)")

    hero = spec.get("hero_3d", {})
    if isinstance(hero, dict) and hero.get("archetype"):
        arch = str(hero["archetype"]).strip().lower()
        if arch not in VALID_ARCHETYPES:
            errors.append(f"Invalid hero_3d.archetype '{arch}'. Must be one of: {sorted(VALID_ARCHETYPES)}")
        if arch == "custom" and not (hero.get("custom_glb_path") or spec.get("assets", {}).get("custom_glb_path")):
            errors.append("hero_3d.archetype is 'custom' but no custom_glb_path was provided")
    if isinstance(hero, dict) and hero.get("secondary_archetype"):
        sec = str(hero["secondary_archetype"]).strip().lower()
        if sec not in VALID_ARCHETYPES:
            errors.append(f"Invalid hero_3d.secondary_archetype '{sec}'")

    slides = spec.get("slides")
    if slides is None:
        warnings.append("Legacy v1 spec detected (no 'slides' array); prefer schema_version 2 composable slides")
    elif not isinstance(slides, list) or len(slides) == 0:
        errors.append("'slides' must be a non-empty array")
    else:
        for i, sl in enumerate(slides, start=1):
            if not isinstance(sl, dict):
                errors.append(f"slides[{i}] must be a JSON object")
                continue
            st = sl.get("type", "")
            if st not in VALID_SLIDE_TYPES:
                errors.append(f"slides[{i}] invalid type '{st}'. Valid types: {sorted(VALID_SLIDE_TYPES)}")
                continue
            if st in ("split_horizon_3d", "depth_sandwich_pan", "vertical_elevator_3d") and len(str(sl.get("giant_word", ""))) > 10:
                warnings.append(f"slides[{i}] giant_word '{sl.get('giant_word')}' exceeds 10 chars and may wrap or crowd bezels")
            if st == "portal_flythrough_3d" and sl.get("stage") and sl.get("stage") not in VALID_PORTAL_STAGES:
                errors.append(f"slides[{i}] invalid portal_flythrough_3d stage '{sl.get('stage')}' (expected: {sorted(VALID_PORTAL_STAGES)})")
            if st == "vertical_elevator_3d" and sl.get("elevator_stop") and sl.get("elevator_stop") not in VALID_ELEVATOR_STOPS:
                errors.append(f"slides[{i}] invalid vertical_elevator_3d elevator_stop '{sl.get('elevator_stop')}' (expected: {sorted(VALID_ELEVATOR_STOPS)})")
            if st == "macro_dolly_3d" and not sl.get("callouts"):
                errors.append(f"slides[{i}] macro_dolly_3d requires non-empty 'callouts' array")
            if st == "timeline_morph_step" and not sl.get("stages"):
                errors.append(f"slides[{i}] timeline_morph_step requires non-empty 'stages' array")
            if st == "chart_infographic" and not sl.get("items"):
                errors.append(f"slides[{i}] chart_infographic requires non-empty 'items' array")
            if st == "architecture_diagram" and not sl.get("layers"):
                errors.append(f"slides[{i}] architecture_diagram requires non-empty 'layers' array")
            if st == "comparison_matrix" and not sl.get("rows"):
                errors.append(f"slides[{i}] comparison_matrix requires non-empty 'rows' array")

    if not allow_emojis:
        for path, val in _walk_strings(spec):
            if EMOJI_RE.search(val):
                errors.append(f"Unicode emoji found at {path}: {val[:60]}")

    return errors, warnings


def main():
    parser = argparse.ArgumentParser(description="Validate deck_spec.json before building a morphdeck-3d presentation.")
    parser.add_argument("spec", help="Path to deck_spec.json")
    parser.add_argument("--allow-emojis", action="store_true", help="Allow Unicode emojis inside text strings")
    args = parser.parse_args()

    errors, warnings = validate_spec(args.spec, allow_emojis=args.allow_emojis)
    for w in warnings:
        print(f"WARN: {w}")
    for e in errors:
        print(f"ERROR: {e}")

    if errors:
        print(f"SPEC VALIDATION FAILED ({len(errors)} error(s))")
        sys.exit(1)

    print(f"SPEC VALIDATION PASSED cleanly ({len(warnings)} warning(s)): {args.spec}")
    sys.exit(0)


if __name__ == "__main__":
    main()
