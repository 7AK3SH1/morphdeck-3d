#!/usr/bin/env python3
# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""
JSON Schema Validator & Anti-Regression Gate for morphdeck-3d (v2 & Cinema-Morph v3).

Validates both:
  - v3 Cinema-Morph specs (`universe`, real `.glb`/photorealistic `hero_3d`, 17 slide types)
  - v2 Legacy specs (`theme`, 10 archetypes + custom, 19 legacy slide types)
Accepts either a file path or a parsed dict and returns `(errors, warnings)`.
Rejects missing required fields, malformed hex colors, invalid slide counts, and emojis.
"""

import argparse
import json
import os
import re
import sys

VALID_UNIVERSES = {
    "documentary_press",
    "architectural_portal",
    "sculptural_museum",
    "commercial_studio",
    "executive_swiss",
    "editorial_pop",
}

VALID_CINEMA_SLIDE_TYPES = {
    "newspaper_dive",
    "editorial_print_overlay",
    "odometer_counter",
    "flyby_wipe",
    "cutout_map_stage",
    "poster_depth_reveal",
    "picture_frame_pullback",
    "real_portal_flythrough",
    "refracted_spotlight_dock",
    "parchment_3d_split",
    "vertical_world_elevator",
    "rotational_drift_3d",
    "swiss_bento_morph",
    "timeline_rail_morph",
    "chart_infographic",
    "architecture_blueprint",
    "closing_finale",
}

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
    "custom",
}

VALID_SLIDE_TYPES = {
    "cover_hero",
    "portal_flythrough_3d",
    "vertical_elevator_3d",
    "split_horizon_3d",
    "depth_sandwich_pan",
    "macro_dolly_3d",
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

HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
EMOJI_RE = re.compile(r"[\U0001F300-\U0001FAFF\u2600-\u27BF\u2300-\u23FF]")


def _check_no_emoji(obj, path: str, errors: list[str]) -> None:
    if isinstance(obj, str):
        if EMOJI_RE.search(obj):
            errors.append(f"Emoji detected at '{path}': remove emojis and use typographic labels.")
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            _check_no_emoji(item, f"{path}[{i}]", errors)
    elif isinstance(obj, dict):
        for k, v in obj.items():
            _check_no_emoji(v, f"{path}.{k}", errors)


def validate_spec(spec_or_path) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []

    if isinstance(spec_or_path, (str, os.PathLike)):
        if not os.path.exists(spec_or_path):
            return ([f"Spec file not found: {spec_or_path}"], warnings)
        with open(spec_or_path, "r", encoding="utf-8") as f:
            spec = json.load(f)
    else:
        spec = spec_or_path

    if not isinstance(spec, dict):
        return (["Spec root must be a JSON object."], warnings)

    _check_no_emoji(spec, "root", errors)
    slides = spec.get("slides")
    if not isinstance(slides, list) or not (1 <= len(slides) <= 24):
        errors.append("'slides' must be a non-empty list of 1 to 24 slides.")
        return (errors, warnings)

    pal = spec.get("palette", {})
    if isinstance(pal, dict):
        for ck, cv in pal.items():
            if ck == "is_light":
                continue
            if cv and not HEX_RE.match(str(cv)):
                errors.append(f"palette.{ck} invalid hex color '{cv}' (expected #RRGGBB).")

    is_v3 = int(spec.get("schema_version", 2)) >= 3 or "universe" in spec
    if is_v3:
        univ = spec.get("universe")
        if univ and univ not in VALID_UNIVERSES:
            errors.append(f"Invalid 'universe' '{univ}'. Valid: {sorted(VALID_UNIVERSES)}")

        for i, sl in enumerate(slides, start=1):
            if not isinstance(sl, dict):
                errors.append(f"Slide {i} must be an object.")
                continue
            st = sl.get("type")
            if st not in VALID_CINEMA_SLIDE_TYPES:
                errors.append(f"Slide {i} invalid v3 type '{st}'. Valid: {sorted(VALID_CINEMA_SLIDE_TYPES)}")
                continue
            if st == "odometer_counter" and not sl.get("target_value"):
                errors.append(f"Slide {i} (odometer_counter) requires 'target_value'.")
            elif st == "refracted_spotlight_dock":
                if not sl.get("spotlight") or not sl.get("dock"):
                    errors.append(f"Slide {i} (refracted_spotlight_dock) requires 'spotlight' and 'dock'.")
            elif st == "timeline_rail_morph" and not sl.get("stages"):
                errors.append(f"Slide {i} (timeline_rail_morph) requires 'stages'.")
            elif st == "chart_infographic" and not sl.get("items"):
                errors.append(f"Slide {i} (chart_infographic) requires 'items'.")
            elif st == "architecture_blueprint" and not sl.get("layers"):
                errors.append(f"Slide {i} (architecture_blueprint) requires 'layers'.")
        return (errors, warnings)

    # Legacy v2 validation
    theme = spec.get("theme")
    if theme and theme not in VALID_THEMES:
        errors.append(f"Invalid 'theme' '{theme}'. Valid: {sorted(VALID_THEMES)}")

    hero = spec.get("hero_3d", {})
    if isinstance(hero, dict) and hero.get("archetype") and hero["archetype"] not in VALID_ARCHETYPES:
        errors.append(f"Invalid 'hero_3d.archetype' '{hero['archetype']}'. Valid: {sorted(VALID_ARCHETYPES)}")

    for i, sl in enumerate(slides, start=1):
        if not isinstance(sl, dict):
            errors.append(f"Slide {i} must be an object.")
            continue
        st = sl.get("type")
        if st not in VALID_SLIDE_TYPES:
            errors.append(f"Slide {i} invalid type '{st}'. Valid: {sorted(VALID_SLIDE_TYPES)}")
            continue
        if st == "portal_flythrough_3d" and sl.get("stage") and sl["stage"] not in ("exterior", "interior", "horizon"):
            errors.append(f"Slide {i} (portal_flythrough_3d) invalid stage '{sl['stage']}'.")
        if st == "vertical_elevator_3d" and sl.get("elevator_stop") and sl["elevator_stop"] not in ("upper_sky", "lower_ground"):
            errors.append(f"Slide {i} (vertical_elevator_3d) invalid elevator_stop '{sl['elevator_stop']}'.")
        if st == "chart_infographic" and not sl.get("items"):
            errors.append(f"Slide {i} (chart_infographic) requires 'items'.")
        elif st == "architecture_diagram" and not sl.get("layers"):
            errors.append(f"Slide {i} (architecture_diagram) requires 'layers'.")
    return (errors, warnings)


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate a morphdeck-3d JSON specification.")
    parser.add_argument("--spec", required=True, help="Path to JSON spec file")
    args = parser.parse_args()

    errs, warns = validate_spec(args.spec)
    for w in warns:
        print(f"WARNING: {w}")
    if errs:
        print("VALIDATION FAILED:", file=sys.stderr)
        for e in errs:
            print(f"  - {e}", file=sys.stderr)
        sys.exit(1)
    with open(args.spec, "r", encoding="utf-8") as f:
        data = json.load(f)
    ver = "v3" if (int(data.get("schema_version", 2)) >= 3 or "universe" in data) else "v2"
    print(f"Spec validation passed ({ver}, {len(data['slides'])} slides): {args.spec}")


if __name__ == "__main__":
    main()
