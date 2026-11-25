---
name: morphdeck-3d
description: >
  Cinema-Morph PowerPoint engineering skill for AI coding agents across any
  domain (historical documentaries, luxury architecture, consumer hardware,
  SaaS, biotech, executive decks). Features 6 Visual Universes, 17 native
  PowerPoint Morph engines (broadsheet newspaper dives, organic ink drips,
  rolling odometer counters, architectural portal fly-throughs, diagonal cloud
  fly-by wipes, 4.4x concentric picture-frame pull-backs, sky-split vertical
  world elevators, backdrop-refracted glass docks), real high-poly CC0 .glb
  models (Khronos, Poly Haven, NASA), and open-license 4K photography via
  Wikimedia Commons. Zero toy 3D primitives. Zero emojis.
license: Apache-2.0
metadata:
  author: Mohamed S. Sissoko (7AK3SH1)
  version: "3.0.0"
  repository: https://github.com/7AK3SH1/morphdeck-3d
---

# morphdeck-3d v3.0.0 — Cinema-Morph & Photorealistic 3D

Build **authentic, domain-adaptive Cinema-Morph PowerPoint presentations** with zero generic SaaS templates, zero blocky primitive toy 3D, zero HDRI calibration spheres, and zero Unicode emojis:

1. **6 Distinct Visual Universes**: Pick the palette, typography, paper/obsidian depth, and lighting profile from the project's true subject (`documentary_press`, `architectural_portal`, `sculptural_museum`, `commercial_studio`, `executive_swiss`, `editorial_pop`).
2. **17 Native PowerPoint Morph Engines**: Compose distinct narrative arcs per deck:
   - *Editorial & Documentary*: `newspaper_dive`, `editorial_print_overlay` (organic liquid ink drips + feathered paper cloud), `odometer_counter` (black curtain wipe + masked rolling digit reels + silhouette horizon), `flyby_wipe` (diagonal volumetric cloud sweep), `cutout_map_stage`, `poster_depth_reveal`, `picture_frame_pullback` (true `4.4x` concentric hollow-frame wall pull-back).
   - *Spatial, Product & Technical*: `real_portal_flythrough`, `refracted_spotlight_dock`, `parchment_3d_split`, `vertical_world_elevator` (`1920x2160` sky-split ascent/descent), `rotational_drift_3d`, `swiss_bento_morph`, `timeline_rail_morph`, `chart_infographic`, `architecture_blueprint`, `closing_finale`.
3. **Real High-Poly CC0 `.glb` + 4K Open-License Photography**:
   - Use `scripts/fetch_real_glb.py` (`Khronos glTF-Sample-Assets`, `Poly Haven` glTF->GLB packer, `NASA 3D Resources`) or GLB-rendered RGBA cutouts (`*_cutout_glb`) instead of blocky toy primitives. Reserve `smartphone_pro`/`laptop_studio` strictly for software UI screens.
   - Use `scripts/fetch_real_photo.py` (Wikimedia Commons API queries) or user-supplied imagery for real portals, maps, galleries, and alpine horizons.

---

## Step 0: Domain-to-Universe Routing Matrix

| Project Domain | Recommended `universe` | Real 3D Source (`hero_3d`) | Signature Slide Sequence (Vary Per Project!) |
|---|---|---|---|
| **History, Documentary, Investigative, Media** | `documentary_press` | `khronos/AntiqueCamera` or `nasa/...` | `newspaper_dive` x2 -> `editorial_print_overlay` -> `odometer_counter` -> `flyby_wipe` -> `cutout_map_stage` -> `poster_depth_reveal` -> `picture_frame_pullback` |
| **Luxury Architecture, Optics, Real Estate** | `architectural_portal` | `khronos/AntiqueCamera` or `polyhaven/marble_bust_01` | `real_portal_flythrough(exterior)` -> `real_portal_flythrough(interior)` -> `parchment_3d_split` -> `vertical_world_elevator` x2 -> `refracted_spotlight_dock` -> `closing_finale` |
| **Footwear, Hardware, Automotive, Apparel** | `commercial_studio` | `khronos/MaterialsVariantsShoe` / `CarConcept` / `BoomBox` | `poster_depth_reveal` -> `rotational_drift_3d` -> `swiss_bento_morph` -> `chart_infographic` -> `timeline_rail_morph` -> `closing_finale` |
| **Museums, Culture, Luxury Brand Storytelling** | `sculptural_museum` | `khronos/DragonAttenuation` or `polyhaven` scan | `poster_depth_reveal` -> `picture_frame_pullback` -> `parchment_3d_split` -> `refracted_spotlight_dock` -> `closing_finale` |
| **Enterprise Board, FinOps, Strategy** | `executive_swiss` | `khronos/ABeautifulGame` or `SheenChair` | `swiss_bento_morph` -> `timeline_rail_morph` -> `chart_infographic` -> `architecture_blueprint` -> `closing_finale` |
| **Consumer Creators, Music, Gaming Drops** | `editorial_pop` | `khronos/DamagedHelmet` or `BoomBox` | `editorial_print_overlay` -> `odometer_counter` -> `flyby_wipe` -> `rotational_drift_3d` -> `closing_finale` |

Full schema & field specifications: [references/universes-and-transitions-bible.md](references/universes-and-transitions-bible.md).

---

## Turnkey Workflow

```bash
# 1. Inspect universes & slide engines
python scripts/morphdeck.py universes

# 2. Fetch or test a real CC0 .glb or 4K Wikimedia photograph
python scripts/morphdeck.py fetch-3d --auto "sneaker" --out work/hero.glb
python scripts/morphdeck.py fetch-photo --query "Louvre gallery hall" --out work/wall.png

# 3. Validate JSON spec (v2 or v3)
python scripts/morphdeck.py validate --spec work/deck_spec.json

# 4. Build PPTX + inject 3-point <am3d:ptLight> + export 1080p PNGs + verify XML
python scripts/morphdeck.py build --spec work/deck_spec.json --out Output.pptx --export-dir work/renders
```

- **Windows + PowerPoint**: Uses `scripts/build_cinema_deck.ps1` COM automation, injects 3-point `<am3d:ptLight>` studio rigs via `scripts/inject_3d_studio_lighting.py`, exports crisp `1920x1080` PNGs via `-ExportOnly`, and runs `verify_deck.py --require-lighting`. Always inspect the rendered PNGs before finishing.
- **macOS / Linux / CI**: Automatically uses `scripts/build_cinema_crossplatform.py` + `scripts/render_glb_preview.py` to produce valid ECMA-376 Morph OpenXML `.pptx` packages.
