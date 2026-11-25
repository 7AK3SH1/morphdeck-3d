---
name: morphdeck-3d
description: >
  Project-adaptive 3D glTF + Morph PowerPoint presentation engineering engine
  for AI coding agents. Generates 10 original procedural PBR 3D .glb archetypes
  tailored to each project domain, 7 themes, 19 composable Morph slide methods
  (optical portal fly-throughs, 2-stop vertical world elevators, 4-layer Depth
  Sandwich, 1.55x macro dolly callouts, backdrop-refracted glass docks),
  3-point colored <am3d:ptLight> studio lighting, and verified Windows COM +
  Cross-Platform OpenXML builds.
license: Apache-2.0
metadata:
  author: Mohamed S. Sissoko (7AK3SH1)
  version: "2.0.0"
  repository: https://github.com/7AK3SH1/morphdeck-3d
---

# morphdeck-3d

Build **100% original, from-scratch 3D + Morph PowerPoint presentations** with zero recycled templates and zero Unicode emojis:
1. **Subject-Linked 3D `.glb` Hero (`hero_3d.archetype`)**: Automatically selects and generates a multi-material PBR 3D glTF 2.0 object directly matching the project's subject (`architectural_pavilion`, `ai_quantum_server`, `energy_turbine_core`, `product_drop_capsule`, `aerospace_orbital_craft`, `executive_globe_monolith`, `biotech_helix_core`, `fintech_vault_prism`, `smartphone_pro`, `laptop_studio`, or `custom`), illuminated with a theme-calibrated 3-point `<am3d:ptLight>` studio rig.
2. **Continuous 3D Morph & Cinema Camera Motion**: Keeps `!!Phone3D` persistent across every slide (via PowerPoint COM `.Duplicate().Item(1)` with `ppEffectMorph` on Windows, or studio PBR renders + ECMA-376 `<mc:AlternateContent><p159:morph/>` XML on macOS/Linux), combining optical portal aperture fly-throughs (`4.6x` bulkhead zoom), `1920x2160` 2-stop vertical world elevators (`-520 pt` camera pan), staggered off-canvas velocity cascades, orthogonal L-trajectory exits, and vertical `srcRect` curtain-wipe image reveals.
3. **4-Layer Depth Sandwich & Backdrop-Refracted Glass**: Layers atmospheric horizons behind giant occluded display type (`!!pan_bg` -> `!!pan_word` -> `!!pan_fg` -> `!!Phone3D`), paired with true Gaussian backdrop-refracted frosted glass docks and contact shadows.
4. **Distinct Architecture Per Project**: Never reuse the exact same slide sequence across two projects. Pick the archetype, theme, slide count (6-8 slides), and 3D Morph choreography to fit the domain.

---

## Step 0: Project Diagnosis Matrix (Archetype + Theme + Choreography)

| Project Domain | `hero_3d.archetype` | `theme` | Distinct 3D + Morph Slide Choreography |
|---|---|---|---|
| **Architecture, Real Estate, Eco-Design** | `architectural_pavilion` | `editorial_ivory` | `cover_hero(asymmetric_split)` -> `split_horizon_3d` -> `macro_dolly_3d` (1.55x dolly + callouts) -> `timeline_morph_step` x3 -> `closing_finale(editorial_statement)` |
| **AI, Cloud, Cybersecurity, Backend** | `ai_quantum_server` (+ `laptop_studio`) | `cyber_blueprint` | `cover_hero(centered_clean)` -> `kpi_bento(left_anchor)` -> `architecture_diagram` -> `chart_infographic` x2 -> `macbook_studio_3d` -> `closing_finale(dual_3d)` |
| **Biotech, Genomics, Pharma, MedTech** | `biotech_helix_core` | `biotech_luminance` | `cover_hero(centered_clean)` -> `depth_sandwich_pan(wide)` -> `depth_sandwich_pan(zoom_right)` -> `macro_dolly_3d` -> `chart_infographic(radial_rings)` -> `closing_finale` |
| **CleanTech, Energy, Solar, Smart Grids** | `energy_turbine_core` | `obsidian_kinetic` | `cover_hero(asymmetric_split)` -> `split_horizon_3d` -> `macro_dolly_3d` -> `chart_infographic(comparison_bars)` -> `timeline_morph_step` x2 -> `closing_finale` |
| **FinTech, Security, Cryptography** | `fintech_vault_prism` | `monochrome_luxury` | `cover_hero(centered_clean)` -> `problem_solution_split` -> `architecture_diagram` -> `comparison_matrix` -> `closing_finale(executive_cta)` |
| **Product Launch, Retail Drops** | `product_drop_capsule` | `kinetic_poster` | `cover_hero(bold_poster)` -> `problem_solution_split` -> `kpi_bento(top_hero_strip)` -> `spotlight_dock` x2 (refracted dock + `srcRect` wipe + L-exit) -> `closing_finale(editorial_statement)` |
| **FinOps, Corporate Strategy, Consulting** | `executive_globe_monolith` | `executive_clean` | `cover_hero(asymmetric_split)` -> `chapter_divider_morph` -> `kpi_bento(top_hero_strip)` -> `chart_infographic` -> `comparison_matrix` -> `closing_finale(executive_cta)` |
| **Aerospace, Mobility, Drones, Defense** | `aerospace_orbital_craft` | `obsidian_kinetic` | `portal_flythrough_3d(exterior)` -> `portal_flythrough_3d(interior)` -> `vertical_elevator_3d(upper_sky)` -> `vertical_elevator_3d(lower_ground)` -> `macro_dolly_3d` -> `closing_finale` |
| **Mobile + Web Consumer App / SaaS** | `smartphone_pro` (+ `laptop_studio`) | `obsidian_kinetic` | `cover_hero(3d_staged)` -> `split_horizon_3d` -> `orbit_steps_3d(upper_stats)` -> `orbit_steps_3d(obsidian_steps)` -> `macbook_studio_3d` -> `closing_finale(dual_3d)` |

---

## Hard Bans

- **Never use `smartphone_pro`/`laptop_studio` as the sole hero for non-software topics** (e.g., architecture, biotech, clean energy). Always link `hero_3d.archetype` to the project subject.
- **Never skip verification**: always run `validate_spec.py` before build and `verify_deck.py` after build (`0` XML errors, `0` external links, `<mc:AlternateContent>` + `<p159:morph>` on every slide, `0` emojis).
- **Never insert Unicode emojis** inside JSON specs, headings, labels, or `.pptx` text runs.
- **Photorealistic Photography Rule**: Whenever project photos exist or an image generation tool (`image-generator`) is available, populate `assets.horizon_photo`, `assets.upper_photo`, and `spotlights[*].portrait_image` with real 16:9 visuals so `render_studio_assets.py` builds true photographic Split-Horizons and 4-Layer Depth Sandwiches.

---

## Turnkey Workflow

Discover all capabilities first if needed:
```bash
python scripts/morphdeck.py list
```

### Step 1: Initialize & Author `deck_spec.json`
```bash
python scripts/morphdeck.py init --template ai --output work/deck_spec.json
```
Edit `work/deck_spec.json` to match the project topic, character limits in [references/themes-and-methods-catalog.md](references/themes-and-methods-catalog.md), and real/generated photography paths.

### Step 2: Validate, Render, Build & Verify in One Command
```bash
python scripts/morphdeck.py build work/deck_spec.json --out Output_Presentation.pptx --renders work/renders
```
- On **Windows with Microsoft PowerPoint**, `--engine auto` uses PowerPoint COM (`scripts/build_deck.ps1 -SkipExport`), injects a theme-calibrated 3-point `<am3d:ptLight>` studio lighting rig via `scripts/inject_3d_studio_lighting.py`, re-opens via COM (`-ExportOnly`) to export `1920x1080` lit slide PNGs into `work/renders`, and enforces `GLBs >= 1` via `verify_deck.py`.
- On **macOS / Linux / CI** (or with `--engine crossplatform`), `scripts/build_deck_crossplatform.py` + `scripts/render_glb_preview.py` renders PBR studio views for every 3D camera angle and packages an OpenXML-compliant Morph `.pptx`.

### Step 3: Visual PNG Inspection
When running on Windows COM with exported renders (`work/renders/slide_*.png`), visually inspect the rendered PNGs before final delivery to confirm zero text clipping, high contrast, and crisp 3D framing.

### Standalone 3D Tools
- Preview any archetype in seconds without PowerPoint:
  ```bash
  python scripts/morphdeck.py preview-3d --archetype biotech_helix_core --brand "HELIXIS" --accent1 "#10B981" --output preview.png
  ```
- Bake a UI screenshot onto any external custom `.glb`:
  ```bash
  python scripts/bake_glb_screen.py --input custom.glb --mesh DisplayMesh --material 2 --image ui.png --output baked.glb
  ```
