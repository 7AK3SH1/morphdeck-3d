# Changelog

All notable changes to **morphdeck-3d** will be documented in this file.

## [2.0.0] - 2026-11-25 - Cinema Motion & Studio Optics

### Added
- **Slide Method #18: `portal_flythrough_3d`** (`exterior` -> `interior` -> `horizon`): Procedural dual-aperture optical bulkhead masks (`portal_outer.png`, `portal_inner.png`) with punched chamfered transparent viewports and HUD brackets that zoom `4.2x-4.6x` past the camera under Morph while the live 3D `.glb` subject flies forward through the bay.
- **Slide Method #19: `vertical_elevator_3d`** (`upper_sky` -> `lower_ground`): Seamless `1920x2160` double-height atmospheric panorama (`elev_bg.png`) paired with a sculpted ridge cutout (`elev_fg.png`) that travels `-520 pt` vertically across two camera stops while keeping giant occluded typography behind the horizon ridge and the live 3D subject orbiting in front.
- **Theme-Calibrated 3-Point 3D Studio Lighting Injector (`scripts/inject_3d_studio_lighting.py`)**: Automatically upgrades COM-embedded `<am3d:model3d>` XML nodes across all 7 themes with calibrated `<am3d:ambientLight>` + 3 colored `<am3d:ptLight>` emitters (warm key + theme accent fill + colored rim backlight), synchronized with `scripts/render_glb_preview.py` for cross-platform parity.
- **True Backdrop-Refracted Glassmorphism (`render_refracted_glass_dock`)**: Crops and Gaussian-blurs (`r=24`) the exact underlying slide background coordinates before compositing frosted tint glass and top specular rims onto `spotlight_dock` (`s*_dock_refract.png`).
- **Turnkey Aerospace Example (`examples/aerospace_portal_cinema.json`)**: Demonstrates `portal_flythrough_3d`, `vertical_elevator_3d`, staggered `macro_dolly_3d`, and `closing_finale` with `aerospace_orbital_craft`.

### Improved
- **Staggered-Velocity Morph & `srcRect` Curtain-Wipe Reveals**: `spotlight_dock`, `macro_dolly_3d`, `portal_flythrough_3d`, and `vertical_elevator_3d` pre-stage incoming actors at staggered off-canvas coordinates for natural velocity cascades, exit previous `spotlight_dock` typography orthogonally upward, morph pagination dot dimensions (`7pt` vs `14pt`), and animate vertical `CropTop`/`CropBottom` (`<a:srcRect>`) curtain wipes on incoming spotlight backdrops.
- **COM Export & Lighting Pipeline**: Added `-SkipExport` / `-ExportOnly` flags to `scripts/build_deck.ps1` and `<am3d:ptLight` telemetry reporting to `scripts/verify_deck.py` so exported `1920x1080` PNGs render with full 3-point colored 3D studio illumination.

## [1.0.0] - 2026-11-25

### Added
- **10 Procedural PBR 3D `.glb` Archetypes** (`scripts/generate_subject_glb.py`):
  - `architectural_pavilion`, `ai_quantum_server`, `energy_turbine_core`, `product_drop_capsule`, `aerospace_orbital_craft`, `executive_globe_monolith`, `biotech_helix_core`, `fintech_vault_prism`, `smartphone_pro`, and `laptop_studio` (100% original geometry, zero third-party binaries).
- **7 Built-in Visual Themes** (`scripts/render_studio_assets.py`):
  - `obsidian_kinetic`, `editorial_ivory`, `cyber_blueprint`, `kinetic_poster`, `executive_clean`, `biotech_luminance`, `monochrome_luxury`.
- **17 Composable Slide Methods**:
  - Including `macro_dolly_3d` (1.55x tactical 3D camera dolly with leader callouts), 4-layer `depth_sandwich_pan` (`!!pan_bg` -> `!!pan_word` -> `!!pan_fg` -> `!!Phone3D`), `split_horizon_3d`, `orbit_steps_3d`, `timeline_morph_step`, `kpi_bento`, `chart_infographic`, `architecture_diagram`, `problem_solution_split`, `comparison_matrix`, `chapter_divider_morph`, `editorial_split`, `spotlight_dock`, `macbook_studio_3d`, `full_bleed_visual`, and `closing_finale`.
- **Dual Build Engines**:
  - Windows Native PowerPoint COM Builder (`scripts/build_deck.ps1`) embedding `.glb` once and morphing across every slide via `.Duplicate().Item(1)`.
  - Cross-Platform OpenXML Builder (`scripts/build_deck_crossplatform.py`) + Pure-NumPy Studio GLB Renderer (`scripts/render_glb_preview.py`) for macOS, Linux, and headless CI.
- **Unified Turnkey CLI & Validators**:
  - `scripts/morphdeck.py` (`list`, `init`, `validate`, `preview-3d`, `build`, `verify`), `scripts/validate_spec.py`, `scripts/verify_deck.py`, and `scripts/bake_glb_screen.py`.
- **Multi-Agent Installers & Verification**:
  - `install.ps1`, `install.sh`, `tests/test_morphdeck.py`, GitHub Actions CI, Apache-2.0 `LICENSE`, and mandatory `NOTICE` attribution.
