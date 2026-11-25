# Changelog

All notable changes to **morphdeck-3d** will be documented in this file.

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
