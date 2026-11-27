# Changelog

All notable changes to **morphdeck-3d** will be documented in this file.

## [5.0.0] - 2026-11-27 - Optical Scrim Precision, Depth-Sandwich Framing & Borderless Glass Docks

### Improved & Fixed
- **Generous 3D Camera Framing (`scripts/render_glb_preview.py`)**: Expanded GLB projection distance (`cam_dist = span * 2.75`, normalized `w >= h` frames across `poster_depth_reveal`, `rotational_drift_3d`, `swiss_bento_morph`, `chart_infographic`, `timeline_rail_morph`, and `closing_finale`) so wide hardware models like `MaterialsVariantsShoe` never clip heel or toe edges.
- **Behind-Subject Depth Sandwich & Legible Gallery Placards (`poster_depth_reveal`, `picture_frame_pullback`)**: Enlarged `!!poster_word` to `152pt-184pt` layered behind `!!Hero3D` / `!!poster_hero` so typography interlocks physically behind the 3D foreground, and upgraded left gallery subtitles to high-contrast `#CBD5E1` slate with drop-shadow depth.
- **Borderless Backdrop-Refracted Spotlight Dock (`refracted_spotlight_dock`)**: Replaced hard outlined pill boxes with soft vertical-fade frosted frost gliders (`render_soft_dock_glider`), crisp `16:9` destination thumbnails, `1.2px` white destination underlines (`!!DockLine*`), dynamic 3-or-4-item rail math, and synchronized `<a:srcRect t="-100000" b="100000"/>` curtain wipes across both COM and cross-platform OpenXML engines.
- **3-Stage Portal Continuity (`real_portal_flythrough`) & Frosted Elevator Triptych (`vertical_world_elevator`)**: Full `exterior -> interior -> horizon` transition bridge (`!!pt_hor_title`/`!!pt_hor_body`) plus horizon-reflected frost triptych (`elev_rcard_h_*`) directly sliced from the `1920x2160` alpine panorama (`use_hero_3d: false` option removes mid-air hardware over cliffs).
- **Feathered Odometer Slot Masks & Non-Periodic Paper Deckle Edges (`scripts/studio_cutout.py`)**: Upgraded odometer reels to `90pt` / `112pt` rows with smooth gradient-feathered slot apertures (`render_odometer_masks`), replaced periodic `sin()` waves with Gaussian-smoothed Brownian deckle paper edges, and added optical top/bottom-left scrims on bright landscapes (`flyby_wipe`, `portal_bg`).
- **Fixed**: Removed trailing comma bug in `scripts/build_cinema_deck.ps1` cross-platform/skill bundle `timeline_rail_morph` pre-staging arguments.

## [4.0.0] - 2026-11-26 - True-Morph Physics & Museum-Grade 3D Optics

### Added
- **Unbroken True-Morph Pre/Post-Staging & Optical Alpha Dissolves (`<a:alphaModFix amt="0"/>`)**: Unified across both `scripts/build_cinema_deck.ps1` (Windows COM) and `scripts/build_cinema_crossplatform.py` (macOS/Linux/CI) so every adjacent slide pair animates `>= 2` shared `!!` Morph actors with staggered velocity curves, pre-masked odometer strips, and live `<am3d:model3d>` concentric frame-shrink.
- **True Cantilevered Infinity-Pool Architectural Portals & Mullioned Interiors (`render_architectural_portal_exterior`, `render_architectural_portal_interior`)**: Synthesizes perspective architectural pavilions with live water-rippled reflections of the 4K background landscape and floor-to-ceiling steel mullions (replacing rectangular sticker punches).
- **Cloud-Proof 1D Dynamic-Programming Sky-Split Horizon Matting (`split_horizon_cutout`)**: Guarantees `100%` cloud-free upper sky (`alpha = 0`) behind giant `132pt` display typography while preserving crisp mountain silhouettes below.
- **2x-SSAA Z-Buffered Smooth-Phong & Bilinear-PBR `.glb` Renderer (`scripts/render_glb_preview.py`)**: Full glTF 2.0 scene-graph node traversal (`TRS` quaternions + `4x4` matrices), floating-point Z-buffer, smooth vertex-normal interpolation, bilinear PBR texture lookup, and 2x supersampled anti-aliasing.
- **Juried Wikimedia Featured & Quality Photography Selector (`scripts/fetch_real_photo.py`)**: Prioritizes `Featured_pictures_on_Wikimedia_Commons` and `Quality_images` while rejecting tourist crowds and photo montages.
- **Museum Gallery Spotlight Wall, Mitered Gilt Frame, Deckle Vellum & Optical ND Scrims (`scripts/studio_cutout.py`)**: Eliminates clunky opaque UI boxes in favor of optical exposure gradients (`apply_tonal_text_scrim`), tourist-free parabolic-spotlight museum walls (`render_gallery_wall_room`), 3D mitered gilt frames, and non-repetitive archival broadsheet articles.
- **Automated Morph Pair Continuity Gate (`verify_deck.py --require-morph-pairs`)**: Fails verification if any slide transition in a v3/v4 presentation has zero moving `!!` Morph actors.

## [3.0.0] - 2026-11-25 - Cinema-Morph & Photorealistic 3D

### Added
- **6 Cinema-Morph Visual Universes**: `documentary_press`, `architectural_portal`, `sculptural_museum`, `commercial_studio`, `executive_swiss`, and `editorial_pop`.
- **17 Signature Morph & Optical Slide Engines** (`scripts/prepare_cinema_assets.py` + `scripts/build_cinema_deck.ps1` + `scripts/build_cinema_crossplatform.py`):
  - `newspaper_dive` (`frontpage` -> `masthead_spread` -> `collage`), `editorial_print_overlay` (organic liquid ink drips via `render_organic_ink_drips` + feathered paper cloud reading window + 3D cutout), `odometer_counter` (black curtain wipe + masked multi-column vertical digit reels + crimson horizon silhouette), `flyby_wipe` (diagonal volumetric cloud sweep + contrast scrim + rotating object), `cutout_map_stage` (gliding cartographic map + KPI panel), `poster_depth_reveal` (dropping accent monolith + behind-subject giant display year), and `picture_frame_pullback` (true `4.4x` concentric camera pull-back shrinking the live poster into a hollow transparent-aperture museum gold frame).
  - Real photography `real_portal_flythrough`, `refracted_spotlight_dock`, `parchment_3d_split`, `vertical_world_elevator` (`1920x2160` sky-split ascent/descent), `rotational_drift_3d`, `swiss_bento_morph`, `timeline_rail_morph`, `chart_infographic`, `architecture_blueprint`, and `closing_finale`.
- **Real High-Poly CC0 `.glb` Resolver (`scripts/fetch_real_glb.py`)**: Resolves authentic PBR `.glb` geometry from Khronos `glTF-Sample-Assets`, Poly Haven CC0 photogrammetry (packed directly to binary `.glb`), and NASA 3D Resources instead of synthetic toy primitives.
- **Open-License Photography Fetcher & Studio Cutout Engine (`scripts/fetch_real_photo.py` + `scripts/studio_cutout.py`)**: Fetches real 4K architecture, gallery, alpine, and historical cartography via Wikimedia Commons API (`fetch_wikimedia_photo`), plus perimeter flood-fill studio segmentation, architectural portal punching, sky splitting, organic ink drips, soft paper clouds, cloud banks, horizon silhouettes, and hollow museum frames.
- **Expanded Turnkey CLI (`scripts/morphdeck.py`)**: Added `universes`, `fetch-3d`, `fetch-photo`, `cutout`, unified v2/v3 `build` routing, and `--require-lighting` OpenXML QA verification.

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
