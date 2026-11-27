# Cinema-Morph v5 Visual Universes & True-Morph Reference

Detailed contract for `morphdeck-3d` v5.0.0 (`schema_version: 3`, `4`, and `5`).

## 1. Visual Universes (`universe`)

- **`documentary_press`** *(light)* — Archival warm broadsheet paper `#EAE3D2`, crisp crimson `#B91C1C`, charcoal ink `#161618`, serif Georgia display typography.
- **`architectural_portal`** *(light)* — Warm travertine `#F4EFE6`, burnt sienna `#C2410C`, teal `#0D9488`, Georgia + Bahnschrift editorial hierarchy.
- **`sculptural_museum`** *(dark)* — Midnight gallery obsidian `#0A0C12`, terracotta `#E07A5F`, museum gold `#D4AF37`, high-contrast spotlight mood.
- **`commercial_studio`** *(dark)* — Deep studio black `#080A0F`, crimson `#E11D48`, electric cyan `#38BDF8`, bold Impact + Bahnschrift kinetic typography.
- **`executive_swiss`** *(light)* — Pure Swiss slate `#F8FAFC`, royal blue `#2563EB`, tight geometric borders `#CBD5E1`, data-first readability.
- **`editorial_pop`** *(dark)* — Deep violet `#0F0C18`, magenta `#EC4899`, cyan `#06B6D4`, amber `#FBBF24`, expressive poster scale.

## 2. Real 3D & Open-License Photography Resolution

### `hero_3d` (Deck-Wide Live `<am3d:model3d>` Asset)
- Explicit Khronos GLB: `{"source": "khronos", "name": "MaterialsVariantsShoe"}` *(reliable studio opaque PBR: `AntiqueCamera`, `MaterialsVariantsShoe`, `DamagedHelmet`, `GlamVelvetSofa`, `SheenChair`, `BoomBox`, `Lantern`, `CarConcept`, `WaterBottle`; avoid `ChronographWatch` in COM due to glass transmission rendering white)*.
- Poly Haven Photogrammetry (packed to `.glb`): `{"source": "polyhaven", "name": "marble_bust_01", "res": "1k"}`
- NASA 3D Spacecraft: `{"source": "nasa", "name": "Voyager"}`
- Automatic keyword resolver: `{"auto_query": "vintage camera lens"}`
- Custom user `.glb`: `{"glb_path": "assets/custom.glb"}`

### Photography Fields (`*_query`, `*_url`, or local path)
Every visual slot (`bg_query`, `portal_outer_query`, `portal_inner_query`, `portal_bg_query`, `wall_query`, `framed_art_query`, `poster_bg_query`, `lower_query`, `map_query`, `panorama_query`, `archival_queries`) resolves automatically through `scripts/fetch_real_photo.py` (Wikimedia Commons API landscape search, >= 1200 px wide, progressive short-query fallback, local cache). Keep queries concise (`"Matterhorn peak snow"`, `"Louvre Pyramid courtyard"`, `"Theatrum Orbis Terrarum world map"`, `"Louvre gallery hall painting room"`). Never use HDRI preview images.

## 3. All 17 Cinema-Morph v4 Slide Types

| Slide `type` | Signature Transition & Visual Grammar | Key JSON Fields |
|---|---|---|
| `newspaper_dive` | Broadsheet newspaper (`frontpage` -> `masthead_spread` 2.5x zoom -> `collage` 3.6x zoom) | `stage`, `newspaper_headline`, `newspaper_subhead`, `dateline`, `columns`, `archival_queries`, `badge` |
| `editorial_print_overlay` | Zoomed newsprint + deckle-edge archival tears (`ep_drips.png`) + vellum reading sheet (`ep_paper_cloud.png`) + cutout | `overline`, `headline`, `body`, `portrait_cutout` or `portrait_cutout_glb`, `badge` |
| `odometer_counter` | Pre-masked black curtain sweep (`!!odo_mask_top`/`!!odo_mask_bot`) + rolling digit reels + 3-layer twilight horizon | `bg_hex`, `overline`, `target_value` *(e.g. `"1 196 164"`)*, `sublabel` |
| `flyby_wipe` | Diagonal volumetric cloud bank sweep + rotating object cutout + optional 3D model + optical ND scrim | `overline`, `headline`, `body`, `bg_query`, `flyby_cutout_glb`, `use_hero_3d` |
| `cutout_map_stage` | Historical cartography cutout gliding on the left + frameless editorial KPI narrative | `overline`, `headline`, `body`, `map_query`, `kpis` (`[{value, label}]`) |
| `poster_depth_reveal` | Dropping accent band + giant display word (`124pt`) layered behind hero object / GLB | `giant_word`, `poster_bg_query`, `use_hero_3d`, `hero_cutout` |
| `picture_frame_pullback` | True `4.4x` concentric Morph zoom shrinking prior poster & live `!!Hero3D` into a 3D mitered gilt frame (`!!hollow_frame`) | `wall_query`, `frame_style` (`ornate_gold`/`modern_black`), `frame_x/y/w/h`, `overline`, `headline`, `subtitle`, `cta` |
| `real_portal_flythrough` | Photorealistic infinity-pool & gallery architectural fly-through with `<a:alphaModFix amt="0"/>` optical dissolve | `stage`, `portal_outer_query`, `portal_inner_query`, `portal_bg_query`, `outer_box`, `inner_box`, `kpis`, `bottom_cue` |
| `refracted_spotlight_dock` | Staggered borderless Gaussian backdrop-refracted glass dock + sleek white-underlined glider + frameless hero text | `dock_index`, `bg_query`, `spotlight`, `dock` (`[{title, subtitle, query}]`) |
| `parchment_3d_split` | Upper archival paper + lower terrain photo + centered giant display word + `256pt` side text margins around 3D hero | `giant_word`, `left_label`, `left_text`, `right_label`, `right_text`, `lower_query`, `kpis` |
| `vertical_world_elevator` | `1920x2160` 2-stop vertical pan (`upper_sky` -> `lower_ground`) with cloud-proof 1D DP sky-split typography behind peaks | `elevator_stop`, `giant_word`, `panorama_query`, `headline`, `body`, `cards` |
| `rotational_drift_3d` | Staggered left spec cards + horizontal callout connectors pointing to angled 3D hero | `overline`, `headline`, `callouts` (`[{title, desc}]`) |
| `swiss_bento_morph` | Asymmetric top telemetry strip + 3 bottom hairline/accent tiles + top-right 3D hero | `overline`, `headline`, `hero_tile`, `tiles` |
| `timeline_rail_morph` | Horizontal milestone rail with progress bar fill + active capsule node + metric drawer | `active_index`, `overline`, `headline`, `stages` |
| `chart_infographic` | 2x supersampled `comparison_bars` or `radial_rings` chart PNG + right highlight card + 3D hero | `chart_type`, `overline`, `headline`, `items`, `highlight` |
| `architecture_blueprint` | Multi-layer stacked topology diagram with flow badges + side pillar cards + 3D hero | `overline`, `headline`, `layers`, `flows`, `pillars` |
| `closing_finale` | Bold editorial summary, 3 metric KPI tiles, CTA pill, and final 3D hero angle | `overline`, `headline`, `summary`, `metrics`, `cta` |

## 4. Morph Naming Continuity Contract (`!!` Prefix & Optical Dissolve)

Both PowerPoint COM (`ppEffectMorph`) and the cross-platform builder enforce unbroken `!!`-prefixed Morph continuity verified via `verify_deck.py --require-morph-pairs`:
- `!!portal_outer` & `!!portal_inner`: Pre-staged framed (`960x540`), then post-staged zoomed (`4.6x`) with `MORPH_ALPHA_ZERO` (`<a:alphaModFix amt="0"/>`) so PowerPoint smoothly dissolves the outgoing architectural foreground as the camera flies through its aperture.
- `!!elev_sky`, `!!elev_word`, `!!elev_fg`: Pan `270pt` vertically in unison while `!!elev_word` stays locked behind the cloud-free 1D DP mountain ridge.
- `!!news_sheet`: Scaled across `newspaper_dive` stages into `editorial_print_overlay`.
- `!!odo_curtain`, `!!odo_mask_top`, `!!odo_mask_bot` & `!!odo_col_*`: Pre-staged off-canvas together so digit reels never appear unmasked during sweep-in.
- `!!flyby_obj` & `!!flyby_cloud`: Pre-staged off-canvas at `x = -880` (`< 0` so zero glow enters the prior slide).
- `!!wall_room`, `!!frame_art`, `!!hollow_frame`, `!!poster_band`, `!!poster_word`, `!!Hero3D`: Shrink concentrically on `picture_frame_pullback`.
- `!!Hero3D` & `!!HeroShadow`: Preserved ONLY on slides that visibly display the 3D subject so zero orphan models appear on unrelated slides.
