---
name: morphdeck-3d
description: >
  True-Morph & Museum-Grade 3D PowerPoint engineering skill for AI coding agents
  across any domain (historical documentaries, luxury architecture, consumer
  hardware, SaaS, biotech, executive decks). Features 6 Visual Universes, 17
  unbroken native PowerPoint Morph engines (cantilevered infinity-pool portal
  fly-throughs with optical alpha dissolves, cloud-proof 1D horizon sky-split
  elevators, broadsheet newspaper dives, deckle-edge archival tears, masked
  odometer reels, 4.4x concentric museum gilt-frame pull-backs with live 3D
  shrink, borderless backdrop-refracted glass docks), 2x-SSAA Z-buffered smooth
  PBR CC0 .glb rendering (Khronos, Poly Haven, NASA), and juried Wikimedia
  Featured 4K photography. Zero toy 3D primitives. Zero box clutter. Zero emojis.
license: Apache-2.0
metadata:
  author: Mohamed S. Sissoko (7AK3SH1)
  version: "4.0.0"
  repository: https://github.com/7AK3SH1/morphdeck-3d
---

# morphdeck-3d v4.0.0 — True-Morph Physics & Museum-Grade 3D Optics

Build **cinema-grade, domain-adaptive PowerPoint presentations** with unbroken multi-object Morph continuity, open-canvas editorial typography, real PBR 3D geometry, and juried photography:

1. **5 Physical Laws of True-Morph Choreography (Enforced on Every Slide Pair)**:
   - **Pre-Staging & Post-Staging Continuity (`!!` Exact Pairing)**: Every incoming element on Slide `N` is pre-staged off-canvas on Slide `N-1`, and outgoing portals/elevators post-stage zoomed (`4.5x`) with `<a:alphaModFix amt="0"/>` optical dissolve. Verified automatically via `verify_deck.py --require-morph-pairs`.
   - **Staggered Kinematics**: Multi-item glass docks, callout cards, and bento tiles pre-stage at staggered coordinates (`y = 620..870` or `x = -460..-820`) so PowerPoint Morph produces cascading velocity curves in a single transition.
   - **Open-Canvas Editorial Discipline**: Never slap clunky opaque rectangles or SaaS pill badges over photography. Full-bleed slides use optical neutral-density exposure gradients (`apply_tonal_text_scrim`) so frameless display typography reads with crisp contrast.
   - **Cloud-Proof 1D Seam Sky-Split & Infinity-Pool Portals**: `split_horizon_cutout` computes a continuous 1D dynamic-programming ridge curve so `100%` of the upper sky is cloud-free (`alpha = 0`) behind giant `132pt` display letters, while `real_portal_flythrough` synthesizes cantilevered architectural pavilions with live water-rippled reflections of the 4K horizon.
   - **Grounded 3D Placement & Concentric Frame Shrink**: Standing 3D subjects rest naturally on architectural floors, plinths, or archival splits (never floating mid-air on mountain cliffs), and `picture_frame_pullback` shrinks live `<am3d:model3d>` hardware concentrically right inside a 3D mitered gilt frame on a spotlighted museum gallery wall.

2. **6 Distinct Visual Universes & 17 Native Morph Engines**:
   - *Universes*: `documentary_press`, `architectural_portal`, `sculptural_museum`, `commercial_studio`, `executive_swiss`, `editorial_pop`.
   - *Editorial & Documentary*: `newspaper_dive`, `editorial_print_overlay`, `odometer_counter` (pre-masked rolling reels + 3-layer twilight ridge), `flyby_wipe`, `cutout_map_stage`, `poster_depth_reveal`, `picture_frame_pullback`.
   - *Spatial, Product & Enterprise*: `real_portal_flythrough`, `refracted_spotlight_dock`, `parchment_3d_split`, `vertical_world_elevator`, `rotational_drift_3d`, `swiss_bento_morph`, `timeline_rail_morph`, `chart_infographic`, `architecture_blueprint`, `closing_finale`.

3. **2x-SSAA Z-Buffered PBR `.glb` + Juried Wikimedia Featured Photography**:
   - `scripts/fetch_real_glb.py` + `scripts/render_glb_preview.py`: Full glTF 2.0 scene-graph traversal, per-pixel floating-point Z-buffer, smooth interpolated vertex normals, and bilinear PBR texture sampling at 2x SSAA.
   - `scripts/fetch_real_photo.py`: Prioritizes `Featured_pictures_on_Wikimedia_Commons` and `Quality_images` while rejecting tourist crowds and photo montages.

---

## Step 0: Domain-to-Universe Routing Matrix

| Project Domain | Recommended `universe` | Real 3D Source (`hero_3d`) | Signature Slide Sequence (Vary Per Project!) |
|---|---|---|---|
| **History, Documentary, Investigative, Media** | `documentary_press` | `khronos/AntiqueCamera` or `nasa/...` | `newspaper_dive` x2 -> `editorial_print_overlay` -> `odometer_counter` -> `flyby_wipe` -> `cutout_map_stage` -> `poster_depth_reveal` -> `picture_frame_pullback` |
| **Luxury Architecture, Travel, Optics, Real Estate** | `architectural_portal` | `khronos/AntiqueCamera` or `polyhaven/marble_bust_01` | `real_portal_flythrough(exterior)` -> `real_portal_flythrough(interior)` -> `parchment_3d_split` -> `vertical_world_elevator` x2 -> `refracted_spotlight_dock` -> `closing_finale` |
| **Footwear, Hardware, Automotive, Apparel** | `commercial_studio` | `khronos/MaterialsVariantsShoe` / `CarConcept` / `BoomBox` | `poster_depth_reveal` -> `rotational_drift_3d` -> `swiss_bento_morph` -> `chart_infographic` -> `timeline_rail_morph` -> `closing_finale` |
| **Museums, Culture, Luxury Brand Storytelling** | `sculptural_museum` | `khronos/DragonAttenuation` or `polyhaven` scan | `poster_depth_reveal` -> `picture_frame_pullback` -> `parchment_3d_split` -> `refracted_spotlight_dock` -> `closing_finale` |
| **Enterprise Board, FinOps, Strategy** | `executive_swiss` | `khronos/ABeautifulGame` or `SheenChair` | `swiss_bento_morph` -> `timeline_rail_morph` -> `chart_infographic` -> `architecture_blueprint` -> `closing_finale` |
| **Consumer Creators, Music, Gaming Drops** | `editorial_pop` | `khronos/DamagedHelmet` or `BoomBox` | `editorial_print_overlay` -> `odometer_counter` -> `flyby_wipe` -> `rotational_drift_3d` -> `closing_finale` |

Full schema & Morph continuity specifications: [references/universes-and-transitions-bible.md](references/universes-and-transitions-bible.md).

---

## Turnkey Workflow

```bash
# 1. Inspect universes & slide engines
python scripts/morphdeck.py universes

# 2. Fetch or preview a real CC0 .glb or juried Wikimedia Featured photograph
python scripts/morphdeck.py fetch-3d --auto "sneaker" --out work/hero.glb
python scripts/morphdeck.py preview-3d --glb work/hero.glb --theme commercial_studio --out work/preview.png
python scripts/morphdeck.py fetch-photo --query "Lauterbrunnen valley Swiss Alps" --out work/alps.png

# 3. Validate JSON spec
python scripts/morphdeck.py validate --spec work/deck_spec.json

# 4. Build PPTX + inject 3-point <am3d:ptLight> & <a:alphaModFix> + export 1080p PNGs + verify Morph pairs
python scripts/morphdeck.py build --spec work/deck_spec.json --out Output.pptx --export-dir work/renders
```

- **Windows + PowerPoint**: Uses `scripts/build_cinema_deck.ps1` (or `build_cinema_deck.ps1.txt` in text-only skill bundles) COM automation, injects 3-point `<am3d:ptLight>` studio rigs and `<a:alphaModFix amt="0"/>` optical dissolves via `scripts/inject_3d_studio_lighting.py`, exports crisp `1920x1080` PNGs, and runs `verify_deck.py --require-lighting --require-morph-pairs`. Always inspect the rendered PNGs before finishing.
- **macOS / Linux / CI**: Uses `scripts/build_cinema_crossplatform.py` + `scripts/render_glb_preview.py` (2x-SSAA Z-buffer PBR) with full pre/post-staging Morph continuity.
