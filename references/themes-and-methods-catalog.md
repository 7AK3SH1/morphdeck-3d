# Themes, 3D Archetypes & Composable Slide Methods Schema (v2)

Use this reference to author `deck_spec.json`. Every archetype, theme, and slide method executes across both the Windows PowerPoint COM engine (`build_deck.ps1`) and the cross-platform OpenXML engine (`build_deck_crossplatform.py`).

---

## 1. Root Spec Envelope

```json
{
  "schema_version": 2,
  "theme": "obsidian_kinetic | editorial_ivory | cyber_blueprint | kinetic_poster | executive_clean | biotech_luminance | monochrome_luxury",
  "brand": {
    "name": "Project Name",
    "title_part1": "PROJ",
    "title_part2": "ECT",
    "tagline": "Concise single-line value proposition",
    "monogram": "PR",
    "logo_path": ""
  },
  "palette": {
    "accent1": "#2563EB",
    "accent2": "#10B981",
    "accent3": "#F59E0B",
    "chassis_color": "#2563EB"
  },
  "hero_3d": {
    "archetype": "architectural_pavilion | ai_quantum_server | energy_turbine_core | product_drop_capsule | aerospace_orbital_craft | executive_globe_monolith | biotech_helix_core | fintech_vault_prism | smartphone_pro | laptop_studio | custom",
    "secondary_archetype": "laptop_studio | smartphone_pro | ...",
    "custom_glb_path": ""
  },
  "assets": {
    "horizon_photo": "",
    "upper_photo": "",
    "mobile_ui_image": "",
    "desktop_ui_image": "",
    "custom_glb_path": ""
  },
  "slides": []
}
```

---

## 2. All 10 Procedural PBR 3D `.glb` Archetypes

Every archetype is synthesized cleanly into binary glTF 2.0 by `scripts/generate_subject_glb.py` with zero external binary dependencies:

| `hero_3d.archetype` | Target Domain | Procedural 3D Geometry & Self-Lit PBR Materials |
|---|---|---|
| `architectural_pavilion` | Architecture, Real Estate, Construction | Stepped travertine plinth, bronze portico columns, 14 Douglas fir louver fins, infinity basin, cantilevered roof + illuminated circular oculus |
| `ai_quantum_server` | AI, Cloud, Data Centers, DevOps | Anodized titanium frame, 7 GPU tensor blades, glowing RDMA vertical bus bars, self-lit silicon core, dual orbital gimbal rings |
| `energy_turbine_core` | CleanTech, Solar, Wind, Smart Grids | Industrial base podium, LFP storage tower, tilted crystalline solar arrays, 3-blade aerodynamic rotor & hub |
| `product_drop_capsule` | E-Commerce, Hardware Drops, Retail | Levitating plinth, brand-wrapped monolith cylinder & metallic crown dome, intersecting kinetic speed rings |
| `aerospace_orbital_craft` | Aerospace, Mobility, Drones, Defense | Multi-stage metallic fuselage, 4 swept delta fins, nosecone fairing, 3 plasma thruster bells |
| `executive_globe_monolith` | FinOps, Consulting, Global Enterprise | Tiered architectural plinth, illuminated global node sphere, brushed titanium & champagne gold meridian gimbal rings |
| `biotech_helix_core` | Biotech, Genomics, Pharma, MedTech | 20-step DNA double-helix backbone & glowing base-pair rungs inside a containment pedestal and dual biosensor rings |
| `fintech_vault_prism` | Fintech, Security, Cryptography | Hexagonal encrypted crystal vault core, 6 corner armor ribs, and 3 interlocking titanium security torus rings |
| `smartphone_pro` (alias `iphone`) | Mobile Apps, Field Ops, Consumer SaaS | Rounded anodized titanium frame, triple-lens camera island, Dynamic Island notch, custom-rendered OLED screen quad |
| `laptop_studio` (alias `macbook`) | Developer Tools, Web Consoles, B2B SaaS | Aluminum unibody deck, 60-key cap grid, glass trackpad, barrel hinge, angled bezel & live dashboard display quad |
| `custom` | Any external `.glb` asset (`custom_glb_path`) | Optional screen baking via `scripts/bake_glb_screen.py` |

---

## 3. All 7 Visual Theme Systems

| Theme ID | Mood & Canvas | Typography Stack | Default 3D Archetype |
|---|---|---|---|
| `obsidian_kinetic` | Deep studio `#080A0F` + warm coral/cyan | `Bahnschrift` + `Segoe UI` + `Consolas` | `smartphone_pro` |
| `editorial_ivory` | Warm paper `#F4EFE6` + burnt sienna | `Georgia` + `Segoe UI` + `Consolas` | `architectural_pavilion` |
| `cyber_blueprint` | Deep slate `#070C17` + cyan/indigo | `Bahnschrift` + `Consolas` + `Segoe UI` | `ai_quantum_server` |
| `kinetic_poster` | Contrast noir `#0B0C10` + crimson/cobalt | `Impact` + `Bahnschrift` + `Segoe UI` | `product_drop_capsule` |
| `executive_clean` | Boardroom navy `#0A1120` + royal/teal | `Bahnschrift` + `Segoe UI` + `Consolas` | `executive_globe_monolith` |
| `biotech_luminance` | Deep emerald `#061412` + bio-mint/cyan | `Bahnschrift` + `Consolas` + `Segoe UI` | `biotech_helix_core` |
| `monochrome_luxury` | Haute obsidian `#09090B` + champagne gold | `Georgia` + `Segoe UI` + `Consolas` | `fintech_vault_prism` |

---

## 4. Composable Slide Method Reference (19 Methods)

- `cover_hero`: `style` in `centered_clean | asymmetric_split | bold_poster | 3d_staged`. Pre-stages off-canvas Morph actors when Slide 2 is `split_horizon_3d`, `depth_sandwich_pan`, or `macro_dolly_3d`.
- `portal_flythrough_3d`: `stage` in `exterior | interior | horizon`. Generates dual optical bulkhead masks (`portal_outer.png`, `portal_inner.png`) with punched chamfered viewports that zoom `4.2x-4.6x` past the camera under Morph while the live 3D subject flies forward into full view.
- `vertical_elevator_3d`: `elevator_stop` in `upper_sky | lower_ground`. Traverses `-520 pt` vertically across a `1920x2160` world panorama (`elev_bg.png`) and foreground ridge cutout (`elev_fg.png`), keeping giant occluded typography behind the horizon while orbiting the live 3D hero cleanly in front of all scene layers.
- `split_horizon_3d`: Contrasting upper canvas + lower horizon strip + giant behind-model typography (`max 10 chars`) + 4 KPI cards.
- `depth_sandwich_pan`: 4-layer optical pan (`!!pan_bg` -> `!!pan_word` -> `!!pan_fg` foreground terrain ridge -> `!!Phone3D`). `pan_state`: `wide | zoom_right | zoom_left`.
- `macro_dolly_3d`: 1.55x close-up 3D camera dolly on stage right with 3 staggered-velocity frosted-glass callout cards, horizontal leader lines, and target dots.
- `orbit_steps_3d`: `variant` in `upper_stats | obsidian_steps` (`side`: `left | right`).
- `timeline_morph_step`: 2-4 consecutive slides sharing `stages[]` and incrementing `active_index` while rotating the 3D hero progressively.
- `kpi_bento`: `layout` in `left_anchor | top_hero_strip`. Pairs translucent frosted-glass metric containers with live 3D hero orbit.
- `chart_infographic`: `chart_type` in `comparison_bars | radial_rings`. Rendered at `2x` supersampling inside frosted glass alongside 3D hero + insight cards.
- `architecture_diagram`: Multi-layer system topology blueprint (`2560x1592` supersampled) + 2 right-column spec cards + isometric 3D hero.
- `problem_solution_split`: 3D hero serves as the physical center bridge between Left (Legacy) and Right (Transformed) frosted panels.
- `comparison_matrix`: 4-column feature benchmark matrix with highlighted featured column flanked by the 3D subject model.
- `chapter_divider_morph`: Oversized chapter numeral, thesis statement, active bottom TOC rail, and dramatic 3D right stage.
- `editorial_split`: Pairs rounded framed photography (`image_side`: `left | right`) with overlapping 3D hero and 2 KPI proof cards.
- `spotlight_dock`: Consecutive slides incrementing `dock_index` glide `!!DockGlider` horizontally across a true backdrop-refracted glass dock (`s*_dock_refract.png`), pre-stage staggered X entrances + orthogonal upward exits, morph pagination dot scale (`7pt` vs `14pt`), and reveal incoming backdrops via vertical `srcRect` curtain wipes.
- `macbook_studio_3d`: Showcases secondary desktop 3D model (`laptop_studio`) alongside 3 frosted engineering pillars.
- `full_bleed_visual`: Full-bleed atmospheric backdrop + bottom frosted caption bar + overlaid 3D subject model.
- `closing_finale`: `style` in `editorial_statement | executive_cta | dual_3d`.

---

## 5. Cinema Motion & 3D Studio Lighting Architecture

- **3-Point Colored Studio Lighting (`scripts/inject_3d_studio_lighting.py`)**: Every `<am3d:model3d>` container receives a theme-matched `<am3d:ambientLight>` plus three `<am3d:ptLight>` emitters (warm key, theme accent fill, colored rim backlight) matching our NumPy shader in `scripts/render_glb_preview.py`.
- **Staggered Off-Canvas Velocity**: Text and card groups pre-stage at linearly increasing off-canvas distances (`-460`, `-560`, `-680`, `-800` pt or `+1080`, `+1240`, `+1400` pt) so fixed-duration Morph produces organic acceleration curves.
- **Backdrop-Refracted Frosted Glass**: Dock surfaces sample and Gaussian-blur (`radius=24`) the exact underlying pixels of the slide background before applying frosted glass tint and a `2px` top specular rim.
