# 3D glTF + Morph Architecture

## 1. Dual-Engine Strategy

**morphdeck-3d** separates *asset synthesis* from *presentation assembly* so AI agents get reliable output on every OS:

1. **Asset Stage (`render_studio_assets.py` + `generate_subject_glb.py`)**:
   - Synthesizes binary glTF 2.0 (`.glb`) 3D subject models with analytical normals, forward-facing UVs, 5 PBR materials, and embedded `1024x1024` self-lit textures.
   - Pre-renders translucent frosted-glass containers (`card_*.png`, `dock_*.png`), seamless 4-layer Depth-Sandwich cutouts (`s*_pan_fg.png`), and `2x`-supersampled charts and system diagrams.
2. **Assembly Stage**:
   - **Windows PowerPoint COM (`build_deck.ps1`)**: Embeds each `.glb` via native `Shapes.Add3DModel($glbPath, 0, 1, ...)`, propagates the shared `!!Phone3D` / `!!Laptop3D` relationships via `$prevSlide.Duplicate().Item(1)`, applies native `ppEffectMorph` (`EntryEffect = 3954`), and exports real GPU-rendered `1920x1080` slide PNGs.
   - **Cross-Platform OpenXML (`build_deck_crossplatform.py` + `render_glb_preview.py`)**: On macOS, Linux, and headless CI containers, renders each `.glb` pose cleanly at its slide rotation `(rx, ry, rz)` using NumPy/Pillow PBR shading, builds valid `.pptx` slides with `python-pptx`, and injects ECMA-376 `<mc:AlternateContent><mc:Choice Requires="p159"><p:transition ...><p159:morph option="byObject"/></p:transition></mc:Choice>` XML wrappers.

---

## 2. GLB Deduplication via `.Duplicate()` (COM Engine)

Inserting `.glb` models via `Add3DModel` on every slide would duplicate a multi-megabyte binary inside `ppt/media/` N times. Instead:
- Slide 1 inserts `!!Phone3D` once (with `LinkToFile = 0`, `SaveWithDocument = 1`).
- Every subsequent slide clones `$prevSlide.Duplicate().Item(1)`, deletes non-persistent shapes backwards (`$k = Count .. 1`), and updates `.Left`, `.Top`, `.Width`, `.Height`, and `Model3D.RotationX/Y/Z`.
- All slides reference the single `ppt/media/model3d1.glb` part, keeping file size compact and giving PowerPoint identical persistent object IDs for smooth 60fps 3D Morph interpolation.

---

## 3. OpenXML Morph Transition Compatibility

Never inject naked `<p159:morph/>` directly inside `<p:transition>` without `<mc:AlternateContent>`. Doing so fails ECMA-376 strict validation in older schema parsers and triggers PowerPoint repair warnings. Both engines guarantee the compliant wrapper:

```xml
<mc:AlternateContent
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main">
  <mc:Choice Requires="p159">
    <p:transition spd="slow"
      xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
      p14:dur="1800">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback xmlns="">
    <p:transition spd="slow"><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

---

## 4. 4-Layer Depth-Sandwich Z-Order Stack

On `depth_sandwich_pan` slides, shapes are stacked cleanly back-to-front:
1. `!!pan_bg` (`ZOrder = 1`, sky + atmospheric backdrop panning under camera zoom)
2. `!!pan_word` (giant display headline partially occluded along its lower baseline)
3. `!!pan_fg` (seamless foreground ridge/subject alpha cutout masking the bottom edge of `!!pan_word`)
4. `!!PhoneShadow` + `!!Phone3D` + `pan_cap_bg` (live rotating 3D model and frosted glass caption card in foreground)
