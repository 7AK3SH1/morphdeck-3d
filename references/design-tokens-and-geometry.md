# Design Tokens & Coordinate Geometry

## 1. Widescreen 16:9 Canvas
- Dimensions: `960 x 540 pt` (`13.333 x 7.500 in`, `12,192,000 x 6,858,000 EMUs`).
- Render Export Resolution: `1920 x 1080 px` (`2x` Retina scale from point grid).
- PowerPoint COM Colors: Always pass BGR 24-bit integers (`R + G*256 + B*65536`), converted automatically by `HexToBgr`.

## 2. Persistent `!!` Morph Naming Convention
PowerPoint Morph (`byObject` / explicit `!!` semantic matching) pairs elements identically across consecutive slides by exact shape name:
- `!!Phone3D`, `!!PhoneShadow`: Primary subject 3D model + soft drop shadow.
- `!!Laptop3D`, `!!LaptopShadow`: Secondary desktop 3D model + shadow.
- `!!SplitLower`, `!!SplitUpper`, `!!GiantWord`: Horizon split panels and behind-model typography.
- `!!pan_bg`, `!!pan_word`, `!!pan_fg`: 4-layer Depth Sandwich pan actors.
- `!!tl_rail`, `!!tl_progress`, `!!tl_pointer`, `!!tl_node_0..3`, `!!tl_drawer`: Sliding timeline Morph rig.
- `!!GlassDock`, `!!DockGlider`, `!!DockThumb1..4`, `!!DockLbl1..4`, `!!DockSub1..4`: Horizontal spotlight dock switcher.

## 3. Off-Canvas Morph Staging Coordinates
To make elements sweep cleanly onto and off the slide instead of fading abruptly:
- **Sub-canvas bottom wait/exit**: `y = 565..585 pt` (`!!SplitLower` rises from `565 pt` to `275 pt` on Slide 2, then exits to `585 pt` on Slide 3).
- **Overhead wait/exit**: `y = -250..-230 pt` (`!!GiantWord` and `!!pan_word` descend from `-230 pt` into canvas and exit upward on transition).
- **Left park for dual-3D handoff**: `x = -240 pt, ry = -55 deg` on `macbook_studio_3d` so the primary hero model glides smoothly back onto center-left during `closing_finale(dual_3d)`.

## 4. Typography Vertical Clearance Rules
- **Display Serif (`Georgia`)**: Generous ascender/cap height requires `122 pt` max at `y = 38 pt` on `split_horizon_3d` so uppercase glyphs never clip against the top slide edge (`y = 0`).
- **Grotesk Display (`Bahnschrift`) / Poster (`Impact`)**: Safe up to `148 pt` at `y = 14..28 pt`.
- **Contrast Discipline**: All text over photography or dark studio backgrounds must sit over pre-composited frosted-glass cards (`alpha = 205..225` with `2px` top specular rim) or directional vignette gradients (`fade_frac >= 0.66`).
