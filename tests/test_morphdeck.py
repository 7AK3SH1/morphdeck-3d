#!/usr/bin/env python3
# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
"""Comprehensive test and repository hygiene suite for morphdeck-3d."""

import glob
import json
import os
import re
import struct
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
import numpy as np
from PIL import Image, ImageStat

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(REPO_ROOT, "scripts")
EXAMPLES_DIR = os.path.join(REPO_ROOT, "examples")

if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from generate_subject_glb import ARCHETYPES, generate_subject_glb  # noqa: E402
from inject_3d_studio_lighting import inject_3d_studio_lighting  # noqa: E402
from prepare_cinema_assets import prepare_all  # noqa: E402
from render_glb_preview import render_glb_to_image  # noqa: E402
from render_studio_assets import generate_all  # noqa: E402
from studio_cutout import (  # noqa: E402
    render_cloud_sweep,
    render_hollow_gold_frame,
    render_organic_ink_drips,
    render_silhouette_horizon,
    render_soft_paper_cloud,
)
from validate_spec import validate_spec  # noqa: E402
from build_cinema_crossplatform import build_cinema_crossplatform  # noqa: E402
from build_deck_crossplatform import build_crossplatform_pptx  # noqa: E402
from verify_deck import verify_pptx  # noqa: E402

EMOJI_RE = re.compile("[\U0001F300-\U0001FAFF\u2600-\u27BF\u2300-\u23FF]", flags=re.UNICODE)


class TestMorphDeck3D(unittest.TestCase):
    def test_all_10_archetypes_generate_valid_glb(self):
        self.assertEqual(len(ARCHETYPES), 10)
        with tempfile.TemporaryDirectory() as tmp:
            for arch in ARCHETYPES:
                out_glb = os.path.join(tmp, f"{arch}.glb")
                generate_subject_glb(
                    archetype=arch,
                    out_glb=out_glb,
                    brand_name="TESTBRAND",
                    accent1_hex="#2563EB",
                    accent2_hex="#10B981",
                    accent3_hex="#F59E0B",
                )
                self.assertTrue(os.path.exists(out_glb))
                self.assertGreater(os.path.getsize(out_glb), 4096)

                with open(out_glb, "rb") as f:
                    magic, version, total_len = struct.unpack("<4sII", f.read(12))
                    self.assertEqual(magic, b"glTF")
                    self.assertEqual(version, 2)
                    self.assertEqual(total_len, os.path.getsize(out_glb))
                    jlen, jtype = struct.unpack("<I4s", f.read(8))
                    self.assertEqual(jtype, b"JSON")
                    gltf = json.loads(f.read(jlen).decode("utf-8"))

                self.assertEqual(gltf["asset"]["generator"], "morphdeck-3d-glb-engine-1.0")
                self.assertIn("Mohamed S. Sissoko", gltf["asset"]["copyright"])
                self.assertGreaterEqual(len(gltf["meshes"][0]["primitives"]), 3)

    def test_glb_preview_renderer_produces_shaded_rgba(self):
        with tempfile.TemporaryDirectory() as tmp:
            glb_path = os.path.join(tmp, "helix.glb")
            generate_subject_glb("biotech_helix_core", glb_path, brand_name="HELIXIS")
            img = render_glb_to_image(glb_path, 360, 360, rx=14, ry=-24, rz=2)
            self.assertEqual(img.mode, "RGBA")
            self.assertEqual(img.size, (360, 360))
            alpha_stat = ImageStat.Stat(img.split()[3])
            self.assertGreater(alpha_stat.mean[0], 5.0)
            self.assertLess(alpha_stat.mean[0], 240.0)

    def test_all_examples_pass_spec_validation(self):
        examples = sorted(glob.glob(os.path.join(EXAMPLES_DIR, "*.json")))
        self.assertGreaterEqual(len(examples), 7)
        for path in examples:
            errors, _ = validate_spec(path)
            self.assertEqual(errors, [], f"Validation errors in {path}: {errors}")

    def test_invalid_spec_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = os.path.join(tmp, "bad.json")
            with open(bad, "w", encoding="utf-8") as f:
                json.dump({
                    "schema_version": 2,
                    "theme": "invalid_theme",
                    "brand": {"name": "Bad"},
                    "palette": {"accent1": "NOTHEX"},
                    "hero_3d": {"archetype": "nonexistent_mesh"},
                    "slides": [
                        {"type": "unknown_slide"},
                        {"type": "portal_flythrough_3d", "stage": "bad_stage"},
                        {"type": "vertical_elevator_3d", "elevator_stop": "bad_stop"},
                    ],
                }, f)
            errors, _ = validate_spec(bad)
            self.assertGreaterEqual(len(errors), 6)

    def test_crossplatform_e2e_build_and_verify(self):
        spec = os.path.join(EXAMPLES_DIR, "biotech_therapeutics.json")
        with tempfile.TemporaryDirectory() as tmp:
            assets_dir = os.path.join(tmp, "assets")
            out_pptx = os.path.join(tmp, "biotech.pptx")
            generate_all(spec, assets_dir)
            build_crossplatform_pptx(spec, assets_dir, out_pptx)
            self.assertTrue(os.path.exists(out_pptx))
            rc = verify_pptx(out_pptx, require_3d=False)
            self.assertEqual(rc, 0)

    def test_v2_cinema_assets_and_crossplatform_e2e(self):
        spec = os.path.join(EXAMPLES_DIR, "aerospace_portal_cinema.json")
        with tempfile.TemporaryDirectory() as tmp:
            assets_dir = os.path.join(tmp, "assets")
            out_pptx = os.path.join(tmp, "aerospace.pptx")
            generate_all(spec, assets_dir)

            with Image.open(os.path.join(assets_dir, "portal_outer.png")) as im_po:
                p_out = im_po.convert("RGBA")
            with Image.open(os.path.join(assets_dir, "portal_inner.png")) as im_pi:
                p_in = im_pi.convert("RGBA")
            self.assertEqual(p_out.size, (1920, 1080))
            self.assertEqual(p_in.size, (1920, 1080))
            out_zero_frac = float(np.mean(np.array(p_out)[..., 3] == 0))
            in_zero_frac = float(np.mean(np.array(p_in)[..., 3] == 0))
            self.assertGreater(out_zero_frac, 0.15)
            self.assertLess(out_zero_frac, 0.50)
            self.assertGreater(in_zero_frac, out_zero_frac)

            with Image.open(os.path.join(assets_dir, "elev_bg.png")) as ev_bg:
                self.assertEqual(ev_bg.size, (1920, 2160))
            with Image.open(os.path.join(assets_dir, "elev_fg.png")) as im_efg:
                ev_fg = im_efg.convert("RGBA")
                self.assertEqual(ev_fg.size, (1920, 2160))
                fg_alpha = np.array(ev_fg)[..., 3]
            self.assertGreater(float(np.mean(fg_alpha == 0)), 0.10)
            self.assertGreater(float(np.mean(fg_alpha == 255)), 0.65)

            build_crossplatform_pptx(spec, assets_dir, out_pptx)
            rc = verify_pptx(out_pptx, require_3d=False)
            self.assertEqual(rc, 0)

    def test_refracted_dock_asset_generation(self):
        spec = os.path.join(EXAMPLES_DIR, "product_drop_launch.json")
        with tempfile.TemporaryDirectory() as tmp:
            assets_dir = os.path.join(tmp, "assets")
            generate_all(spec, assets_dir)
            with Image.open(os.path.join(assets_dir, "s4_dock_refract.png")) as dock_im:
                self.assertEqual(dock_im.size, (1760, 256))
                self.assertEqual(dock_im.mode, "RGBA")

    def test_inject_3d_studio_lighting_preserves_namespaces(self):
        sample_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
            'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="am3d">'
            '<am3d:model3d><am3d:spPr/><am3d:camera/><am3d:trans/>'
            '<am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr><am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight>'
            '<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="100000" g="100000" b="100000"/></am3d:clr></am3d:ptLight>'
            '<am3d:extLst/>'
            '</am3d:model3d></mc:Choice><mc:Fallback/></mc:AlternateContent>'
            '</p:spTree></p:cSld>'
            '<mc:AlternateContent><mc:Choice Requires="p159"><p:transition><p159:morph option="byObject"/></p:transition></mc:Choice></mc:AlternateContent>'
            '</p:sld>'
        )
        with tempfile.TemporaryDirectory() as tmp:
            mock_pptx = os.path.join(tmp, "mock_3d.pptx")
            with zipfile.ZipFile(mock_pptx, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("ppt/slides/slide1.xml", sample_xml.encode("utf-8"))
            upgraded = inject_3d_studio_lighting(mock_pptx, theme="cyber_blueprint")
            self.assertEqual(upgraded, 1)
            with zipfile.ZipFile(mock_pptx, "r") as z:
                out_xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
            self.assertIn("<am3d:model3d>", out_xml)
            self.assertIn("<p159:morph", out_xml)
            self.assertEqual(out_xml.count("<am3d:ambientLight>"), 1)
            self.assertEqual(out_xml.count("<am3d:ptLight"), 3)
            self.assertIn('r="40000"', out_xml)
            self.assertIn('<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="92000"', out_xml)
            ET.fromstring(out_xml)

    def test_v3_studio_cutout_generators(self):
        with tempfile.TemporaryDirectory() as tmp:
            drips_p = render_organic_ink_drips(os.path.join(tmp, "drips.png"))
            cloud_p = render_soft_paper_cloud(os.path.join(tmp, "paper_cloud.png"))
            sweep_p = render_cloud_sweep(os.path.join(tmp, "cloud_sweep.png"))
            sil_p = render_silhouette_horizon(os.path.join(tmp, "silhouette.png"))
            frame_p = render_hollow_gold_frame(os.path.join(tmp, "hollow_frame.png"), frame_style="ornate_gold", border_px=84)

            for p in (drips_p, cloud_p, sweep_p, sil_p, frame_p):
                self.assertTrue(os.path.exists(p))
                with Image.open(p) as im:
                    self.assertEqual(im.mode, "RGBA")

            # Hollow frame center aperture MUST be transparent (alpha == 0) and outer border opaque (alpha == 255)
            with Image.open(frame_p) as hf:
                arr = np.array(hf)
                self.assertEqual(int(arr[hf.height // 2, hf.width // 2, 3]), 0)
                self.assertEqual(int(arr[10, 10, 3]), 255)

    def test_v3_cinema_crossplatform_offline_e2e(self):
        spec = {
            "schema_version": 3,
            "universe": "documentary_press",
            "brand": {"name": "Offline Dispatch", "monogram": "OD"},
            "slides": [
                {"type": "newspaper_dive", "stage": "frontpage", "headline": "Archival Test", "badge": "EDITION"},
                {"type": "editorial_print_overlay", "overline": "CH 01", "headline": "Ink Overlay", "body": "Offline test."},
                {"type": "odometer_counter", "overline": "RECORDS", "target_value": "204 800", "sublabel": "VERIFIED"},
                {"type": "poster_depth_reveal", "giant_word": "1945", "use_hero_3d": False},
                {"type": "picture_frame_pullback", "headline": "Gallery Frame", "subtitle": "Concentric pull-back"},
                {"type": "closing_finale", "headline": "Complete", "summary": "Verified offline v3 deck.", "metrics": [{"value": "6", "label": "UNIVERSES"}]},
            ],
        }
        errs, _ = validate_spec(spec)
        self.assertEqual(errs, [])
        with tempfile.TemporaryDirectory() as tmp:
            spec_p = os.path.join(tmp, "v3_spec.json")
            assets_d = os.path.join(tmp, "assets")
            out_pptx = os.path.join(tmp, "v3_out.pptx")
            with open(spec_p, "w", encoding="utf-8") as f:
                json.dump(spec, f)
            prepare_all(spec_p, assets_d)
            build_cinema_crossplatform(spec_p, assets_d, out_pptx)
            self.assertEqual(verify_pptx(out_pptx, require_3d=False), 0)

    def test_v3_lighting_universes(self):
        sample_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
            'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="am3d"><am3d:model3d><am3d:spPr/><am3d:extLst/></am3d:model3d></mc:Choice></mc:AlternateContent>'
            '</p:spTree></p:cSld></p:sld>'
        )
        for univ in ("documentary_press", "architectural_portal", "commercial_studio"):
            with tempfile.TemporaryDirectory() as tmp:
                mock_pptx = os.path.join(tmp, f"{univ}.pptx")
                with zipfile.ZipFile(mock_pptx, "w", zipfile.ZIP_DEFLATED) as z:
                    z.writestr("ppt/slides/slide1.xml", sample_xml.encode("utf-8"))
                self.assertEqual(inject_3d_studio_lighting(mock_pptx, theme=univ), 1)
                with zipfile.ZipFile(mock_pptx, "r") as z:
                    out_xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
                self.assertEqual(out_xml.count("<am3d:ptLight"), 3)
                ET.fromstring(out_xml)

    def test_repository_hygiene_zero_emojis_zero_leaks(self):
        # Checked via byte decoding so this test file contains zero literal restricted strings.
        forbidden = [
            bytes([121, 101, 114, 97]).decode("ascii"),
            bytes([121, 101, 114, 114, 97]).decode("ascii"),
            bytes([112, 97, 116, 114, 101, 111, 110]).decode("ascii"),
            bytes([107, 110, 105, 103, 104, 116, 46, 112, 112, 116, 120]).decode("ascii"),
            bytes([106, 111, 117, 114, 110, 101, 121, 46, 112, 112, 116, 120]).decode("ascii"),
            bytes([97, 114, 116, 105, 115, 97, 110, 99, 111, 110, 110, 101, 99, 116]).decode("ascii"),
            bytes([98, 97, 109, 97, 107, 111]).decode("ascii"),
            bytes([99, 58, 92, 117, 115, 101, 114, 115, 92]).decode("ascii"),
        ]
        exts = (".py", ".ps1", ".sh", ".md", ".json", ".yml", ".yaml", ".toml", ".txt")
        for root, dirs, files in os.walk(REPO_ROOT):
            dirs[:] = [d for d in dirs if d not in (".git", "__pycache__", ".pytest_cache")]
            for fn in files:
                if not fn.endswith(exts):
                    continue
                fp = os.path.join(root, fn)
                with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                    txt = f.read()
                self.assertIsNone(EMOJI_RE.search(txt), f"Emoji detected in {fp}")
                low = txt.lower()
                for tok in forbidden:
                    self.assertNotIn(tok, low, f"Forbidden token found in {fp}")


if __name__ == "__main__":
    unittest.main()
