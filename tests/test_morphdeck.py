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
from PIL import Image, ImageStat

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(REPO_ROOT, "scripts")
EXAMPLES_DIR = os.path.join(REPO_ROOT, "examples")

if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from generate_subject_glb import ARCHETYPES, generate_subject_glb  # noqa: E402
from render_glb_preview import render_glb_to_image  # noqa: E402
from render_studio_assets import generate_all  # noqa: E402
from validate_spec import validate_spec  # noqa: E402
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
        self.assertGreaterEqual(len(examples), 6)
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
                    "slides": [{"type": "unknown_slide"}],
                }, f)
            errors, _ = validate_spec(bad)
            self.assertGreaterEqual(len(errors), 4)

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
