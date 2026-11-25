# Contributing to morphdeck-3d

Thank you for your interest in contributing to **morphdeck-3d** (`https://github.com/7AK3SH1/morphdeck-3d`).

## Guidelines

1. **Originality & Clean IP**: Submit only original code, procedural 3D meshes, or permissively licensed assets you authored or have explicit rights to contribute. Do not commit proprietary binaries or third-party templates.
2. **Design & Typography Discipline**:
   - Never introduce Unicode emojis in code, docs, specs, or slide templates.
   - Preserve cross-platform compatibility (Windows PowerPoint COM + macOS/Linux `python-pptx` OpenXML builder).
   - Keep all 10 procedural 3D archetypes self-contained inside `scripts/generate_subject_glb.py` so zero external `.glb` binaries are required.
3. **Commit Format**: Follow [Conventional Commits](https://www.conventionalcommits.org/) (`feat:`, `fix:`, `docs:`, `test:`, `chore:`).
4. **Pull Request Checklist**:
   - Install dependencies: `pip install -r requirements.txt`
   - Run the unit and integration suite: `python -m unittest discover -s tests -v`
   - Verify any added or modified spec examples with: `python scripts/validate_spec.py examples/<file>.json`
   - Confirm zero temporary `.pptx` or `.glb` artifacts are staged before committing.

## Licensing of Contributions

By submitting a pull request, you agree that your contributions will be licensed under the **Apache License, Version 2.0**, subject to the attribution requirements in `NOTICE`.
