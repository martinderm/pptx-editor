#!/usr/bin/env python3
"""Project self-test suite for pptx-editor.

Creates a synthetic fixture presentation in a temporary directory,
tests extraction, ops, patching, precondition failures, markdown conversion,
and validates key contracts without leaving any test artifacts.

Usage:
  python scripts/selftest.py
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
OPS = SCRIPTS / "pptx_ops.py"
EXTRACT = SCRIPTS / "extract_pptx_for_llm.py"
PATCH = SCRIPTS / "apply_pptx_patch.py"
MD2PPTX = SCRIPTS / "markdown_to_pptx.py"
PPTX2MD = SCRIPTS / "pptx_to_markdown.py"
DEFAULT_TPL = ROOT / "templates" / "default-presentation.pptx"


def run_py(cmd_args: list[str], cwd: Path) -> subprocess.CompletedProcess:
    full_cmd = [sys.executable, *cmd_args]
    return subprocess.run(full_cmd, cwd=str(cwd), text=True, capture_output=True, check=True)


def build_fixture_presentation(path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title slide
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Fixture Presentation"
    if len(s1.placeholders) > 1:
        s1.placeholders[1].text = "Automated Self-Test Suite"
    s1.notes_slide.notes_text_frame.text = "Notes for Slide 1"

    # Slide 2: Bullets
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Features and Components"
    tf = s2.placeholders[1].text_frame
    p0 = tf.paragraphs[0]
    p0.text = "Core Engine"
    p0.level = 0
    p1 = tf.add_paragraph()
    p1.text = "Extraction sub-module"
    p1.level = 1
    s2.notes_slide.notes_text_frame.text = "Explain core features."

    # Slide 3: Table
    s3 = prs.slides.add_slide(prs.slide_layouts[1])
    s3.shapes.title.text = "Data Matrix"
    tbl_shape = s3.shapes.add_table(2, 2, Inches(1.5), Inches(2.0), Inches(10.0), Inches(2.0))
    tbl = tbl_shape.table
    tbl.cell(0, 0).text = "H1"
    tbl.cell(0, 1).text = "H2"
    tbl.cell(1, 0).text = "Val1"
    tbl.cell(1, 1).text = "Val2"

    prs.save(str(path))


def main() -> int:
    print("=== Starting pptx-editor self-test suite ===")
    with tempfile.TemporaryDirectory(prefix="pptx-editor-selftest-") as tmpdir:
        t = Path(tmpdir)
        fixture_pptx = t / "fixture.pptx"
        extract_json = t / "extract.json"
        rag_json = t / "rag.json"
        patch_json = t / "patch.json"
        patched_pptx = t / "patched.pptx"
        exported_md = t / "exported.md"
        regenerated_pptx = t / "regenerated.pptx"

        # 1. Build fixture
        print("[1/6] Building synthetic fixture presentation...")
        build_fixture_presentation(fixture_pptx)
        assert fixture_pptx.exists(), "Fixture PPTX was not created."

        # 2. Test pptx_ops (stats & text)
        print("[2/6] Testing pptx_ops.py stats and text...")
        p_stats = run_py([str(OPS), "stats", "--in", str(fixture_pptx), "--json"], t)
        stats_env = json.loads(p_stats.stdout)
        assert stats_env["success"] is True
        assert stats_env["data"]["slide_count"] == 3
        assert stats_env["data"]["shapes"]["tables"] >= 1

        p_text = run_py([str(OPS), "text", "--in", str(fixture_pptx), "--json"], t)
        text_env = json.loads(p_text.stdout)
        assert text_env["success"] is True
        assert len(text_env["data"]["slides"]) == 3

        # 3. Test extract_pptx_for_llm
        print("[3/6] Testing extract_pptx_for_llm.py (v1 and RAG)...")
        run_py([
            str(EXTRACT),
            "--in", str(fixture_pptx),
            "--out", str(extract_json),
            "--rag-output", str(rag_json),
        ], t)
        assert extract_json.exists()
        assert rag_json.exists()
        v1_data = json.loads(extract_json.read_text(encoding="utf-8"))
        assert v1_data["schema"] == "pptx-structure.v1"
        assert v1_data["stats"]["total_slides"] == 3
        assert v1_data["slides"][0]["title"] == "Fixture Presentation"
        assert v1_data["slides"][0]["speaker_notes"]["has_notes"] is True

        rag_data = json.loads(rag_json.read_text(encoding="utf-8"))
        assert rag_data["schema"] == "pptx-rag-chunks.v1"
        assert len(rag_data["blocks"]) >= 3

        # 4. Test apply_pptx_patch
        print("[4/6] Testing apply_pptx_patch.py with preconditions...")
        patch_spec = {
            "ops": [
                {
                    "op": "set_slide_title",
                    "slide_number": 1,
                    "title": "Updated Fixture Title",
                    "expected_old_title": "Fixture",
                },
                {
                    "op": "set_slide_bullets",
                    "slide_number": 2,
                    "bullets": [
                        "Updated Core Engine",
                        {"text": "Refactored Extraction", "level": 1, "bold": True},
                    ],
                    "expected_contains": "Core Engine",
                },
                {
                    "op": "set_speaker_notes",
                    "slide_number": 1,
                    "text": "New presentation notes for slide 1.",
                    "mode": "replace",
                    "expected_contains": "Notes for Slide 1",
                },
                {
                    "op": "fill_table",
                    "slide_number": 3,
                    "cells": [
                        {"row": 2, "col": 1, "text": "UpdatedVal1"},
                        {"row": 2, "col": 2, "text": "UpdatedVal2"},
                    ],
                },
                {
                    "op": "replace_text",
                    "slide_number": 2,
                    "find": "Updated Core",
                    "replace": "Patched Core",
                    "expected_matches": 1,
                },
            ]
        }
        patch_json.write_text(json.dumps(patch_spec), encoding="utf-8")
        run_py([
            str(PATCH),
            "--in", str(fixture_pptx),
            "--out", str(patched_pptx),
            "--patch", str(patch_json),
        ], t)
        assert patched_pptx.exists()

        # Validate patch results via Presentation object
        prs_patched = Presentation(str(patched_pptx))
        assert prs_patched.slides[0].shapes.title.text == "Updated Fixture Title"
        assert "New presentation notes" in prs_patched.slides[0].notes_slide.notes_text_frame.text
        s2_bullets = [p.text for p in prs_patched.slides[1].placeholders[1].text_frame.paragraphs]
        assert "Patched Core Engine" in s2_bullets[0]
        assert "Refactored Extraction" in s2_bullets[1]

        # 4b. Verify Precondition Failure behavior
        bad_patch = t / "bad_patch.json"
        bad_patch.write_text(json.dumps({
            "ops": [
                {
                    "op": "set_slide_title",
                    "slide_number": 1,
                    "title": "Fail Title",
                    "expected_old_title": "NON_EXISTENT_STRING",
                }
            ]
        }), encoding="utf-8")
        failed_res = subprocess.run(
            [sys.executable, str(PATCH), "--in", str(patched_pptx), "--out", str(t / "should_not_exist.pptx"), "--patch", str(bad_patch)],
            cwd=str(t), capture_output=True, text=True
        )
        assert failed_res.returncode != 0, "apply_pptx_patch must fail on precondition mismatch."
        assert not (t / "should_not_exist.pptx").exists(), "Must not write file on precondition failure."

        # 5. Test pptx_to_markdown & markdown_to_pptx
        print("[5/6] Testing pptx_to_markdown.py and markdown_to_pptx.py...")
        run_py([str(PPTX2MD), "--in", str(patched_pptx), "--out", str(exported_md)], t)
        assert exported_md.exists()
        md_text = exported_md.read_text(encoding="utf-8")
        assert "# Updated Fixture Title" in md_text
        assert "Patched Core Engine" in md_text
        assert "> Notes: New presentation notes" in md_text

        run_py([
            str(MD2PPTX),
            "--in", str(exported_md),
            "--out", str(regenerated_pptx),
            "--template", str(DEFAULT_TPL),
        ], t)
        assert regenerated_pptx.exists()
        prs_regen = Presentation(str(regenerated_pptx))
        assert len(prs_regen.slides) == 3

        # 6. Verify default template integrity
        print("[6/7] Verifying default template layout integrity...")
        assert DEFAULT_TPL.exists()
        prs_tpl = Presentation(str(DEFAULT_TPL))
        assert len(prs_tpl.slide_layouts) >= 5

        # 7. Test potx template support & convert-template
        print("[7/7] Testing .potx template support and convert-template...")
        import zipfile
        fixture_potx = t / "fixture.potx"
        with zipfile.ZipFile(fixture_pptx, "r") as z_in, zipfile.ZipFile(fixture_potx, "w") as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                        b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                    )
                z_out.writestr(item, data)
        assert fixture_potx.exists()

        # Test inspecting .potx
        p_inspect_potx = run_py([str(OPS), "inspect", "--in", str(fixture_potx), "--json"], t)
        inspect_env = json.loads(p_inspect_potx.stdout)
        assert inspect_env["success"] is True
        assert inspect_env["data"]["total_slides"] == 3

        # Test convert-template
        converted_pptx = t / "converted_from_potx.pptx"
        run_py([str(OPS), "convert-template", "--in", str(fixture_potx), "--out", str(converted_pptx), "--json"], t)
        assert converted_pptx.exists()

        # Test markdown_to_pptx using .potx template with slide clearing
        potx_out_pptx = t / "built_with_potx.pptx"
        run_py([
            str(MD2PPTX),
            "--in", str(exported_md),
            "--out", str(potx_out_pptx),
            "--template", str(fixture_potx),
            "--clear-template-slides",
        ], t)
        assert potx_out_pptx.exists()
        prs_potx_out = Presentation(str(potx_out_pptx))
        assert len(prs_potx_out.slides) == 3

    print("=== All pptx-editor self-tests PASSED successfully! ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
