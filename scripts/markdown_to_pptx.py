#!/usr/bin/env python3
"""Convert structured Markdown into a PowerPoint (.pptx) presentation based on a template.

Markdown Conventions:
  - Slide separator: '---'
  - Slide title: '# Title'
  - Subtitle / Section: '## Subtitle'
  - Bullet points: '- Item' (indented with 2 or 4 spaces for sub-levels)
  - Tables: Standard GitHub Flavored Markdown '| col1 | col2 |'
  - Speaker notes: Blockquotes '> Notes: ...' or '> Notizen: ...'
  - Two-column layout: Consecutive '## Column 1' / '## Column 2' blocks

Usage:
  python scripts/markdown_to_pptx.py --in presentation.md --out presentation.pptx [--template template.pptx]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt


@dataclass
class SlideData:
    title: str = ""
    subtitle: str = ""
    bullets: list[tuple[int, str]] = field(default_factory=list)  # (level, text)
    columns: list[list[tuple[int, str]]] = field(default_factory=list)
    table_rows: list[list[str]] = field(default_factory=list)
    speaker_notes: list[str] = field(default_factory=list)
    is_two_column: bool = False


def existing_file(path_str: str) -> Path:
    p = Path(path_str)
    if not p.exists():
        raise argparse.ArgumentTypeError(f"Datei nicht gefunden: {path_str}")
    return p


def parse_markdown(text: str) -> list[SlideData]:
    # Strip YAML Front Matter if present
    lines = text.splitlines()
    if len(lines) >= 2 and lines[0].strip() == "---":
        closing_idx = -1
        for idx in range(1, len(lines)):
            if lines[idx].strip() == "---":
                closing_idx = idx
                break
        if closing_idx != -1:
            lines = lines[closing_idx + 1 :]

    # Split slides by '---'
    raw_slides: list[list[str]] = []
    current_slide_lines: list[str] = []

    for line in lines:
        if line.strip() == "---":
            if current_slide_lines:
                raw_slides.append(current_slide_lines)
                current_slide_lines = []
        else:
            current_slide_lines.append(line)
    if current_slide_lines:
        raw_slides.append(current_slide_lines)

    parsed_slides: list[SlideData] = []

    for slide_lines in raw_slides:
        # Check if slide has non-empty lines
        non_empty = [l for l in slide_lines if l.strip()]
        if not non_empty:
            continue

        sd = SlideData()
        current_column: list[tuple[int, str]] | None = None
        in_table = False

        for raw_line in slide_lines:
            line = raw_line.rstrip()
            stripped = line.strip()

            if not stripped:
                continue

            # Speaker notes
            if stripped.startswith("> Notes:") or stripped.startswith("> Notizen:"):
                notes_text = re.sub(r"^>\s*(Notes|Notizen):\s*", "", stripped)
                sd.speaker_notes.append(notes_text)
                continue
            elif stripped.startswith(">"):
                sd.speaker_notes.append(stripped.lstrip(">").strip())
                continue

            # Title
            if stripped.startswith("# ") and not sd.title:
                sd.title = stripped[2:].strip()
                continue

            # Subtitle / Columns
            if stripped.startswith("## "):
                sub_text = stripped[3:].strip()
                if not sd.subtitle and not sd.bullets and not sd.table_rows:
                    sd.subtitle = sub_text
                else:
                    # Could be column header
                    sd.is_two_column = True
                    current_column = []
                    sd.columns.append(current_column)
                    current_column.append((0, sub_text))
                continue

            # Tables
            if stripped.startswith("|") and stripped.endswith("|"):
                in_table = True
                # Skip markdown table divider |---|---|
                if re.match(r"^\|(\s*:?-+:?\s*\|)+$", stripped):
                    continue
                cells = [c.strip() for c in stripped[1:-1].split("|")]
                sd.table_rows.append(cells)
                continue
            else:
                in_table = False

            # Bullets
            bullet_match = re.match(r"^(\s*)([-*+]|\d+\.)\s+(.*)$", line)
            if bullet_match:
                indent = len(bullet_match.group(1))
                level = min(5, indent // 2)
                b_text = bullet_match.group(3).strip()
                if sd.is_two_column and sd.columns:
                    sd.columns[-1].append((level, b_text))
                else:
                    sd.bullets.append((level, b_text))
                continue

            # Plain paragraph text as bullet level 0
            if sd.is_two_column and sd.columns:
                sd.columns[-1].append((0, stripped))
            else:
                sd.bullets.append((0, stripped))

        parsed_slides.append(sd)

    return parsed_slides


def build_presentation(slides: list[SlideData], template_path: Path | None = None) -> Presentation:
    if template_path and template_path.exists():
        prs = Presentation(str(template_path))
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    layouts = prs.slide_layouts
    layout_title = layouts[0] if len(layouts) > 0 else None
    layout_content = layouts[1] if len(layouts) > 1 else layout_title
    layout_two_content = layouts[3] if len(layouts) > 3 else layout_content

    for idx, sd in enumerate(slides):
        # Determine layout
        if idx == 0 and sd.title and sd.subtitle and not sd.bullets and not sd.table_rows and not sd.is_two_column:
            # Title slide
            slide = prs.slides.add_slide(layout_title)
            if slide.shapes.title:
                slide.shapes.title.text = sd.title
            if len(slide.placeholders) > 1:
                slide.placeholders[1].text = sd.subtitle

        elif sd.is_two_column and len(sd.columns) >= 2:
            # Two content layout
            slide = prs.slides.add_slide(layout_two_content)
            if slide.shapes.title and sd.title:
                slide.shapes.title.text = sd.title

            # Left and right content placeholders
            content_phs = [sh for sh in slide.placeholders if sh != slide.shapes.title]
            for col_idx, col_items in enumerate(sd.columns[: len(content_phs)]):
                ph = content_phs[col_idx]
                tf = ph.text_frame
                tf.word_wrap = True
                tf.clear()
                for b_idx, (level, b_text) in enumerate(col_items):
                    p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                    p.text = b_text
                    p.level = level

        else:
            # Standard Title and Content
            slide = prs.slides.add_slide(layout_content)
            if slide.shapes.title and sd.title:
                slide.shapes.title.text = sd.title

            content_ph = None
            for sh in slide.placeholders:
                if sh != slide.shapes.title:
                    content_ph = sh
                    break

            # If table
            if sd.table_rows:
                rows = len(sd.table_rows)
                cols = max(len(r) for r in sd.table_rows) if rows else 1
                # Place table nicely
                top = Inches(2.0)
                left = Inches(1.5)
                width = Inches(10.33)
                height = Inches(0.5 * rows)

                table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
                tbl = table_shape.table
                for r_i, r_data in enumerate(sd.table_rows):
                    for c_i, c_text in enumerate(r_data):
                        if c_i < cols:
                            tbl.cell(r_i, c_i).text = c_text

            elif sd.bullets and content_ph:
                tf = content_ph.text_frame
                tf.word_wrap = True
                tf.clear()
                for b_idx, (level, b_text) in enumerate(sd.bullets):
                    p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                    p.text = b_text
                    p.level = level

        # Speaker notes
        if sd.speaker_notes:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = "\n".join(sd.speaker_notes)

    return prs


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Markdown to PowerPoint (.pptx)")
    parser.add_argument("--in", dest="infile", type=existing_file, required=True, help="Eingabe Markdown-Datei")
    parser.add_argument("--out", dest="outfile", type=Path, required=True, help="Ausgabe PPTX-Datei")
    parser.add_argument("--template", dest="template", type=existing_file, help="Optionale Vorlage (.pptx)")
    args = parser.parse_args()

    md_text = args.infile.read_text(encoding="utf-8")
    slides_data = parse_markdown(md_text)

    # If no template specified, check default in templates/
    template_path = args.template
    if not template_path:
        def_tpl = Path(__file__).resolve().parents[1] / "templates" / "default-presentation.pptx"
        if def_tpl.exists():
            template_path = def_tpl

    prs = build_presentation(slides_data, template_path)

    # Atomic write
    args.outfile.parent.mkdir(parents=True, exist_ok=True)
    temp_file = tempfile.NamedTemporaryFile(
        dir=str(args.outfile.parent),
        prefix=f".tmp_{args.outfile.stem}_",
        suffix=".pptx",
        delete=False,
    )
    temp_path = Path(temp_file.name)
    temp_file.close()

    try:
        prs.save(str(temp_path))
        os.replace(str(temp_path), str(args.outfile))
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass

    out_info = {
        "action": "markdown_to_pptx",
        "success": True,
        "message": f"Generated presentation with {len(slides_data)} slides",
        "in": str(args.infile).replace("\\", "/"),
        "out": str(args.outfile).replace("\\", "/"),
        "slides_count": len(slides_data),
    }
    print(json.dumps(out_info, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
