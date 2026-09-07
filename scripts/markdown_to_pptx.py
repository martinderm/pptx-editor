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
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches, Pt

try:
    from pptx_ops import clear_slides, load_presentation
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from pptx_ops import clear_slides, load_presentation


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


def parse_frontmatter(text: str) -> tuple[dict[str, str], str]:
    """Extract YAML frontmatter between opening and closing '---'."""
    meta: dict[str, str] = {}
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return meta, text

    closing_idx = -1
    for idx in range(1, len(lines)):
        if lines[idx].strip() == "---":
            closing_idx = idx
            break

    if closing_idx == -1:
        return meta, text

    fm_text = "\n".join(lines[1:closing_idx])
    remaining_text = "\n".join(lines[closing_idx + 1 :])

    try:
        import yaml

        parsed = yaml.safe_load(fm_text)
        if isinstance(parsed, dict):
            for k, v in parsed.items():
                if v is not None:
                    meta[str(k).lower().strip()] = str(v).strip()
            return meta, remaining_text
    except Exception:
        pass

    for line in fm_text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" in line:
            k, v = line.split(":", 1)
            meta[k.lower().strip()] = v.strip().strip("'\"")

    return meta, remaining_text


def validate_footer_metadata(
    meta: dict[str, str],
    cli_footer: str | None = None,
    cli_presenter: str | None = None,
    cli_event: str | None = None,
    cli_date: str | None = None,
    allow_no_footer: bool = False,
) -> tuple[str, list[str]]:
    """Validate and build footer text from metadata or CLI arguments."""
    if allow_no_footer:
        return "", []

    if cli_footer:
        return cli_footer.strip(), []
    if meta.get("footer") or meta.get("fusszeile"):
        return (meta.get("footer") or meta.get("fusszeile", "")).strip(), []

    presenter = (
        cli_presenter
        or meta.get("presenter")
        or meta.get("vortragende")
        or meta.get("vortragender")
        or meta.get("author")
        or ""
    )
    event = cli_event or meta.get("event") or meta.get("ort") or meta.get("veranstaltung") or ""
    date = cli_date or meta.get("date") or meta.get("datum") or ""

    missing = []
    if not presenter:
        missing.append("presenter (Vortragende/r)")
    if not event:
        missing.append("event (Ort/Event)")
    if not date:
        missing.append("date (Datum)")

    parts = [p.strip() for p in [presenter, event, date] if p.strip()]
    footer_text = " | ".join(parts)
    return footer_text, missing


def parse_markdown(text: str) -> tuple[dict[str, str], list[SlideData]]:
    # Extract YAML Front Matter if present
    meta, remaining_text = parse_frontmatter(text)

    # Split slides by '---'
    lines = remaining_text.splitlines()

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

    return meta, parsed_slides


def add_formatted_runs(paragraph: Any, text: str, default_bold: bool = False) -> None:
    """Parse inline Markdown (**bold**, *italic*) and add formatted runs."""
    pattern = re.compile(r"(\*\*.*?\*\*|\*.*?\*|__.*?__|_.*?_)")
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        if (part.startswith("**") and part.endswith("**")) or (part.startswith("__") and part.endswith("__")):
            run = paragraph.add_run()
            run.text = part[2:-2]
            run.font.bold = True
        elif (part.startswith("*") and part.endswith("*")) or (part.startswith("_") and part.endswith("_")):
            run = paragraph.add_run()
            run.text = part[1:-1]
            run.font.italic = True
        else:
            run = paragraph.add_run()
            run.text = part
            if default_bold:
                run.font.bold = True


def apply_slide_number_and_footer(
    slide: Any,
    slide_num: int,
    total_slides: int,
    footer_text: str,
    prs: Presentation,
    skip_slide_number: bool = False,
) -> None:
    """Clone and populate slide number and footer placeholders from layout, with fallbacks."""
    layout = slide.slide_layout
    existing_types = set()
    for ph in slide.placeholders:
        try:
            existing_types.add(ph.placeholder_format.type)
        except Exception:
            pass

    # 1. Clone layout placeholders if missing on the slide
    if layout and hasattr(layout, "placeholders"):
        for l_ph in layout.placeholders:
            try:
                p_type = l_ph.placeholder_format.type
                if p_type == PP_PLACEHOLDER.SLIDE_NUMBER and not skip_slide_number and p_type not in existing_types:
                    slide.shapes.clone_placeholder(l_ph)
                    existing_types.add(p_type)
                elif p_type == PP_PLACEHOLDER.FOOTER and footer_text and p_type not in existing_types:
                    slide.shapes.clone_placeholder(l_ph)
                    existing_types.add(p_type)
            except Exception:
                pass

    # 2. Populate placeholders
    has_num_ph = False
    has_footer_ph = False

    for s_ph in slide.placeholders:
        try:
            p_type = s_ph.placeholder_format.type
            if p_type == PP_PLACEHOLDER.SLIDE_NUMBER:
                if skip_slide_number:
                    s_ph.text = ""
                else:
                    s_ph.text = str(slide_num)
                has_num_ph = True
            elif p_type == PP_PLACEHOLDER.FOOTER and footer_text:
                s_ph.text = footer_text
                has_footer_ph = True
        except Exception:
            pass

    # 3. Fallbacks if template layout does not have these placeholders
    if not has_num_ph and not skip_slide_number:
        tb = slide.shapes.add_textbox(
            prs.slide_width - Inches(1.5),
            prs.slide_height - Inches(0.55),
            Inches(1.2),
            Inches(0.4),
        )
        p = tb.text_frame.paragraphs[0]
        p.text = str(slide_num)
        p.font.size = Pt(10)

    if not has_footer_ph and footer_text:
        tb = slide.shapes.add_textbox(
            Inches(1.0),
            prs.slide_height - Inches(0.55),
            prs.slide_width - Inches(3.0),
            Inches(0.4),
        )
        p = tb.text_frame.paragraphs[0]
        p.text = footer_text
        p.font.size = Pt(10)


def parse_bool(val: Any, default: bool = False) -> bool:
    """Parse boolean value from string or bool, with default."""
    if val is None:
        return default
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    if s in ("true", "1", "yes", "ja"):
        return True
    if s in ("false", "0", "no", "nein"):
        return False
    return default


def build_presentation(
    slides: list[SlideData],
    template_path: Path | None = None,
    clear_template_slides: bool = True,
    footer_text: str = "",
    title_slide_number: bool = False,
) -> Presentation:
    if template_path and template_path.exists():
        prs = load_presentation(template_path)
        if clear_template_slides:
            clear_slides(prs)
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

            # Left and right content placeholders only
            content_phs = [
                sh
                for sh in slide.placeholders
                if sh != slide.shapes.title
                and getattr(sh, "placeholder_format", None)
                and sh.placeholder_format.type in (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY)
            ]
            if not content_phs:
                content_phs = [sh for sh in slide.placeholders if sh != slide.shapes.title]

            for col_idx, col_items in enumerate(sd.columns[: len(content_phs)]):
                ph = content_phs[col_idx]
                tf = ph.text_frame
                tf.word_wrap = True
                tf.clear()
                for b_idx, (level, b_text) in enumerate(col_items):
                    p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                    p.text = ""
                    p.level = level
                    is_col_header = (b_idx == 0)
                    add_formatted_runs(p, b_text, default_bold=is_col_header)

        else:
            # Standard Title and Content
            slide = prs.slides.add_slide(layout_content)
            if slide.shapes.title and sd.title:
                slide.shapes.title.text = sd.title

            content_ph = None
            for sh in slide.placeholders:
                if sh != slide.shapes.title and getattr(sh, "placeholder_format", None):
                    if sh.placeholder_format.type in (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY):
                        content_ph = sh
                        break
            if not content_ph:
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
                            cell = tbl.cell(r_i, c_i)
                            cell.text = ""
                            cp = cell.text_frame.paragraphs[0]
                            add_formatted_runs(cp, c_text, default_bold=(r_i == 0))

            elif sd.bullets and content_ph:
                tf = content_ph.text_frame
                tf.word_wrap = True
                tf.clear()
                for b_idx, (level, b_text) in enumerate(sd.bullets):
                    p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                    p.text = ""
                    p.level = level
                    add_formatted_runs(p, b_text)

        # Apply Slide Number & Footer
        skip_num = (idx == 0) and not title_slide_number
        apply_slide_number_and_footer(
            slide=slide,
            slide_num=idx + 1,
            total_slides=len(slides),
            footer_text=footer_text,
            prs=prs,
            skip_slide_number=skip_num,
        )

        # Speaker notes
        if sd.speaker_notes:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = "\n".join(sd.speaker_notes)

    return prs


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Markdown to PowerPoint (.pptx)")
    parser.add_argument("--in", dest="infile", type=existing_file, required=True, help="Eingabe Markdown-Datei")
    parser.add_argument("--out", dest="outfile", type=Path, required=True, help="Ausgabe PPTX-Datei")
    parser.add_argument("--template", dest="template", type=existing_file, help="Optionale Vorlage (.pptx oder .potx)")
    parser.add_argument(
        "--clear-template-slides",
        dest="clear_template_slides",
        action="store_true",
        default=True,
        help="Entfernt vorhandene Dummy-Folien aus der Vorlage (Standard: True)",
    )
    parser.add_argument(
        "--keep-template-slides",
        dest="clear_template_slides",
        action="store_false",
        help="Bestehende Folien in der Vorlage beibehalten",
    )
    parser.add_argument("--presenter", dest="presenter", type=str, help="Vortragende/r (falls nicht im Frontmatter)")
    parser.add_argument("--event", dest="event", type=str, help="Ort/Event (falls nicht im Frontmatter)")
    parser.add_argument("--date", dest="date", type=str, help="Datum (falls nicht im Frontmatter)")
    parser.add_argument("--footer", dest="footer", type=str, help="Manueller Fußzeilentext")
    parser.add_argument(
        "--allow-no-footer",
        dest="allow_no_footer",
        action="store_true",
        help="Erlaubt Folien ohne Fußzeile",
    )
    parser.add_argument(
        "--strict-footer",
        dest="strict_footer",
        action="store_true",
        help="Bricht mit Fehler ab, wenn Fußzeilen-Metadaten unvollständig sind",
    )
    parser.add_argument(
        "--title-slide-number",
        dest="title_slide_number",
        action="store_true",
        default=None,
        help="Foliennummer auch auf der Titelfolie anzeigen",
    )
    parser.add_argument(
        "--no-title-slide-number",
        dest="title_slide_number",
        action="store_false",
        help="Keine Foliennummer auf der Titelfolie anzeigen (Standard)",
    )
    args = parser.parse_args()

    md_text = args.infile.read_text(encoding="utf-8")
    meta, slides_data = parse_markdown(md_text)

    # Determine whether title slide should display slide number
    meta_title_num = parse_bool(
        meta.get("title_slide_number")
        or meta.get("first_slide_number")
        or meta.get("number_title_slide")
        or meta.get("slide_number_on_title"),
        default=False,
    )
    show_title_num = args.title_slide_number if args.title_slide_number is not None else meta_title_num

    # Validate footer metadata
    footer_text, missing = validate_footer_metadata(
        meta=meta,
        cli_footer=args.footer,
        cli_presenter=args.presenter,
        cli_event=args.event,
        cli_date=args.date,
        allow_no_footer=args.allow_no_footer,
    )

    if missing and not args.allow_no_footer:
        msg = (
            f"Fußzeilen-Metadaten unvollständig ({', '.join(missing)}). "
            "Bitte im Frontmatter der Markdown-Datei angeben (presenter, event, date) oder per CLI übergeben."
        )
        if args.strict_footer:
            raise ValueError(msg)
        else:
            print(f"WARNUNG: {msg}", file=sys.stderr)

    # If no template specified, check default in templates/
    template_path = args.template
    if not template_path:
        def_tpl = Path(__file__).resolve().parents[1] / "templates" / "default-presentation.pptx"
        if def_tpl.exists():
            template_path = def_tpl

    prs = build_presentation(
        slides=slides_data,
        template_path=template_path,
        clear_template_slides=args.clear_template_slides,
        footer_text=footer_text,
        title_slide_number=show_title_num,
    )

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
        "footer": footer_text,
        "metadata": meta,
    }
    print(json.dumps(out_info, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
