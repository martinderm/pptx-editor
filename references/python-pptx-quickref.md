# python-pptx Quick Reference

## Setup & Dependency

- Standard-Installation:
  - `python -m pip install python-pptx`
- Minimale Version: `python-pptx>=1.0.0`

## Common Operations

### 1. Presentation öffnen und speichern
```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation("presentation.pptx")
# oder leere Präsentation:
# prs = Presentation()

# Folienabmessungen prüfen oder setzen (z. B. 16:9 Widescreen):
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

prs.save("output.pptx")
```

### 2. Folien und Standard-Layouts
Die Standardlayouts des Master-Templates (0-basiert):
- `0`: Title Slide (Title 1, Subtitle 2)
- `1`: Title and Content (Title 1, Content 2)
- `2`: Section Header (Title 1, Text 2)
- `3`: Two Content (Title 1, Content 2, Content 3)
- `4`: Comparison (Title 1, Text 2, Content 3, Text 4, Content 5)
- `5`: Title Only (Title 1)
- `6`: Blank
- `7`: Content with Caption
- `8`: Picture with Caption

```python
slide_layout = prs.slide_layouts[1]  # Title and Content
slide = prs.slides.add_slide(slide_layout)

# Titel setzen
title_shape = slide.shapes.title
if title_shape:
    title_shape.text = "Folientitel"
```

### 3. Text & Bullet-Listen
```python
# Auf Body-Platzhalter oder Textfeld zugreifen
body_shape = slide.placeholders[1]
tf = body_shape.text_frame
tf.word_wrap = True

# Erster Absatz
p0 = tf.paragraphs[0]
p0.text = "Erster Hauptpunkt"
p0.level = 0

# Weiterer Absatz (Unterpunkt)
p1 = tf.add_paragraph()
p1.text = "Unterpunkt Ebene 1"
p1.level = 1

# Formatierung auf Run-Ebene
run = p1.add_run()
run.text = " (hervorgehoben)"
run.font.bold = True
run.font.size = Pt(14)
```

### 4. Tabellen
```python
table_shape = slide.shapes.add_table(rows=3, cols=2, left=Inches(1), top=Inches(2), width=Inches(8), height=Inches(3))
tbl = table_shape.table

# Zelle ansprechen (0-basiert intern)
cell = tbl.cell(row_idx=0, col_idx=0)
cell.text = "Kopfzeile 1"
```

### 5. Vortragsnotizen (Speaker Notes)
```python
# Speaker Notes abrufen / erstellen
notes_slide = slide.notes_slide
tf_notes = notes_slide.notes_text_frame
current_notes = tf_notes.text

# Neue Notizen setzen
tf_notes.text = "Hier ist der Redetext für den Präsentierenden."
```

### 6. Bilder austauschen
```python
# In bestehendem Picture-Shape / Placeholder:
rel_id = pic_shape._element.xpath('.//a:blip/@r:embed')[0]
image_part = slide.part.related_part(rel_id)
image_part._blob = Path("neues_bild.png").read_bytes()
```

## LLM-Editing & Patch-Workflow

1. **Extrahieren:** `python scripts/extract_pptx_for_llm.py --in folien.pptx --out structure.v1.json`
2. **Patch planen:** LLM plant gezielte Operationen (`set_slide_title`, `set_slide_bullets`, `replace_text`, `set_speaker_notes`, `fill_table`, `replace_image`).
3. **Deterministisch anwenden:** `python scripts/apply_pptx_patch.py --in folien.pptx --out patched.pptx --patch patch.json`
4. **Validieren:** Pre-Conditions (`expected_matches`, `expected_contains`) sichern gegen unerwünschte Nebeneffekte ab.
