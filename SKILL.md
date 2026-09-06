---
name: pptx-editor
description: "Arbeiten mit PowerPoint-Präsentationen (.pptx) über die Python-Bibliothek python-pptx: Folien analysieren, hierarchische LLM-Extraktion, deterministisches Patching mit Pre-Conditions, Vorlagen-basierte Markdown-Generierung, Markdown-Export und automatisierter PDF-/PNG-Export via Windows-COM."
---

# PowerPoint PPTX Editor

Nutze diesen Skill für die wiederholbare und automatisierte Erstellung, Analyse, Bearbeitung und Validierung von PowerPoint-Präsentationen (.pptx).

## Workflow

1. **Voraussetzungen prüfen:**
   - Sicherstellen, dass `python-pptx` installiert ist:
     ```bash
     pip install "python-pptx>=1.0.0"
     ```
2. **Inspektion & Statistik:**
   - Verschaffe dir mit `scripts/pptx_ops.py stats` oder `scripts/pptx_ops.py text` einen ersten Überblick über die Folien.
3. **Strukturierte LLM-Extraktion:**
   - Extrahiere die Präsentation in das Schema `pptx-structure.v1` für präzise LLM-Verarbeitung:
     ```bash
     python scripts/extract_pptx_for_llm.py --in presentation.pptx --out structure.json
     ```
4. **Deterministischer Edit-Plan (Patching):**
   - Plane gezielte Operationen (`set_slide_title`, `set_slide_bullets`, `set_speaker_notes`, `replace_text`, `fill_table`, `replace_image`).
   - Wende den Patch sicher mit Pre-Conditions an:
     ```bash
     python scripts/apply_pptx_patch.py --in in.pptx --out out.pptx --patch patch.json
     ```
5. **Markdown-Workflows:**
   - Aus strukturiertem Markdown direkt neue Folien generieren (mit Template-CI):
     ```bash
     python scripts/markdown_to_pptx.py --in slides.md --out slides.pptx [--template template.pptx]
     ```
   - Bestehende Präsentation in Markdown konvertieren:
     ```bash
     python scripts/pptx_to_markdown.py --in slides.pptx --out slides.md
     ```
6. **Visuelle QA & Export (Windows COM):**
   - Auf Windows-Hosts für PDF- und hochauflösende PNG-Exporte (1080p für multimodale Vision-Reviews):
     ```powershell
     pwsh -ExecutionPolicy Bypass -File scripts/export_pptx_to_pdf.ps1 -InputPath slides.pptx -OutputPath slides.pdf
     pwsh -ExecutionPolicy Bypass -File scripts/export_pptx_to_images.ps1 -InputPath slides.pptx -OutputDir ./preview_images
     ```
7. **Regressionstests:**
   - Führe vor Übergaben `scripts/selftest.py` aus.

---

## Standard-Befehle

### 1. Text & Statistiken (`pptx_ops.py`)
- Reinen Text aller Folien anzeigen:
  ```bash
  python scripts/pptx_ops.py text --in presentation.pptx [--no-notes] [--json]
  ```
- Statistiken (Folienanzahl, Wörter, Shapes, ungenutzte Platzhalter):
  ```bash
  python scripts/pptx_ops.py stats --in presentation.pptx [--json]
  ```
- Einfaches Suchen & Ersetzen über alle Folien hinweg:
  ```bash
  python scripts/pptx_ops.py replace --in in.pptx --out out.pptx --find "Alt" --replace "Neu" [--slide 1] [--include-notes] [--json]
  ```
- Detaillierte Folien- und Shape-Inspektion:
  ```bash
  python scripts/pptx_ops.py inspect --in presentation.pptx [--slide 1] [--json]
  ```

### 2. LLM-Extraktion (`extract_pptx_for_llm.py`)
- Hierarchische Struktur für LLMs:
  ```bash
  python scripts/extract_pptx_for_llm.py --in in.pptx --out structure.json
  ```
- Zusätzlicher RAG-Output (Blocks und Chunks für Retrieval):
  ```bash
  python scripts/extract_pptx_for_llm.py --in in.pptx --out structure.json --rag-output rag.json
  ```

### 3. Deterministisches Patching (`apply_pptx_patch.py`)
- Anwenden von Patch-JSON:
  ```bash
  python scripts/apply_pptx_patch.py --in in.pptx --out out.pptx --patch patch.json
  ```

---

## JSON Schemas & Patch-Format

### Patch-Datei-Format (`pptx-patch-schema.json`)

```json
{
  "ops": [
    {
      "op": "set_slide_title",
      "slide_number": 1,
      "title": "Neuer Folientitel",
      "expected_old_title": "Alter Titel"
    },
    {
      "op": "set_slide_bullets",
      "slide_number": 2,
      "expected_contains": "Kernpunkt",
      "bullets": [
        "Erster Hauptpunkt",
        { "text": "Detaillierter Unterpunkt", "level": 1, "bold": true }
      ]
    },
    {
      "op": "set_speaker_notes",
      "slide_number": 2,
      "text": "Redetext für den Vortragenden.",
      "mode": "replace"
    },
    {
      "op": "fill_table",
      "slide_number": 3,
      "cells": [
        { "row": 1, "col": 1, "text": "Kopfzeile 1" },
        { "row": 2, "col": 2, "text": "Neuer Wert", "mode": "replace" }
      ]
    },
    {
      "op": "replace_text",
      "slide_number": 1,
      "find": "Suchbegriff",
      "replace": "Ersatztext",
      "expected_matches": 1
    }
  ]
}
```

### Sicherheitsregeln für Patches:
- **`expected_matches`**: Schützt vor unbemerkten Mehrfachtreffern oder Nulltreffern. Weicht die tatsächliche Trefferzahl ab, bricht das Skript hart ab.
- **`expected_contains`**: Stellt sicher, dass die Folie nicht inzwischen anderweitig modifiziert wurde.
- **Atomarität**: Änderungen werden erst bei 100% fehlerfreiem Durchlauf atomar in die Zieldatei geschrieben.

---

## Markdown-Konventionen für Folien

Folien können einfach in Markdown definiert werden:

```markdown
---
# Haupttitel der Präsentation
## Untertitel für den Vortrag

---

# Agenda und Überblick
- Erster Hauptpunkt
  - Unterpunkt Ebene 1
- Zweiter Hauptpunkt

> Notes: Hier sind die Vortragsnotizen für diese Folie.

---

# Datenübersicht
| Phase | Dauer | Status |
| --- | --- | --- |
| Konzeption | 1 Woche | Abgeschlossen |
| Umsetzung | 2 Wochen | In Arbeit |
```

---

## Practical Guidance & Best Practices

- **Master-Layouts & CI:** Bei Erstellung immer die Vorlage (`--template templates/default-presentation.pptx`) nutzen, um CI-Vorgaben, Folienmaße (16:9) und Formatvorlagen beizubehalten.
- **1-basierte Foliennummerierung:** Alle Parameter (`--slide`, `slide_number`, `row`, `col`) sind für intuitive menschliche und LLM-Interaktion 1-basiert.
- **Multimodale Review:** Nach größeren Layout-Änderungen auf Windows-Hosts `export_pptx_to_images.ps1` ausführen und die Folien-PNGs zur visuellen Abnahme begutachten.
