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
     python scripts/markdown_to_pptx.py --in slides.md --out slides.pptx [--template template.pptx|template.potx] [--clear-template-slides]
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
- Detaillierte Folien- und Shape-Inspektion (unterstützt `.pptx` und `.potx`):
  ```bash
  python scripts/pptx_ops.py inspect --in presentation.pptx [--slide 1] [--json]
  ```
- PowerPoint-Vorlage (`.potx`) in bearbeitbare `.pptx` konvertieren:
  ```bash
  python scripts/pptx_ops.py convert-template --in template.potx --out template.pptx [--json]
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

Folien werden als strukturiertes Markdown definiert. Als schlüsselfertige Vorlage steht das generische Template [`templates/presentation-template.md`](templates/presentation-template.md) bereit, das alle unterstützten Features und das vollständige Frontmatter enthält.

### 1. Pflicht-Frontmatter für Metadaten & Fußzeile
Jede Präsentationsdatei muss am Dateianfang ein YAML-Frontmatter mit Metadaten für die Fußzeile und Steuerungseinstellungen enthalten:

```yaml
---
presenter: "Vorname Nachname"
event: "Konferenz / Event / Anlass"
date: "01.01.2026"
title_slide_number: false   # Optional: true | false (steuert, ob die Titelfolie nummeriert wird; Standard: false)
---
```

> [!IMPORTANT]
> **Agenten-Regel — Fehlende Metadaten erfragen:**
> Wenn in einer zu verarbeitenden Markdown-Präsentation kein Frontmatter mit `presenter`, `event` und `date` angegeben ist, **muss der KI-Agent den Nutzer vor der Generierung gezielt danach fragen**. Metadaten dürfen niemals eigenmächtig erfunden werden.

### 2. Folienaufbau in Markdown
Folien werden durch `---` getrennt:

```markdown
---
presenter: "Vorname Nachname"
event: "Konferenz / Event"
date: "01.01.2026"
title_slide_number: false
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

### 3. Foliennummerierung & Fußzeilen-Verhalten
- **Foliennummerierung:**
  - Jede Inhaltsfolie erhält eine fortlaufende Foliennummer (2, 3...).
  - **Steuerung der Titelfolie via Frontmatter:** Ob auf der ersten Folie (Titelfolie) eine Nummerierung erscheint, wird über `title_slide_number: true | false` im Markdown-Frontmatter bestimmt (Standard: `false`).
  - **Layout-Integration:** Wenn die verwendete Vorlage (`.pptx` oder `.potx`) entsprechende Layout-Platzhalter (`PP_PLACEHOLDER.SLIDE_NUMBER`) definiert, werden diese automatisch geklont und befüllt. Dadurch werden Position, Schriftart, Farbe und Layout-Elemente der jeweiligen Vorlage exakt beibehalten. Fehlt der Platzhalter im Template, wird eine konsistente Nummerierung im unteren Folienbereich eingefügt.
- **Fußzeile:**
  - Die Fußzeile wird automatisch im Format `Vortragende/r | Ort/Event | Datum` aus den Frontmatter-Feldern erzeugt.
  - Wenn die Vorlage einen Layout-Platzhalter (`PP_PLACEHOLDER.FOOTER`) besitzt, wird dieser geklont und befüllt (unter Beibehaltung von Typografie und Ausrichtung der Vorlage). Andernfalls wird ein sauberer Fallback generiert.

---

## Ordnerstruktur für Präsentationen (Multi-Version & Events)

Präsentationen können für verschiedene Anlässe, Events oder Zielgruppen angepasst werden. Hierfür gilt folgende Standard-Ordnerstruktur (z. B. unter `memory/operations/marketing/presentations/<topic-slug>/`):

```
presentations/<topic-slug>/
├── README.md                           # Übersicht zum Foliensatz, Anlässen & Zielgruppen
├── master/                             # Kanonischer Master-Foliensatz
│   ├── deck.md                         # Vollständiger Master-Inhalt
│   ├── deck.pptx                       # Master PowerPoint
│   └── deck.pdf                        # Master PDF
└── events/                             # Event-, Datums- oder zielgruppenspezifische Versionen
    ├── <YYYY-MM-DD>_<event-slug-1>/    # Konkrete Event-Instanz 1
    │   ├── deck.md                     # Angepasstes Markdown (mit spezifischem Frontmatter)
    │   ├── deck.pptx                   # Generierte Event-Präsentation
    │   ├── deck.pdf                    # PDF für Vortrag / Handout
    │   └── previews/                   # Rendered Folien-PNGs (Visuelle QA)
    │       ├── slide-01.png
    │       └── ...
    └── <YYYY-MM-DD>_<event-slug-2>/    # Weitere Event-Instanz 2
        ├── deck.md
        ├── deck.pptx
        ├── deck.pdf
        └── previews/
```

**Vorteile:**
1. **Klare Trennung:** Der Master bleibt die vollständige inhaltliche Basis; Events enthalten die jeweilige gekürzte/spezifisch gefilterte Instanz.
2. **Eigenständigkeit:** Jedes Event-Verzeichnis ist in sich geschlossen mit eigenem Frontmatter, PPTX, PDF und Previews.
3. **Reproduzierbarkeit:** Bei Nachfragen ist genau nachvollziehbar, welche Folien bei welchem Event gezeigt wurden.

---

## Practical Guidance & Best Practices

- **Master-Layouts & CI:** Bei Erstellung immer die Vorlage (`--template templates/default-presentation.pptx` oder `.potx`) nutzen, um CI-Vorgaben, Folienmaße (16:9) und Formatvorlagen beizubehalten.
- **1-basierte Foliennummerierung:** Alle Parameter (`--slide`, `slide_number`, `row`, `col`) sind für intuitive menschliche und LLM-Interaktion 1-basiert.
- **Multimodale Review:** Nach größeren Layout-Änderungen auf Windows-Hosts `export_pptx_to_images.ps1` ausführen und die Folien-PNGs zur visuellen Abnahme begutachten.
