# PPTX Editor — Shared Skill

PowerPoint-Automatisierung für Agenten und Entwickler: Text-/Strukturanalyse, hierarchische LLM-Extraktion, deterministisches Patching, vorlagenbasierte Markdown-Generierung und Windows-COM-Exporte.

## Features

- **CLI Envelope & Ops (`pptx_ops.py`):** Text-Extraktion, Präsentations-Statistiken, gezieltes Suchen & Ersetzen, Folien-Inspektion mit standardisiertem JSON-Envelope (`--json`).
- **Hierarchische LLM-Extraktion (`extract_pptx_for_llm.py`):** Überführt Präsentationen in ein sauberes JSON-Schema (`pptx-structure.v1`) inkl. Shapes, Bullet-Levels, Tabellen und Speaker Notes sowie RAG-Chunking.
- **Deterministisches Patching (`apply_pptx_patch.py`):** Minimal-Invasives Schreiben mit Sicherheits-Preconditions (`expected_matches`, `expected_contains`) und Erhalt von Layout und Typografie.
- **Markdown-Konvertierung (`markdown_to_pptx.py` & `pptx_to_markdown.py`):** Folien aus strukturiertem Markdown erstellen oder bestehende Folien in Markdown exportieren.
- **Windows-COM Export (`.ps1`):** Automatisierter Export nach PDF und 1080p-PNGs je Folie für visuelle Layout-Reviews mit Vision-Modellen.
- **Autonome Testsuite (`selftest.py`):** Vollständiger End-to-End Test ohne Artefakte im Repository.

## Quick Start

```bash
# Voraussetzungen
pip install -r requirements.txt

# Statistiken ermitteln
python scripts/pptx_ops.py stats --in presentation.pptx --json

# Struktur für LLM extrahieren
python scripts/extract_pptx_for_llm.py --in presentation.pptx --out structure.json

# Patch anwenden
python scripts/apply_pptx_patch.py --in in.pptx --out out.pptx --patch patch.json

# Folien aus Markdown erzeugen
python scripts/markdown_to_pptx.py --in presentation.md --out presentation.pptx

# Selbsttest ausführen
python scripts/selftest.py
```

## Standard Template für `AGENTS.md`

```markdown
## pptx-editor

Dieser Workspace nutzt den Shared Skill bei `../skills/pptx-editor/` für die Erstellung, Analyse und Bearbeitung von PowerPoint-Präsentationen (.pptx).
Siehe [SKILL.md](../skills/pptx-editor/SKILL.md) für Befehle, Workflows und JSON-Schemas.
```

## Lizenz

MIT License (Copyright 2026 Martin Derm). Siehe [LICENSE](LICENSE).
