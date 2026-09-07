---
# ==============================================================================
# PRÄSENTATIONS-METADATEN & STEUERUNG
# ==============================================================================
# Pflichtangaben für die automatische Fußzeile (Format: 'presenter | event | date'):
presenter: "Vorname Nachname"
event: "Konferenz- oder Veranstaltungsname"
date: "01.01.2026"

# Steuerung der Foliennummerierung auf der Titelfolie:
# false = Titelfolie bleibt unnummeriert (Standard)
# true  = Titelfolie erhält ebenfalls die Foliennummer '1'
title_slide_number: false

# Optional: Manueller Fußzeilentext (überschreibt bei Angabe die Kombination oben):
# footer: "Benutzerdefinierte Fußzeile"
---

# Haupttitel der Präsentation
## Aussagekräftiger Untertitel für den Vortrag

> Notes: Begrüßung der Teilnehmer, kurze persönliche Vorstellung und Einführung in die Zielsetzung des Vortrags.

---

# Standard-Inhaltsfolie mit gegliederter Aufzählung
## Übersicht über Kernpunkte und Argumente

- **Erster Hauptpunkt:** Wichtige Kernbotschaft mit **fetter Hervorhebung**.
  - Detaillierter Unterpunkt (Ebene 1) mit *kursiver Ergänzung*.
  - Weiterer Aspekt auf zweiter Gliederungsebene.
- **Zweiter Hauptpunkt:** Konkrete Maßnahmen und nächste Schritte.
  - Weitere Vertiefung oder methodische Erläuterung.
- **Dritter Hauptpunkt:** Zusammenfassende Erkenntnis für das Publikum.

> Notes: Hier stehen die Vortragsnotizen für den Referenten. Diese werden in PowerPoint als Notizen unterhalb der Folie abgelegt und sind in der Referentenansicht sichtbar.

---

# Zwei-Spalten-Layout: Direkter Vergleich
## Gegenüberstellung zweier Perspektiven oder Phasen

## Spalte 1: Bisherige Praxis
- **Hoher Zeitaufwand:** Manuelle Erstellung jeder einzelnen Folie.
- **Inkonsistente Vorlagen:** Abweichende Farben, Ränder und Schriftgrößen.
- **Fehleranfälligkeit:** Manuelles Übertragen von Daten birgt Risiken.
- **Eingeschränkte Versionierung:** Binäre PPTX-Dateien sind schwer zu vergleichen.

## Spalte 2: Automatisierter Workflow
- **Trennung von Inhalt & Design:** Markdown für Inhalte, PowerPoint-Vorlage für CI.
- **Konsistente Master-Layouts:** Automatische Platzierung von Nummern und Fußzeilen.
- **Revisionssicherheit:** Änderungen werden transparent über Git versioniert.
- **Multimodale QA:** Automatisierte PDF- und PNG-Exporte zur visuellen Abnahme.

> Notizen: Erläutern Sie hier die Vorteile des systematischen Wandels vom manuellen Folienbau zur strukturierten Erstellung.

---

# Strukturierte Tabellenübersicht
## Projektphasen, Meilensteine und Status im Überblick

| Phase | Meilenstein & Fokus | Zeitrahmen | Status |
| --- | --- | --- | --- |
| **Phase 1: Konzeption** | Vorlagenanalyse und Festlegung der CI-Standards | Woche 1–2 | Abgeschlossen |
| **Phase 2: Umsetzung** | Markdown-Transformation & Layout-Automation | Woche 3–4 | In Bearbeitung |
| **Phase 3: Rollout** | Multi-Event-Bereitstellung & visuelle QA | Woche 5–6 | Geplant |

> Notes: Tabellen im GitHub-Flavored-Markdown-Format werden automatisch in native PowerPoint-Tabellen mit formatierter Kopfzeile konvertiert.

---

# Fazit & Handlungsaufforderung
## Die wichtigsten Erkenntnisse auf einen Blick

- **Struktur schafft Verlässlichkeit:** Saubere Workflows sparen Zeit und sichern CI-Konformität.
- **Event-Organisation:** Separate Event-Instanzen gewährleisten vollständige Reproduzierbarkeit.
- **Menschliche Führung:** Die KI liefert akribische Vorarbeit, die inhaltliche Freigabe bleibt beim Menschen.
- **Diskussion & Ausblick:** Vielen Dank für Ihre Aufmerksamkeit! Fragen und Anmerkungen?

> Notes: Zusammenfassendes Fazit ziehen, Dank an die Teilnehmer richten und in die Fragerunde überleiten.
