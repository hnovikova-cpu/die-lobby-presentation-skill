---
name: powerpoint
description: "Die Lobby branded PowerPoint skill using the official POTX template and brand-kit.json. Use when the user says: 'erstelle eine Präsentation', 'create a presentation', 'mach eine PowerPoint', 'make a PowerPoint', 'füge Folien hinzu', 'add slides to presentation', 'prüfe die Präsentation', 'audit presentation', 'fix brand violations', 'exportiere als PDF', 'export to PDF', 'brand check', 'Markencheck', 'neue Folien', or 'open pptx'. Orchestrates four workflows: (1) create new presentation from brief, (2) add slides to existing PPTX, (3) audit and fix brand violations, (4) export to PDF. Always uses the official POTX template and reads brand-kit.json first. Never overwrites originals — always saves as a new file."
disable-model-invocation: false
---

# powerpoint — Die Lobby Branded Presentation Skill

This skill creates, extends, audits, and exports PowerPoint presentations that follow Die Lobby's brand identity. It uses the official `.potx` template for slide layouts and `brand-kit.json` as the single source of truth for colors, fonts, and tone.

**Every operation saves a NEW file — originals are never overwritten.**

---

## Knowledge Base

| Path | Role |
|---|---|
| `/Users/halyna/Documents/CLAUDE/Skill-powerpoint/brand-kit.json` | Brand colors, fonts, tone of voice — read first every time |
| `/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/2024-10-22_Lobby_PPT-Template.potx` | Official template with 14 slide layouts and slide master |
| `/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/Samu_Barriererfreiheit.pptx` | Reference presentation (accessibility topic, 18 slides) |
| `/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/TCD_TYPO3_Dokumentation_2026-04-2_Orig.pptx` | Reference presentation (technical documentation, large) |

**Always read `brand-kit.json` before any design or audit work.**

---

## Entry Point — Always Ask First

When this skill is triggered, **always start by asking the user what they want to do:**

```
Was möchtest du tun?
[A] Neue Präsentation erstellen
[B] Folien zu bestehender Datei hinzufügen
[C] Präsentation auf Brand-Konformität prüfen und korrigieren
[D] PPTX als PDF exportieren
```

Then gather the specifics before executing. Load the relevant reference file for the chosen workflow.

---

## Four Workflows

### A — Create New Presentation
Load [workflow.md](references/workflow.md#create) for step-by-step details.

Quick summary:
1. Gather: topic, audience, number of slides (or propose an outline)
2. Copy POTX to temp `.pptx`, open with python-pptx
3. Add slides using the right layouts from [slide-layouts.md](references/slide-layouts.md)
4. Apply brand colors/fonts from [brand-rules.md](references/brand-rules.md)
5. **CHECKPOINT** — show slide plan to user before generating
6. Ask for save path → save as new `.pptx`

### B — Add Slides to Existing PPTX
Load [workflow.md](references/workflow.md#extend) for step-by-step details.

Quick summary:
1. User provides path to existing file
2. Inspect existing slides to understand context/style
3. Propose new slides with layout choices
4. **CHECKPOINT** — confirm with user
5. Append slides, save as `[original_name]_extended.pptx`

### C — Audit & Fix Brand Violations
Load [workflow.md](references/workflow.md#audit) for step-by-step details.

Quick summary:
1. User provides path to file
2. Scan all shapes: colors, fonts, backgrounds
3. Compare against brand palette and font list
4. Report violations with slide numbers
5. **CHECKPOINT** — show findings, ask to fix all or selectively
6. Apply fixes, save as `[original_name]_branded.pptx`

### D — Export to PDF
Load [workflow.md](references/workflow.md#export) for step-by-step details.

Quick summary:
1. User provides PPTX path (or uses file just created)
2. Run LibreOffice headless conversion
3. Save PDF in same folder as PPTX

---

## python-pptx Core Pattern

Load [python-pptx-guide.md](references/python-pptx-guide.md) for all code patterns.

The key constraint: **POTX files cannot be opened directly by python-pptx**. Always copy to a temp `.pptx` first:

```python
import shutil, tempfile
from pptx import Presentation

POTX = '/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/2024-10-22_Lobby_PPT-Template.potx'
tmp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
shutil.copy(POTX, tmp.name)
prs = Presentation(tmp.name)
```

---

## Brand Quick Reference

Full rules in [brand-rules.md](references/brand-rules.md). Core values:

| Token | Hex | Use |
|---|---|---|
| Neon Romance | `#e50040` | Accent, CTA elements |
| Wild Berry | `#b2003b` | Deep red accent |
| Teal With It | `#007080` | **Primary heading color** |
| Frozen Boubble | `#00dee0` | Backgrounds, accents |
| Tennis Ball | `#d4ff4d` | Highlight, signal color |
| Midnight Dreams | `#001f33` | **Body text, dark backgrounds** |
| White | `#ffffff` | Default background |
| Black | `#000000` | Logo only |

Fonts: **Neue Kabel Extrabold** (headings) · **Agenda One Medium** (body) · **Agenda One Bold** (emphasis)

Text sizing: compact slides (≤5 blocks) → sub-headers **20 pt**, body **18 pt** · dense slides → sub-headers 16–18 pt, body min **15 pt** · hard floor: 15 pt body, 13 pt captions

Slide size: **13.33 × 7.50 inches** (16:9 widescreen)

---

## Critical Rules

- **Read `brand-kit.json` first** — never invent colors or fonts
- **Never overwrite originals** — always save as a new file with a descriptive suffix
- **Use POTX layouts** — pick the right named layout from the 14 available; never place shapes arbitrarily without layout context
- **Checkpoint before saving** — show the user what will be created/changed before writing the file
- **Ask for save path every time** — user decides where the file goes
- **Body text on dark backgrounds is always White** — body text on light backgrounds is always Midnight Dreams
- **Headings use Teal With It** on light backgrounds; **Frozen Boubble** on dark backgrounds
