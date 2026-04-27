# Workflow Details — Die Lobby PowerPoint Skill

---

## A — Create New Presentation {#create}

### Step 1: Brief
Ask the user:
- **Topic / Titel** — what is the presentation about?
- **Zielgruppe** — who is the audience? (client, internal team, conference)
- **Anzahl Folien** — approximate slide count, or should Claude propose an outline?
- **Ton** — formal/professional or approachable/dynamic? (default: professional, human — per brand tone)
- **Sprache** — German or English? (default: German — brand language is de-DE)

### Step 2: Propose Outline
Based on the brief, propose a slide-by-slide structure:

```
Folie 1:  Titelfolie          — "Titel · Untertitel · Datum"
Folie 2:  Titel und Inhalt    — Agenda / Überblick
Folie 3:  Abschnittsüberschrift — Kapitel 1
Folie 4:  Titel und Inhalt    — Inhalt Kapitel 1
Folie 5:  Bild mit Titel      — Beispiel / Visual
...
Letzte:   Titelfolie          — Vielen Dank · Kontakt
```

Get user approval on the outline before generating.

### Step 3: Generate
Run a Python script using the patterns from [python-pptx-guide.md](python-pptx-guide.md):
1. Copy POTX to temp file
2. Open with `Presentation()`
3. Add slides in order using `layout_map[name]`
4. Fill title and body placeholders
5. Apply brand fonts and colors from [brand-rules.md](brand-rules.md)

### Step 4: Checkpoint
Before saving, report what was created:
```
Ergebnis:
- 8 Folien erstellt
- Layouts verwendet: Titelfolie (×2), Titel und Inhalt (×4), Abschnittsüberschrift (×1), Bild mit Titel (×1)
- Farben: Teal With It (Überschriften), Midnight Dreams (Text), Neon Romance (Akzente)
- Schriften: Neue Kabel Extrabold (Titel), Agenda One Medium (Text)

Wo soll die Datei gespeichert werden?
```

### Step 5: Save
Ask user for save path. Save with `prs.save(path)`. Confirm:
```
Gespeichert: /pfad/zur/datei.pptx
```

---

## B — Add Slides to Existing PPTX {#extend}

### Step 1: Open the existing file
```python
prs = Presentation('/path/to/existing.pptx')
```

Note the existing slide count, layouts used, and general style (background color, heading style).

### Step 2: Inspect existing slides
Scan the first 3 slides to understand the visual style being used:
- Background colors
- Heading font and color
- Whether it uses brand layouts or custom shapes

Report back to user: "Ich sehe X Folien. Die Präsentation verwendet [Layouts / Stilmerkmale]. Ich werde neue Folien im gleichen Stil ergänzen."

### Step 3: Gather new slide brief
- **Inhalt** — what should the new slides cover?
- **Anzahl** — how many slides to add?
- **Position** — at the end, or after a specific slide number?

### Step 4: Propose new slides
Show the outline before generating. Get approval.

### Step 5: Add slides
Append slides using `prs.slides.add_slide()`. Apply brand style consistently with existing slides.

### Step 6: Save as new file
Never overwrite the original. Save as:
```python
import os
base = os.path.splitext(original_path)[0]
output = base + '_extended.pptx'
prs.save(output)
```

Confirm: `Neue Datei gespeichert: [path]_extended.pptx (Original unverändert)`

---

## C — Audit & Fix Brand Violations {#audit}

### Step 1: Open the file
```python
prs = Presentation('/path/to/file.pptx')
```

### Step 2: Run audit
Use `audit_presentation(prs)` from [python-pptx-guide.md](python-pptx-guide.md).

Collect all violations grouped by type:
- **Farben** (off-brand colors)
- **Schriften** (wrong fonts)
- **Hintergründe** (non-brand backgrounds)

### Step 3: Report findings
Present a clear summary:
```
Audit-Ergebnis: 23 Verstöße in 8 Folien

Farben (12 Verstöße):
  Folie 3: Füllfarbe #336699 auf Form "Rectangle 1" (nicht in Markenpalette)
  Folie 5: Textfarbe #444444 auf "Textbox 2"
  ...

Schriften (11 Verstöße):
  Folie 2: Schrift "Calibri" statt "Neue Kabel" / "Agenda One"
  ...

Soll ich alle Verstöße automatisch korrigieren, oder möchtest du auswählen?
```

### Step 4: Fix (with user approval)
**Option A — Fix all:** Apply all corrections automatically.
**Option B — Fix selectively:** User specifies which slides or violation types.

Replacement strategy:
- Off-brand heading color → replace with Teal With It `#007080`
- Off-brand body text color → replace with Midnight Dreams `#001f33`
- Off-brand fill color → replace with the nearest brand color (ask user if unclear)
- Wrong heading font → replace with `Neue Kabel`
- Wrong body font → replace with `Agenda One`

### Step 5: Save as new file
```python
base = os.path.splitext(original_path)[0]
output = base + '_branded.pptx'
prs.save(output)
```

Confirm: `Korrigierte Datei gespeichert: [path]_branded.pptx`
Report: `X von Y Verstößen korrigiert.`

---

## D — Export to PDF {#export}

### Step 1: Confirm the PPTX path
Either the file just created, or user provides a path.

### Step 2: Check for LibreOffice
```python
import shutil
lo = shutil.which('libreoffice') or shutil.which('soffice')
if not lo:
    print("LibreOffice nicht gefunden.")
    print("Bitte installieren: https://www.libreoffice.org/download/")
    # Stop here and guide user
```

If not found, tell the user and stop. Do not attempt other PDF methods unless user requests.

### Step 3: Convert
```python
import subprocess, os
out_dir = os.path.dirname(os.path.abspath(pptx_path))
subprocess.run([lo, '--headless', '--convert-to', 'pdf', '--outdir', out_dir, pptx_path], 
               check=True, timeout=120)
```

### Step 4: Confirm
```
PDF exportiert: /pfad/zur/datei.pdf
```

---

## User Interaction Principles

- **Always ask before doing** — for every workflow, get confirmation before writing files
- **Checkpoint before save** — show slide summary or violation report before writing
- **Save path from user** — never decide the output path autonomously
- **Never overwrite** — suffix output files with `_extended`, `_branded`, `_copy`, or a date
- **Language** — match the language of the user's request (German preferred per brand)
- **Tone** — professional, clear, direct — per brand tone of voice
