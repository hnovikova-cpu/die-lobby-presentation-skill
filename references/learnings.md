# Learnings — Die Lobby PowerPoint Skill

This file captures real-world discoveries made during actual presentation creation sessions. Update it whenever something unexpected happens — a python-pptx quirk, a layout behavior, a brand decision, or a workflow improvement.

---

## Format

Each entry:
```
### YYYY-MM-DD — Short title
**Situation:** what we were trying to do
**Discovery:** what we found
**Fix / Rule:** what to do going forward
```

---

## Entries

### 2026-04-24 — POTX cannot be opened by copying to .pptx
**Situation:** Tried to open the POTX template by copying to a temp `.pptx` file.
**Discovery:** python-pptx still refuses with `content type is 'presentationml.template.main+xml'` — the content type is embedded in `[Content_Types].xml` inside the ZIP, not derived from the file extension.
**Fix / Rule:** Must rewrite the content type string inside the ZIP before opening with python-pptx:
```python
import shutil, tempfile, zipfile, os
from pptx import Presentation

def open_potx_as_pptx(potx_path):
    tmp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
    tmp.close()
    with zipfile.ZipFile(potx_path, 'r') as zin, zipfile.ZipFile(tmp.name, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                data = data.replace(
                    b'presentationml.template.main+xml',
                    b'presentationml.presentation.main+xml'
                )
            zout.writestr(item, data)
    return Presentation(tmp.name), tmp.name
```

### 2026-04-24 — POTX template contains 1 placeholder slide
**Situation:** After opening the POTX template, it always has 1 existing slide.
**Discovery:** The template slide is a design placeholder — it must be removed before saving, or it will appear as slide 1 in the final output.
**Fix / Rule:** After adding all content slides, remove the initial template slide using XML manipulation:
```python
from pptx.oxml.ns import qn
sldIdLst = prs.slides._sldIdLst
first = sldIdLst[0]
r_id = first.get(qn('r:id'))
sldIdLst.remove(first)
prs.part._rels.pop(r_id)  # use .pop(), not del — _Relationships doesn't support item deletion
```

### 2026-04-24 — POTX layouts have a phantom idx=13 subtitle placeholder at the top
**Situation:** Generated slides had empty "Untertitel hinzufügen" block consuming the top ~1" of each slide, pushing titles down to y=2.10".
**Discovery:** Every layout (except Titelfolie) has an `idx=13` "Subtitle 2" placeholder at `top=0.91"` — above the title at `top=2.10"`. When left unfilled it renders as prompt text and wastes vertical space.
**Fix / Rule:** After `add_slide()`, always remove idx=13 and reposition idx=0 (title) to top=0.45":
```python
for ph in list(slide.placeholders):
    if ph.placeholder_format.idx == 13:
        ph._element.getparent().remove(ph._element)
# Then move title up:
for ph in slide.placeholders:
    if ph.placeholder_format.idx == 0:
        ph.top = Inches(0.45); ph.height = Inches(1.2)
```

### 2026-04-24 — Correct placeholder geometry (confirmed by user correction on slide 3)
**Situation:** User manually corrected slide 3 and saved as `_corrected.pptx`. Diff revealed the exact values they want.
**Discovery:** The correct geometry (measured from user's correction):
- Title (idx=0): `top=0.8186"`, `height=1.1698"`, `left=1.77"`, `width=9.84"`
- Content (idx=1 / idx=2): `top=2.333"`, `height=4.0"` — content ends at 6.33", safely above footer at 6.63"
- Font size: **16pt** for all body text (user bumped 13pt → 16pt; shrink-to-fit handles overflow)
**Rule:** Always use these values as defaults. Do not use 0.40" for title top — it sits too high. Do not use 2.10" for content top — it overlaps with long titles.

### 2026-04-24 — User correction workflow: diff _corrected.pptx to extract exact values
**Situation:** User made manual fixes and saved with `_corrected` suffix.
**Discovery:** `ph_info()` diff between original and corrected file is the fastest way to extract exact geometry changes:
```python
def ph_info(slide):
    info = {}
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        info[idx] = {'top': round(ph.top/914400,4), 'left': round(ph.left/914400,4),
                     'width': round(ph.width/914400,4), 'height': round(ph.height/914400,4)}
        try: info[idx]['font_pt'] = ph.text_frame.paragraphs[0].runs[0].font.size.pt
        except: info[idx]['font_pt'] = None
    return info
```
**Rule:** When user says "I corrected X, apply to all" — always diff first, extract values, then regenerate from scratch.

### 2026-04-24 — Content placeholder geometry after removing idx=13
**Situation:** After removing idx=13, content area starts at top=3.32" — too low, content overflows.
**Fix / Rule:** After removing idx=13 and repositioning title, reposition content placeholders to start at top=1.85" with height=4.45". For Zwei Inhalte split: left at left=1.77" w=4.67", right at left=6.94" w=4.67".

### 2026-04-24 — Shrink text to fit via XML bodyPr attribute
**Situation:** Long bullet lists overflow placeholder boundaries.
**Discovery:** `tf.auto_size` property doesn't always work via python-pptx API. Direct XML is reliable:
```python
bodyPr = tf._txBody.find(qn('a:bodyPr'))
if bodyPr is not None:
    bodyPr.set('shrink', '1')
```

### 2026-04-24 — Zwei Inhalte layout placeholder indices
**Situation:** Needed to fill two content columns on a slide.
**Discovery:** `Zwei Inhalte` layout has: idx=0 (title), idx=1 (left content), idx=2 (right content).

### 2026-04-24 — Abschnittsüberschrift layout name contains newline
**Situation:** Tried `layout_map['Abschnittsüberschrift']` — KeyError.
**Discovery:** The XML stores the name as `'Abschnitts-\nüberschrift'` with a literal newline.
**Fix / Rule:** Always access this layout by substring match:
```python
section_layout = next(l for l in prs.slide_layouts if 'Abschnitts' in l.name)
```

### 2026-04-27 — Section divider slides: center title vertically, remove empty body placeholder
**Situation:** Abschnittsüberschrift slides had title at top=0.82" with an empty idx=1 body placeholder below.
**Discovery:** User corrected slide 12 — moved title to top=2.5802" (vertically centered), 40pt font, removed the empty idx=1 placeholder entirely. Applied to all section divider slides (those where idx=1 text is empty).
**Fix / Rule:** After `add_slide()` for section dividers, remove empty body and reposition title:
```python
for ph in list(slide.placeholders):
    if ph.placeholder_format.idx == 1 and not ph.text_frame.text.strip():
        ph._element.getparent().remove(ph._element)
for ph in slide.placeholders:
    if ph.placeholder_format.idx == 0:
        ph.top = Emu(int(2.5802 * 914400)); ph.left = Emu(int(1.87 * 914400))
        ph.width = Inches(9.84); ph.height = Emu(int(1.4198 * 914400))
        for para in ph.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = Pt(40)
```

### 2026-04-27 — Content box left stripe width: 0.09" is too thick
**Situation:** Custom content boxes used a 0.09"-wide left stripe for visual accent.
**Discovery:** User thinned all left stripes from 0.09" → 0.05" — looks more refined and less heavy.
**Fix / Rule:** Always use `w=0.05"` for left accent stripes on content boxes. Applies to all patterns: 3-column boxes, left/right project detail, any custom content box with a left stripe.

### 2026-04-27 — 3-column content boxes: height 4.0" is too tall
**Situation:** 3-column boxes (slide 13 pattern) used full height 4.0" starting at top=2.33".
**Discovery:** User reduced box height to 2.077" — boxes end at ~4.41" leaving breathing room before footer. Body text top moved from 2.93" → 3.124" within the shorter box.
**Fix / Rule:** For 3-column content box pattern: `box_h=2.077"`, body textbox top offset from box top = ~0.794" (not 0.6").

### 2026-04-27 — Project detail left/right layout: corrected geometry
**Situation:** Left/right project detail slides (Kontext left + 3 right boxes) had right boxes misaligned.
**Discovery:** User corrected slide 15. Correct geometry:
- Left box: `h=4.193"` (not 4.0")
- Right box 2 (Ergebnis, top=3.65"): `h=1.355"` (not 1.22")
- Right box 3 (Relevanz): `top=5.105"` (not 4.97"), `h=1.411"` (not 1.36")
- Right title textboxes shifted: box1 top 2.48→2.402, box2 top 3.8→3.767, box3 top 5.12→5.215
- Right body textboxes shifted: box1 top 2.93→2.797, box2 top 4.25→4.194, box3 top 5.57→5.618
**Fix / Rule:** Use these corrected values when generating project detail slides (pattern used on slides 15, 17, 19).

### 2026-04-27 — Diff workflow for custom shapes (not just placeholders)
**Situation:** Assumed user corrections on slides 13 and 15 were placeholder-only — found nothing in ph_info() diff.
**Discovery:** User corrections were in custom AUTO_SHAPE and TEXT_BOX elements, not placeholders. Must iterate `slide.shapes` (not just `slide.placeholders`) to find changes.
**Fix / Rule:** When diffing a _corrected.pptx, always compare all shapes by type+position, not just placeholders.
