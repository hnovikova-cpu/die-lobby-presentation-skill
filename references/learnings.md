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

---

## Entries from 2026-05-15 — Audit & Fix of existing PPTX (lego-ki-workshop)

### 2026-05-15 — Bold detection: font name vs. explicit bold attribute
**Situation:** Rebranding an existing PPTX where all fonts were `Gill Sans MT Bold`. Tried to detect bold runs via `run.font.bold`.
**Discovery:** When the original file set bold by choosing a "Bold" font variant (e.g. `Gill Sans MT Bold`) instead of setting `<a:rPr b="1">`, `run.font.bold` returns `None` — not `True`. The bold is encoded in the font name, not the attribute.
**Fix / Rule:** Always check BOTH `run.font.bold is True` AND `'Bold' in original_font_name`. Capture original font names before making any changes:
```python
orig_fonts = {}
for si, slide in enumerate(prs_orig.slides):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        for pi, para in enumerate(shape.text_frame.paragraphs):
            for ri, run in enumerate(para.runs):
                orig_fonts[(si, shape.shape_id, pi, ri)] = (run.font.name or '', run.font.bold)

def is_bold(fn_orig, bold_attr):
    return bold_attr is True or ('Bold' in fn_orig if fn_orig else False)
```

### 2026-05-15 — Triangle decoration: exact XML format required
**Situation:** Added `rtTriangle` shapes to slides via raw XML. PowerPoint opened the file with an error and auto-repaired it (removing the triangles).
**Discovery:** Two issues caused the corruption:
1. `<a:spLocks noGrp="1"/>` inside `<p:cNvSpPr>` — PowerPoint rejects this for non-placeholder shapes
2. Missing `userDrawn="1"` attribute on `<p:nvPr>` — required for shapes added programmatically
**Fix / Rule:** Always use this exact XML template for triangles (mirrors the reference file `Samu_Barriererfreiheit.pptx`):
```python
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def make_triangle(shape_id, name, off_x, off_y, cx, cy, color_hex, rot=None, flipV=False):
    xfrm_attrs = (f' rot="{rot}"' if rot else '') + (' flipV="1"' if flipV else '')
    xml = f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
  <p:nvSpPr>
    <p:cNvPr id="{shape_id}" name="{name}"/>
    <p:cNvSpPr/>
    <p:nvPr userDrawn="1"/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm{xfrm_attrs}>
      <a:off x="{off_x}" y="{off_y}"/>
      <a:ext cx="{cx}" cy="{cy}"/>
    </a:xfrm>
    <a:prstGeom prst="rtTriangle"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="de-DE" dirty="0"/></a:p>
  </p:txBody>
</p:sp>'''
    from lxml import etree
    return etree.fromstring(xml)
```

### 2026-05-15 — Die Lobby triangle decoration: exact dimensions and colors
**Situation:** Needed to match the triangle decorations from the official template (`Samu_Barriererfreiheit.pptx` master).
**Discovery:** Exact values extracted from the slide master XML:
- **Top-left triangle:** `rtTriangle`, `flipV="1"`, pos=(0,0), size=3200400×964692 EMU, color = **Tennis Ball `#d4ff4d`**
- **Bottom-right triangle:** `rtTriangle`, `rot="10800000" flipV="1"`, pos=(8991600, 5893308), same size, color = **Teal With It `#007080`**
**Fix / Rule:** Always use these exact values. Insert TL at index 0 of spTree (behind all content), append BR at the end. Use `get_max_id(slide) + 1` for shape IDs to avoid duplicates:
```python
TRI_CX, TRI_CY = 3200400, 964692
BR_X, BR_Y = 8991600, 5893308

def get_max_id(slide):
    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    ids = []
    for elem in slide.shapes._spTree:
        pr = elem.find(f'.//{{{P}}}cNvPr')
        if pr is not None:
            try: ids.append(int(pr.get('id', 0)))
            except: pass
    return max(ids) if ids else 100

# Add to each slide:
base = get_max_id(slide) + 1
spTree.insert(0, make_triangle(base,   'Triangle TL', 0,    0,    TRI_CX, TRI_CY, 'd4ff4d', flipV=True))
spTree.append(  make_triangle(base+1, 'Triangle BR', BR_X, BR_Y, TRI_CX, TRI_CY, '007080', rot=10800000, flipV=True))
```

### 2026-05-15 — Text size depends on content density; 15pt hard minimum
**Situation:** The lego-ki-workshop presentation has compact slides (few items per slide), while earlier presentations were denser. Original fonts were 11–18pt with no consistent rule.
**Discovery:** Sub-headers (bold labels like "Aufgabe:", "Warm-Up") should be 20pt on compact slides. 15pt is the hard minimum for all body text. Footnotes/captions minimum is 13pt. Never apply these rules to footer text, page numbers, or emoji runs.
**Fix / Rule:** Count text blocks per slide. ≤5 = compact → sub-headers 20pt, body 18pt. >5 = dense → sub-headers 16–18pt, body min 15pt. See `brand-rules.md` → "Font Sizes — Dynamic by Content Density" for the audit fix pattern.

### 2026-05-15 — Content left alignment: all slides must start at x=150
**Situation:** Slides 3–9 had content starting at x=40–50 while slides 1&2 started at x=150.
**Discovery:** x=150 is the canonical left content margin for Die Lobby presentations. It must be consistent across every slide. Right-side decorative elements (emoji icons at x≥900), footer (x=966), and page number (x=1260) are excluded. After shifting, always check right-side overflow (max content right edge = x=1313).
**Fix / Rule:** See `brand-rules.md` → "Content Start X — Critical Alignment Rule" for the exact audit/fix code. Apply this check as part of every Workflow C (audit) and Workflow A (create) run.

### 2026-05-15 — Validate PPTX XML before considering done
**Situation:** Saved a PPTX that python-pptx loaded fine, but PowerPoint reported an error.
**Discovery:** python-pptx's parser is more lenient than PowerPoint's. A file can be "valid" in python-pptx but still fail in PowerPoint due to bad shape attributes.
**Fix / Rule:** After every save, always validate the full ZIP contents:
```python
import os
from lxml import etree

def validate_pptx(path):
    import zipfile
    errors = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.endswith('.xml') or name.endswith('.rels'):
                try: etree.fromstring(z.read(name))
                except Exception as e: errors.append(f'{name}: {e}')
    # Also check for duplicate shape IDs per slide
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
                tree = etree.fromstring(z.read(name))
                ids = [el.get('id') for el in tree.iter() if el.tag.endswith('}cNvPr') and el.get('id')]
                dupes = {x for x in ids if ids.count(x) > 1}
                if dupes: errors.append(f'{name}: duplicate IDs {dupes}')
    return errors  # empty = clean
```

### 2026-05-15 — PowerPoint auto-repair removes unknown/corrupt shapes
**Situation:** PowerPoint opened the branded file with a repair dialog and removed the Triangle BR shapes from slides 1–9.
**Discovery:** When PowerPoint auto-repairs a file, it saves a `*_Repariert.pptx` (macOS German UI) or `*_Repaired.pptx` (English UI) in the same folder. The repaired file is the user's working version — always use it as the new base.
**Fix / Rule:** When user reports a PowerPoint error + provides a `_Repariert` file:
1. Use the `_Repariert` file as the new source
2. Inspect what was removed (compare shape lists)
3. Re-add only the missing elements — do not re-apply all fixes
4. Re-validate before saving
