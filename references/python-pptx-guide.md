# python-pptx Code Guide — Die Lobby PowerPoint Skill

Library version: python-pptx 1.0.2
Install path: `/Users/halyna/Library/Python/3.9/lib/python/site-packages/pptx/`

---

## Critical: Opening the POTX Template

python-pptx cannot open `.potx` files directly — the content type `presentationml.template.main+xml` is embedded **inside** the ZIP, not derived from the file extension. Simply renaming or copying to `.pptx` does NOT work.

**Always rewrite the content type inside the ZIP:**

```python
import shutil, tempfile, zipfile, os
from pptx import Presentation
from pptx.oxml.ns import qn

POTX = '/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/2024-10-22_Lobby_PPT-Template.potx'

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

prs, tmp_path = open_potx_as_pptx(POTX)
# ... build slides ...
prs.save('/path/to/output.pptx')
os.unlink(tmp_path)
```

## Removing the Template's Initial Slide

After opening, the POTX always contains 1 placeholder slide. Remove it after adding your content slides:

```python
sldIdLst = prs.slides._sldIdLst
first = sldIdLst[0]
r_id = first.get(qn('r:id'))
sldIdLst.remove(first)
del prs.part._rels[r_id]
```

---

## Mandatory: clean_slide() — Call After Every add_slide()

The POTX layouts contain a phantom `idx=13` subtitle placeholder at `top=0.91"` that sits **above** the real title (`top=2.10"`). If not removed, it shows "Untertitel hinzufügen" prompt text and blocks the top of every slide.

**Always call `clean_slide()` immediately after `add_slide()`:**

```python
from pptx.util import Inches
from pptx.oxml.ns import qn

def shrink_to_fit(tf):
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('shrink', '1')

def clean_slide(slide, layout_name='default'):
    # 1. Remove phantom subtitle placeholder
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx == 13:
            ph._element.getparent().remove(ph._element)

    if layout_name == 'Titelfolie':
        return  # Titelfolie has its own right-side design — leave it

    # 2a. Section dividers (Abschnittsüberschrift with empty body): center title, 40pt
    if 'Abschnitts' in layout_name:
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
        return

    # 2b. Move title up for content slides
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.top = Emu(int(0.8186 * 914400)); ph.left = Inches(1.77)
            ph.width = Inches(9.84); ph.height = Emu(int(1.1698 * 914400))

    # 3. Expand content to fill freed space
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 1:
            ph.top = Emu(int(2.333 * 914400)); ph.height = Emu(int(4.0 * 914400))
            if layout_name == 'Zwei Inhalte':
                ph.left = Inches(1.77); ph.width = Inches(4.67)
            else:
                ph.left = Inches(1.77); ph.width = Inches(9.84)
        elif idx == 2:  # Zwei Inhalte right column
            ph.top = Emu(int(2.333 * 914400)); ph.left = Inches(6.94)
            ph.width = Inches(4.67); ph.height = Emu(int(4.0 * 914400))
```

---

## Adding Slides with Named Layouts

```python
# Build layout lookup once
layout_map = {l.name: l for l in prs.slide_layouts}

# Add a standard content slide
slide = prs.slides.add_slide(layout_map['Titel und Inhalt'])

# Handle the section layout whose name has a newline
section_layout = next(l for l in prs.slide_layouts if 'Abschnitts' in l.name)
section_slide = prs.slides.add_slide(section_layout)

# List all available layout names
for layout in prs.slide_layouts:
    print(repr(layout.name))
```

---

## Filling Placeholders

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor

TEAL    = RGBColor(0x00, 0x70, 0x80)
MIDNIGHT= RGBColor(0x00, 0x1f, 0x33)

def set_title(slide, text, color=TEAL, size_pt=28):
    title_ph = slide.shapes.title
    if title_ph is None:
        return
    title_ph.text = text
    for para in title_ph.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = 'Neue Kabel'
            run.font.bold = True
            run.font.size = Pt(size_pt)
            run.font.color.rgb = color

def set_body(slide, text, color=MIDNIGHT, size_pt=16, ph_idx=1):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == ph_idx:
            ph.text = text
            for para in ph.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = 'Agenda One'
                    run.font.bold = False
                    run.font.size = Pt(size_pt)
                    run.font.color.rgb = color
            break

def set_bullet_list(slide, items: list[str], ph_idx=1):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == ph_idx:
            tf = ph.text_frame
            tf.clear()
            for i, item in enumerate(items):
                para = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                run = para.add_run()
                run.text = item
                run.font.name = 'Agenda One'
                run.font.size = Pt(16)
                run.font.color.rgb = MIDNIGHT
            break
```

---

## Setting Slide Background Color

```python
from pptx.util import Pt
from pptx.oxml.ns import qn
from lxml import etree

def set_slide_background(slide, rgb: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb
```

---

## Inserting an Image

```python
from pptx.util import Inches

def add_image(slide, image_path: str, left_in, top_in, width_in, height_in=None):
    left   = Inches(left_in)
    top    = Inches(top_in)
    width  = Inches(width_in)
    height = Inches(height_in) if height_in else None
    slide.shapes.add_picture(image_path, left, top, width, height)
```

---

## Adding a Colored Rectangle

```python
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

def add_colored_box(slide, rgb: RGBColor, left_in, top_in, width_in, height_in):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left_in), Inches(top_in),
        Inches(width_in), Inches(height_in)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    shape.line.fill.background()  # no border
    return shape
```

---

## Scanning for Brand Violations

```python
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor

BRAND_HEX = {'e50040','b2003b','007080','00dee0','d4ff4d','001f33','000000','ffffff'}

def audit_presentation(prs) -> list[dict]:
    violations = []
    
    for slide_idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            # Check fill color
            try:
                if shape.fill.type == 1:  # SOLID
                    rgb = str(shape.fill.fore_color.rgb).lower()
                    if rgb not in BRAND_HEX:
                        violations.append({
                            'slide': slide_idx,
                            'shape': shape.name,
                            'issue': 'fill',
                            'value': f'#{rgb}'
                        })
            except (AttributeError, TypeError):
                pass
            
            # Check text runs
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        if run.font.color.type is not None:
                            rgb = str(run.font.color.rgb).lower()
                            if rgb not in BRAND_HEX:
                                violations.append({
                                    'slide': slide_idx,
                                    'shape': shape.name,
                                    'issue': 'text_color',
                                    'value': f'#{rgb}'
                                })
                    except (AttributeError, TypeError):
                        pass
                    
                    if run.font.name and run.font.name not in ('Neue Kabel', 'Agenda One'):
                        violations.append({
                            'slide': slide_idx,
                            'shape': shape.name,
                            'issue': 'font',
                            'value': run.font.name
                        })
    
    return violations
```

---

## PDF Export

Check for LibreOffice first:

```python
import subprocess, shutil, os

def export_to_pdf(pptx_path: str) -> str:
    """Returns path to generated PDF, or raises if LibreOffice not found."""
    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if not lo:
        raise RuntimeError('LibreOffice not found. Install from https://libreoffice.org')
    
    out_dir = os.path.dirname(os.path.abspath(pptx_path))
    result = subprocess.run(
        [lo, '--headless', '--convert-to', 'pdf', '--outdir', out_dir, pptx_path],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(f'LibreOffice error: {result.stderr}')
    
    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + '.pdf'
    return os.path.join(out_dir, pdf_name)
```

---

## Complete Minimal Example — New Presentation

```python
import shutil, tempfile, os, json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

POTX = '/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/2024-10-22_Lobby_PPT-Template.potx'
BRAND = json.load(open('/Users/halyna/Documents/CLAUDE/Skill-powerpoint/brand-kit.json'))

TEAL    = RGBColor(0x00, 0x70, 0x80)
MIDNIGHT= RGBColor(0x00, 0x1f, 0x33)
WHITE   = RGBColor(0xff, 0xff, 0xff)

# Open template
tmp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
tmp.close()
shutil.copy(POTX, tmp.name)
prs = Presentation(tmp.name)

# Remove the placeholder slide that comes with the template (index 0)
# Only if it exists and you want a clean start
xml_slides = prs.slides._sldIdLst
if len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].get('r:id') if False else None  # keep template slides

layout_map = {l.name: l for l in prs.slide_layouts}

# Cover slide
cover = prs.slides.add_slide(layout_map['Titelfolie'])
cover.shapes.title.text = 'Meine Präsentation'
for ph in cover.placeholders:
    if ph.placeholder_format.idx == 1:
        ph.text = 'Untertitel · Die Lobby · 2026'

# Content slide
content = prs.slides.add_slide(layout_map['Titel und Inhalt'])
content.shapes.title.text = 'Übersicht'
for ph in content.placeholders:
    if ph.placeholder_format.idx == 1:
        ph.text = 'Hier kommt der Inhalt'

# Save
output_path = '/path/to/output.pptx'
prs.save(output_path)
os.unlink(tmp.name)
print(f'Saved: {output_path}')
```
