# Brand Rules — Die Lobby PowerPoint

Source of truth: `/Users/halyna/Documents/CLAUDE/Skill-powerpoint/brand-kit.json`

Always read `brand-kit.json` at the start of every session. The values below reflect version 1.0.0 (last updated 2026-04-20).

---

## Color Palette

```python
from pptx.dml.color import RGBColor

# Primary
NEON_ROMANCE   = RGBColor(0xe5, 0x00, 0x40)  # #e50040 — accent, CTA
WILD_BERRY     = RGBColor(0xb2, 0x00, 0x3b)  # #b2003b — deep red accent

# Secondary
TEAL_WITH_IT   = RGBColor(0x00, 0x70, 0x80)  # #007080 — primary heading color
FROZEN_BOUBBLE = RGBColor(0x00, 0xde, 0xe0)  # #00dee0 — backgrounds/accents
TENNIS_BALL    = RGBColor(0xd4, 0xff, 0x4d)  # #d4ff4d — highlight, signal
MIDNIGHT_DREAMS= RGBColor(0x00, 0x1f, 0x33)  # #001f33 — body text, dark BG

# Neutral
BLACK          = RGBColor(0x00, 0x00, 0x00)  # #000000 — logo only
WHITE          = RGBColor(0xff, 0xff, 0xff)  # #ffffff — default background
```

### Allowed Hex Values (for audit comparison)

```python
BRAND_COLORS = {
    '#e50040', '#b2003b',                    # primary reds
    '#007080', '#00dee0', '#d4ff4d', '#001f33',  # secondary
    '#000000', '#ffffff',                    # neutrals
}
```

---

## Color Usage Rules

### Backgrounds
| Background | When to use |
|---|---|
| White `#ffffff` | Default — most slides |
| Midnight Dreams `#001f33` | Dark slides, section dividers for emphasis |
| Frozen Boubble `#00dee0` | Accent backgrounds, callout boxes |
| Tennis Ball `#d4ff4d` | High-energy highlight slides |

### Text on backgrounds
| Background | Heading color | Body text color |
|---|---|---|
| White | Teal With It `#007080` | Midnight Dreams `#001f33` |
| Frozen Boubble | Midnight Dreams `#001f33` | Midnight Dreams `#001f33` |
| Tennis Ball | Midnight Dreams `#001f33` | Midnight Dreams `#001f33` |
| Midnight Dreams | Frozen Boubble `#00dee0` | White `#ffffff` |
| Black | White `#ffffff` | White `#ffffff` |

### Accent / CTA elements
- Buttons, highlighted boxes, key labels: **Neon Romance `#e50040`**
- Icon backgrounds or dividers: **Wild Berry `#b2003b`**
- Signal/warning callouts: **Tennis Ball `#d4ff4d`**

### Logo backgrounds (logo slides only)
Only place the logo on: White, Black, or Midnight Dreams.

---

## Typography

### Font names (exact strings for python-pptx)

```python
FONT_HEADING   = 'Neue Kabel'       # weight: ExtraBold (800)
FONT_BODY      = 'Agenda One'       # weight: Medium (500)
FONT_EMPHASIS  = 'Agenda One'       # weight: Bold (700)

# Fallback chain if brand fonts not installed:
FONT_FALLBACK_HEADING = ['Futura', 'Century Gothic', 'Helvetica', 'Arial']
FONT_FALLBACK_BODY    = ['Helvetica Neue', 'Helvetica', 'Arial']
```

### Applying fonts in python-pptx

**Always pass `on_dark=True` when the slide background is Midnight Dreams.**

```python
from pptx.util import Pt

def style_heading(run, size_pt=28, on_dark=False):
    run.font.name = 'Neue Kabel'
    run.font.bold = True
    run.font.size = Pt(size_pt)
    run.font.color.rgb = FROZEN_BOUBBLE if on_dark else TEAL_WITH_IT

def style_body(run, size_pt=16, on_dark=False):
    run.font.name = 'Agenda One'
    run.font.bold = False
    run.font.size = Pt(size_pt)
    run.font.color.rgb = WHITE if on_dark else MIDNIGHT_DREAMS

def set_dark_bg(slide):
    """Set slide background to Midnight Dreams. Call before writing any text."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = MIDNIGHT_DREAMS
```

**Contrast quick-reference — before writing ANY text, pick the right pair:**

| Slide background | `on_dark` | Heading color | Body color |
|---|---|---|---|
| White `#ffffff` | `False` | Teal With It `#007080` | Midnight Dreams `#001f33` |
| Midnight Dreams `#001f33` | `True` | Frozen Boubble `#00dee0` | White `#ffffff` |
| Frozen Boubble `#00dee0` | `False` | Midnight Dreams `#001f33` | Midnight Dreams `#001f33` |
| Tennis Ball `#d4ff4d` | `False` | Midnight Dreams `#001f33` | Midnight Dreams `#001f33` |

**Rule:** Never put Midnight Dreams text on a Midnight Dreams background — it's invisible. Never put White text on a White background. When in doubt, set `on_dark` based on whether `set_dark_bg()` was called on that slide.

### Font Sizes — Dynamic by Content Density

Before setting sizes, count visible text blocks per slide:
- **≤ 5 text blocks → COMPACT layout**
- **> 5 text blocks → DENSE layout**

| Element | Compact slide | Dense slide | Hard minimum |
|---|---|---|---|
| Cover title | 40–44 pt | 36–40 pt | 36 pt |
| Slide title | 28–32 pt | 24–28 pt | 24 pt |
| Section heading | 44–48 pt | 40–44 pt | 40 pt |
| Sub-header / label | **20 pt** | 16–18 pt | 16 pt |
| Body text | **18 pt** | 15–16 pt | **15 pt** |
| Footnote / caption | 13 pt | 13 pt | 13 pt |

**Hard rules:**
- Never set body or label text below **15 pt**
- Never set footnotes/captions below **13 pt**
- Sub-headers on compact slides are always **20 pt**
- Footer (`www.die-lobby.de`), page numbers, emojis — excluded from these rules

**Audit fix pattern:**
```python
for run in all_runs:
    if is_footer(run) or is_page_number(run) or is_emoji(run): continue
    if run.font.size is None: continue
    pt = run.font.size.pt
    if run.font.bold and pt < 20 and pt < 28:   # sub-header
        run.font.size = Pt(20)
    elif not run.font.bold and pt < 15:          # body text
        run.font.size = Pt(15)
```

---

## Layout & Spacing

- **Slide size:** 13.33 × 7.50 inches (12,192,000 × 6,858,000 EMU)
- **Margin (safe zone):** 0.4 inches from each edge
- **White space:** leave breathing room — avoid filling every pixel
- **Max 3 dominant colors per slide**
- **Headlines:** max 2 lines, bold, concise

### Content Start X — Critical Alignment Rule

**All content on every slide must start at x = 150 units (1,371,600 EMU / ~1.5 inches).**

This is the single most important layout rule. Confirmed from the reference presentation and user correction across 9 slides.

```python
CONTENT_LEFT_X = 150        # display units (÷ 9144 from EMU)
CONTENT_LEFT_EMU = 1371600  # 150 * 9144

# Right margin: keep all content right edge ≤ x=1313 (20px from slide edge)
CONTENT_RIGHT_MAX = 1313
SLIDE_WIDTH_UNITS = 1333
```

**Applies to:**
- Section number / badge (e.g. "01", "07")
- Slide title / heading
- Body text boxes
- Content card left edges
- Bottom footnote / question text

**Does NOT apply to:**
- Footer (`www.die-lobby.de`) — fixed at x=966
- Page number — fixed at x=1260
- Right-side decorative emoji/icons — stay at x≥900
- Triangle decorations — anchored at (0,0) and (983,644)

**When auditing or fixing an existing file:**
```python
SKIP_NAMES = {'Triangle TL', 'Triangle BR', '', 'Text 2', 'Text 3'}  # footer/page-num
CONTENT_LEFT = 150

for elem in slide.shapes._spTree:
    pr = elem.find(f'.//{{{P}}}cNvPr')
    if pr is None or pr.get('name','') in SKIP_NAMES: continue
    xfrm = elem.find(f'.//{{{A}}}xfrm')
    if xfrm is None: continue
    off = xfrm.find(f'{{{A}}}off')
    if off is None: continue
    x = int(off.get('x')) // 9144
    if x >= 900: continue  # right-side decorative
    if x < CONTENT_LEFT:
        delta = CONTENT_LEFT - x
        off.set('x', str((x + delta) * 9144))
```

**After shifting, always check for right-side overflow:**
```python
for elem in slide.shapes._spTree:
    ...
    right = x + cx
    if right > CONTENT_RIGHT_MAX:
        ext.set('cx', str((CONTENT_RIGHT_MAX - x) * 9144))
```

---

## Brand Violation Checklist (for Audit mode)

When scanning an existing PPTX, flag these as violations:

```python
def is_brand_color(rgb_hex: str) -> bool:
    """Check if a hex color string is in the brand palette."""
    return rgb_hex.lower() in BRAND_COLORS

# Check each shape's fill and text color
for slide in prs.slides:
    for shape in slide.shapes:
        # 1. Fill color
        if shape.fill.type == PP_FILL.SOLID:
            hex_val = '#' + str(shape.fill.fore_color.rgb)
            if not is_brand_color(hex_val):
                flag(f'Slide {slide.slide_number}: off-brand fill {hex_val} on "{shape.name}"')
        
        # 2. Text color + font
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.color.type:
                        hex_val = '#' + str(run.font.color.rgb)
                        if not is_brand_color(hex_val):
                            flag(f'Off-brand text color {hex_val}')
                    if run.font.name and run.font.name not in ('Neue Kabel', 'Agenda One'):
                        flag(f'Off-brand font: {run.font.name}')
```

### Common violations to look for
- Font substitutions (Calibri, Arial, Times New Roman used instead of brand fonts)
- RGB colors not in the 8-color palette
- White text on white background (invisible)
- Midnight Dreams text on Midnight Dreams background (invisible)
- Headings not in Teal With It on light slides
- Body text not in Midnight Dreams on light slides / not in White on dark slides
