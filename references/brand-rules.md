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

```python
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

def style_heading(run, size_pt=28):
    run.font.name = 'Neue Kabel'
    run.font.bold = True
    run.font.size = Pt(size_pt)
    run.font.color.rgb = TEAL_WITH_IT

def style_body(run, size_pt=16):
    run.font.name = 'Agenda One'
    run.font.bold = False
    run.font.size = Pt(size_pt)
    run.font.color.rgb = MIDNIGHT_DREAMS
```

### Recommended font sizes

| Element | Size |
|---|---|
| Cover title | 36–44 pt |
| Slide title | 24–32 pt |
| Section heading | 40–48 pt |
| Body text | 14–18 pt |
| Captions / footnotes | 10–12 pt |

---

## Layout & Spacing

- **Slide size:** 13.33 × 7.50 inches (12,192,000 × 6,858,000 EMU)
- **Margin (safe zone):** 0.4 inches from each edge
- **White space:** leave breathing room — avoid filling every pixel
- **Max 3 dominant colors per slide**
- **Headlines:** max 2 lines, bold, concise

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
