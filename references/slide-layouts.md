# Slide Layouts — Die Lobby POTX Template

Official template: `/Users/halyna/Documents/CLAUDE/Skill-powerpoint/input/2024-10-22_Lobby_PPT-Template.potx`

14 layouts are available. Use the **exact German name** when calling `layout_map['Name']` in python-pptx.

---

## Layout Reference

| # | Exact Name | When to Use |
|---|---|---|
| 1 | `Titelfolie` | Cover slide — first slide of every presentation. Title + subtitle + date/author |
| 2 | `Referenz Device Mockup mit Liste` | Show a device screenshot (phone/tablet/browser) alongside a bullet point list |
| 3 | `Referenz Device Mockup` | Device screenshot only — no text alongside, image fills the slide |
| 4 | `Screenshot in Fenster` | Browser or app window screenshot in a framed window graphic |
| 5 | `Fenster für Content auf Weiß` | White background with a framed content area — for diagrams, tables, long text |
| 6 | `Titel mit Bild` | Section slide with title on left + image on right |
| 7 | `Bild mit Titel` | Image on left + title/text on right |
| 8 | `1_Titel mit Bild` | Alternate variant of layout 6 — use when you've already used layout 6 on a prior slide |
| 9 | `1_Bild mit Titel` | Alternate variant of layout 7 |
| 10 | `Titel und Inhalt` | Standard workhorse: title at top + body text/bullets below. Most common layout |
| 11 | `Abschnitts-\nüberschrift` | Section divider slide — large text only, no body content. Use between major chapters |
| 12 | `Zwei Inhalte` | Two-column layout — side-by-side content blocks |
| 13 | `Vergleich` | Comparison layout — two columns with headers (e.g. Before/After, Option A/B) |
| 14 | `Inhalt mit Überschrift` | Heading on left, content area on right |

**Note:** Layout 11's name contains a newline in the XML (`Abschnitts-\nüberschrift`). Access it safely:

```python
layout_map = {l.name: l for l in prs.slide_layouts}
# Handle the newline variant
section_layout = next(l for l in prs.slide_layouts if 'Abschnitts' in l.name)
```

---

## Slide Structure Recommendations

### Typical presentation flow

```
Slide 1:  Titelfolie              ← always
Slide 2:  Titel und Inhalt        ← agenda / overview
Slide N:  Abschnittsüberschrift   ← before each chapter
Slide N+: content layouts         ← Titel und Inhalt, Zwei Inhalte, etc.
Last:     Titelfolie              ← closing/thank-you (reuse title layout)
```

### When to use which content layout

- **Plain text/bullets:** `Titel und Inhalt` (layout 10)
- **Two topics side by side:** `Zwei Inhalte` (layout 12)
- **Comparing options:** `Vergleich` (layout 13)
- **Image with explanation:** `Titel mit Bild` or `Bild mit Titel` (6, 7, 8, 9)
- **App/website demo:** `Screenshot in Fenster` (4) or `Referenz Device Mockup` (3)
- **App demo + bullets:** `Referenz Device Mockup mit Liste` (2)
- **Diagram/table:** `Fenster für Content auf Weiß` (5)
- **Chapter break:** `Abschnittsüberschrift` (11)

---

## Accessing Placeholders

After adding a slide, fill placeholders by index or type:

```python
from pptx.util import Pp
from pptx.enum.text import PP_ALIGN

slide = prs.slides.add_slide(layout_map['Titel und Inhalt'])

for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.name)
    # idx=0 → title, idx=1 → content/body
```

Common placeholder indices:
- `0` — Title
- `1` — Content / body
- `2` — Subtitle (on Titelfolie)
- `10+` — Custom placeholders (images, footers)
