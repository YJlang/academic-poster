---
name: academic-poster
description: Builds print-ready academic/conference posters (A0/A1/A2) as an editable PPTX whose flow layout engine measures every paragraph from font metrics, so a section bar can never overlap or overflow the text above it — overflow-proof by construction, no fixed coordinates to hand-tune. Also native editable charts and tables, gradient/shadow styling, CJK font handling, and a LibreOffice render-and-check loop. Use when the user wants to build a research or conference poster, a large single-slide PPTX, or asks about poster size, columns, sections, figures, tables, charts, equations, fonts, or fixing text that overflows or gets clipped on a poster.
license: MIT
---

# Academic poster builder

Build a research/conference poster as a single oversized, **editable** PPTX
slide. Content (text, theme, figures, tables, charts) lives in a config; the
engine only does layout. The defining feature is a **flow layout that measures
text**, which is what stops sections from overlapping when content is long.

## When to use
- "Make an A1/A0 conference poster", "research poster as PPTX/PowerPoint".
- A large single-slide PPTX with header, multi-column sections, figures, tables.
- Fixing a poster where text overflows, a title box covers a paragraph, fonts
  look inconsistent, or the file opens read-only.

Not for: normal multi-slide decks (use a slides skill), or non-PPTX formats.

## Dependencies
```bash
pip install python-pptx Pillow PyYAML
```
Previewing/exporting also needs **LibreOffice** (`soffice`).

## Quick start
```bash
# 1) generate from a config (all dummy content to start)
python examples/demo_poster.py examples/config.example.yaml out.pptx
# 2) ALWAYS render and look
python scripts/render_preview.py out.pptx
# 3) (macOS) if PowerPoint opens it read-only:
scripts/unlock_macos.sh out.pptx
```

## Workflow
Copy this checklist and work through it:
```
- [ ] 1. Pick paper size + orientation (A1 portrait is the common default)
- [ ] 2. Write the config: header, abstract, left[] / right[] blocks, footer
- [ ] 3. Generate the pptx (Poster(config).save / demo_poster.py)
- [ ] 4. Render a PNG and LOOK at it (render_preview.py)
- [ ] 5. Fix overflow/clipping/balance, regenerate, repeat until both columns
         end just above the footer with no overlap
- [ ] 6. (macOS) strip quarantine so it opens editable; export final PDF from
         the target app for print
```

## Core principles (do these by default)
1. **Flow, never fixed `y`.** Lay blocks top-to-bottom; each returns its height.
   Measure body text from font metrics before drawing. This is the #1 fix for
   "the section bar covers the text." See [reference/layout-engine.md](reference/layout-engine.md).
2. **Measure with the font you draw with.** Set `theme.font_path` to the same
   font as `theme.font`, or predicted height won't match. Keep a small safety
   margin (over-estimate slightly) so the real app never overflows.
3. **Turn off text-box autofit** (`MSO_AUTO_SIZE.NONE`) so the app can't shrink
   text — otherwise letters look inconsistent. See [reference/fonts-rendering.md](reference/fonts-rendering.md).
4. **Native charts/tables**, not images, so they stay editable. See
   [reference/pptx-techniques.md](reference/pptx-techniques.md).
5. **Render and look** after every change. Don't trust the code; trust the PNG.
6. **Keep content out of the engine.** Put all text/colors/figures in the config
   so re-skinning is a one-place change and nothing is hard-coded.

## Config shape (everything is data)
A config is a dict/YAML. See [examples/config.example.yaml](examples/config.example.yaml)
for a complete, commented, dummy poster.

- `page`: `size` (A0/A1/A2/A3), `orientation`, `margin`, `gutter` (mm).
- `theme`: `font`, `font_path`, `body_size`, and hex colors (`primary`,
  `primary_dark`, `primary_light`, `accent`, `ink`). Re-skin by editing colors.
- `header`: `title`, `authors`, `affiliation`, optional `logo`.
- `abstract` (optional, full-width): `label`, `paragraphs`.
- `left` / `right`: ordered lists of **blocks** (the two columns).
- `footer`: `left`, `right`.

### Block types (in `left`/`right`)
| type | keys | notes |
|---|---|---|
| `section` | `title` | colored section bar |
| `text` | `paragraphs[]`, `align` | body; default justified |
| `bullets` | `items[]` | bullet list |
| `figure` | `path`, `caption`, `number`, `height_mm` | real image, aspect-fit |
| `placeholder` | `label`, `caption`, `height_mm` | gray box when no image yet |
| `table` | `data[][]`, `col_widths[]`, `highlight_row`, `caption` | native table |
| `chart` | `categories[]`, `series[]`, `value_range`, `caption` | native bar chart; `series[].point_colors` for per-bar color |
| `equation` | `parts[[text,is_main]]`, `number` | one-line, auto-sized to box |
| `space` | `height_mm` | spacer |

## Programmatic use
```python
import sys; sys.path.insert(0, "scripts")
from poster import Poster
Poster(config_dict).save("out.pptx")
```

## Scripts (run these; don't paste their contents into context)
- `scripts/poster.py` — the engine (`Poster` class). Import it.
- `scripts/measure.py` — font-metric text measurement (used by the engine).
- `scripts/render_preview.py` — `python render_preview.py file.pptx [--crop x0,y0,x1,y1]`.
- `scripts/unlock_macos.sh` — `unlock_macos.sh file.pptx` (macOS read-only fix).

## Reference (load only when needed)
- [reference/layout-engine.md](reference/layout-engine.md) — flow + measuring + safety factor.
- [reference/pptx-techniques.md](reference/pptx-techniques.md) — size/units, native charts & tables, OOXML gradient/shadow, CJK fonts.
- [reference/fonts-rendering.md](reference/fonts-rendering.md) — autofit, finding app-bundled fonts, the render loop.
- [reference/troubleshooting.md](reference/troubleshooting.md) — overflow, font substitution, read-only file, clipboard image, upscaling.

## Notes
- All bundled content is dummy/placeholder. Replace it; never hard-code real
  text in the engine.
- macOS-specific helpers (quarantine, clipboard, app-bundled font paths) are
  marked as such; on other platforms use the cross-platform equivalents noted
  in the reference files.
