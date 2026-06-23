# python-pptx techniques for posters

## Contents
- Poster size (A0/A1/A2) and units
- Native charts (editable)
- Native tables
- Gradients, shadows, rounded corners (OOXML)
- East-Asian / CJK fonts
- Images and logos

## Poster size (A0/A1/A2) and units
A poster is a single oversized slide. Set the slide size explicitly; everything
else is positioned in millimeters with `pptx.util.Mm`.

```python
prs = Presentation()
prs.slide_width  = Mm(594)   # A1 portrait
prs.slide_height = Mm(841)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
```

Paper sizes (portrait, mm): A0 841×1189, A1 594×841, A2 420×594. Swap w/h for
landscape. **Font sizes are physical points** regardless of slide size: at A1,
~40pt title / ~20pt body / ~15pt caption print well.

## Native charts (editable)
Add a real chart (double-click-editable in PowerPoint), not an image:

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
cd = CategoryChartData(); cd.categories = [...]; cd.add_series("Acc (%)", (...))
chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, cd).chart
```

Per-bar color (e.g. to flag weak categories) via points:
```python
for i, p in enumerate(chart.series[0].points):
    p.format.fill.solid(); p.format.fill.fore_color.rgb = color_i
```
Data labels: `plot.has_data_labels=True; plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END`.
See `Poster._chart()`.

## Native tables
`slide.shapes.add_table(rows, cols, x, y, w, h).table`. Set
`columns[c].width`, per-cell `.fill`, `.vertical_anchor`, and the paragraph
alignment. Tip: align **every** cell `CENTER` (including the first column) unless
you have a reason not to — mixed alignment looks accidental. See `Poster._table()`.

## Gradients, shadows, rounded corners (OOXML)
python-pptx has no high-level API for gradients/soft shadows, so edit the shape
XML directly via `shape._element.spPr` and the `a:` namespace:

- Gradient: insert `a:gradFill > a:gsLst > a:gs(a:srgbClr)` + `a:lin` (angle).
- Soft shadow: insert `a:effectLst > a:outerShdw(a:srgbClr + a:alpha)`.
- Rounded rect: `add_shape(MSO_SHAPE.ROUNDED_RECTANGLE)` then `adjustments[0]`.

Copy `Poster._gradient()` / `Poster._shadow()` rather than re-deriving the XML.

## East-Asian / CJK fonts
Setting only the Latin typeface leaves CJK glyphs to a fallback (often ugly or
boxes). Set the East-Asian (`a:ea`) and complex-script (`a:cs`) typefaces on the
run too:

```python
rPr = run._r.get_or_add_rPr()
for tag in ("a:ea", "a:cs"):
    e = rPr.find(qn(tag)) or rPr.makeelement(qn(tag), {}); e.set("typeface", font_name)
```
`Poster._set_font()` does this for every run.

## Images and logos
Fit an image into a box keeping aspect ratio (`Poster._image_fit()`): compute
scale from the box vs image ratio, center it. For a logo over a colored header,
a transparent PNG can blend in — put it on a white badge or use a white/mono
version if it disappears.
