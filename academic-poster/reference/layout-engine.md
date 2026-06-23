# Layout engine: flow, not fixed coordinates

## Contents
- The core problem
- The flow model
- Measuring text (the key trick)
- The safety factor
- Filling the page without overflow
- One-line auto-fit (equations, banners)

## The core problem
A poster is one big slide with a lot of text. The naive approach is to place
each element at a hard-coded `y` (`y += 158`). This **breaks the moment body
text is longer than you guessed**: the next section bar lands on top of the
text. Most "the title box covers the paragraph" bugs come from fixed `y`.

## The flow model
Lay elements out top-to-bottom per column. Each element returns the height it
consumed; the next element starts right below it:

```
y = column_top
for block in column:
    y += render(block, x, y, col_width)   # render returns mm consumed
```

A section bar is just another block. When the paragraph above it grows, the bar
is pushed down automatically — it can never be overlapped. `scripts/poster.py`
implements exactly this in `Poster.build()` / `Poster._block()`.

## Measuring text (the key trick)
To flow, the engine must know a paragraph's rendered height *before* drawing it.
`scripts/measure.py` computes it from font metrics with Pillow:

1. Load the font at 4× point size for resolution: `ImageFont.truetype(path, pt*4)`.
2. Greedy-wrap character by character against the column width
   (`getlength` per char) → line count. Works for CJK and Latin.
3. `height = lines × size × line_spacing × 0.352778 mm`.

This is why `theme.font_path` matters: **measure with the same font you draw
with**, or the predicted height won't match the rendered height.

## The safety factor
Real renderers (PowerPoint, LibreOffice) lay lines out slightly taller than the
raw font metric. So `text_height_mm()` measures with `line_spacing≈1.33` while
the text is *drawn* at `1.27`. Over-estimating height pushes the next block down
a hair — trading a sliver of whitespace for a guarantee of **no overlap**.

If you ever see clipping in the real app but not in the LibreOffice preview,
raise the measuring `line_spacing` (e.g. 1.27 → 1.33). If you see too much gap,
lower it. Never chase a pixel-perfect match: keep a small safety margin.

## Filling the page without overflow
Because every block's height is known, you can balance the two columns by
adjusting content and figure sizes, then re-rendering:

- Both columns should end a little above the footer line.
- To fill a short column: enlarge a figure/chart, or add a sentence — don't just
  inflate gaps (that reads as empty).
- To fix an overflowing column: shrink a figure or trim a sentence.
- Iterate: build → `render_preview.py` → look → adjust. 2–3 passes converge.

## One-line auto-fit (equations, banners)
For an equation that should fill its box on a single line, don't hard-code the
font size. `measure.fit_one_line_size(parts, avail_mm, font_path)` returns the
largest size that fits, accounting for subscripts (rendered at `size×0.66`).
`Poster._equation()` uses it and sets `word_wrap=False` so it stays on one line.
