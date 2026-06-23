# Fonts and rendering

## Contents
- Turn OFF autofit (the "font changed" bug)
- Choosing a font that actually renders
- Finding a font that "isn't installed" (app-bundled fonts)
- The render-and-look verification loop
- Exporting PDF for print

## Turn OFF autofit (the "font changed" bug)
By default a text box may shrink its text to fit when content overflows, so some
boxes end up with smaller letters than others — it looks like the font changed.
Disable it on every text frame:

```python
from pptx.enum.text import MSO_AUTO_SIZE
tf.auto_size = MSO_AUTO_SIZE.NONE
```

With autofit off, text stays at its set size; prevent overflow with the flow
engine + measuring instead (see layout-engine.md). `Poster._tb()` sets this.

## Choosing a font that actually renders
The font NAME written into the pptx must exist on whoever opens it, or the app
substitutes a different font — which also changes text width and can cause
overflow. Prefer fonts present on the target machine/app. For measurement you
need the matching `.ttf`/`.ttc` path so predicted height = rendered height.

## Finding a font that "isn't installed" (app-bundled fonts)
A font missing from the OS may still ship **inside an app bundle**. Example
(macOS): Microsoft PowerPoint bundles many fonts under
`/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/`. If your
target is PowerPoint, you can measure with the bundled `.ttf` and the result
matches PowerPoint exactly. To also preview correctly in LibreOffice, copy that
`.ttf` into `~/Library/Fonts/`. Search before assuming a font is unavailable:

```bash
find / -iname "*<fontname>*.ttf" 2>/dev/null         # be patient / scope it
mdfind -name "<fontname>"                              # macOS Spotlight
```

`examples/demo_poster.py:find_font()` shows a cross-platform fallback list
(Arial / Helvetica / DejaVu / Liberation).

## The render-and-look verification loop
There is no substitute for looking at the rendered image. After each change:

```bash
python scripts/render_preview.py poster.pptx          # -> poster.png (+pdf)
python scripts/render_preview.py poster.pptx --crop 0.5,0.4,1.0,0.62   # zoom a region
```

Then actually open the PNG and check: overflow, clipping at section bars,
font substitution, column balance. The preview uses LibreOffice headless.

**Caveat:** the LibreOffice preview is not byte-identical to PowerPoint
(metrics differ slightly). It's great for layout/overflow, but confirm the final
file in the target app. The measuring safety factor (layout-engine.md) absorbs
the difference.

## Exporting PDF for print
`render_preview.py` also emits a PDF via LibreOffice. For best fidelity, the
safest print export is from the target app itself (File ▸ Export ▸ PDF), since
that uses the app's own fonts and renderer.
