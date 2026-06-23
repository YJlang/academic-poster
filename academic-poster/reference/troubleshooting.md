# Troubleshooting

## Contents
- Text overflows / a section bar covers a paragraph
- Letters are different sizes across boxes
- Font looks wrong / substituted
- "Can't edit the .pptx" — opens read-only (macOS)
- Preview looks fine but PowerPoint overflows
- Equation too small in a big box
- Pasted image has no file (clipboard, macOS)
- Low-resolution logo/image

## Text overflows / a section bar covers a paragraph
Cause: fixed `y` coordinates. Fix: use the flow engine and measure text height
(see layout-engine.md). Every block must return its height and the next block
starts below it. If it still overflows in the real app, raise the measuring
`line_spacing` safety factor.

## Letters are different sizes across boxes
Cause: text-box autofit shrinking overflowing text. Fix:
`tf.auto_size = MSO_AUTO_SIZE.NONE` on every text frame (fonts-rendering.md),
and prevent overflow by flowing/measuring rather than cramming.

## Font looks wrong / substituted
The font NAME isn't available where the file is opened, so the app substitutes
one (also changing widths → possible overflow). Use a font present on the
target app, or find an app-bundled one (fonts-rendering.md). Always set the
`a:ea`/`a:cs` typefaces too for CJK (pptx-techniques.md).

## "Can't edit the .pptx" — opens read-only (macOS)
PowerPoint opens files with a `com.apple.quarantine` attribute in read-only
"Protected View". Strip it:

```bash
xattr -c poster.pptx        # or: scripts/unlock_macos.sh poster.pptx
```

Then close and reopen. (Or click "Enable Editing" in the yellow bar.) A freshly
generated file usually has no quarantine; it often appears after the app has
touched it.

## Preview looks fine but PowerPoint overflows
LibreOffice (preview) and PowerPoint have slightly different line metrics. Keep
the measuring safety factor (`line_spacing≈1.33`) so the layout has headroom,
and confirm the final file in the target app.

## Equation too small in a big box
Don't hard-code the size. Use `measure.fit_one_line_size()` to grow the equation
to the box width on one line, set `word_wrap=False`, and size the box height to
the font (see `Poster._equation()`).

## Pasted image has no file (clipboard, macOS)
If the user pasted an image (e.g. a logo) and there's no file, extract the
clipboard PNG:

```bash
osascript -e 'set f to (open for access POSIX file "/tmp/logo.png" with write permission)' \
          -e 'write (the clipboard as «class PNGf») to f' -e 'close access f'
```

## Low-resolution logo/image
Upscale with Pillow before placing (keeps it crisp at poster scale):
```python
from PIL import Image, ImageFilter
im = Image.open("logo.png").convert("RGBA")
im = im.resize((im.width*4, im.height*4), Image.LANCZOS).filter(ImageFilter.UnsharpMask(2, 80, 2))
im.save("logo_hd.png")
```
