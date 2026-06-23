"""Render a .pptx poster to PNG (and PDF) with LibreOffice for visual review.

Why: previewing the rendered image is the single most reliable way to catch
overflow, clipping, and font-substitution before printing. Run this after every
edit and actually LOOK at the PNG.

Usage:
    python render_preview.py poster.pptx                 # -> poster.png (+ .pdf)
    python render_preview.py poster.pptx --out /tmp/prev # custom output dir
    python render_preview.py poster.pptx --crop 0,0.7,0.5,1.0   # crop a region

Requires LibreOffice (`soffice` on PATH, or /Applications/LibreOffice.app on macOS).
"""
import argparse
import os
import shutil
import subprocess
import sys


def find_soffice():
    for c in ("soffice", "libreoffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.isabs(c) and os.path.exists(c):
            return c
        if shutil.which(c):
            return c
    sys.exit("LibreOffice not found. Install it (macOS: brew install --cask libreoffice).")


def convert(soffice, pptx, outdir, fmt):
    flt = "png:impress_png_Export" if fmt == "png" else "pdf"
    subprocess.run([soffice, "--headless", "--convert-to", flt, "--outdir", outdir, pptx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + "." + fmt)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None, help="output dir (default: alongside pptx)")
    ap.add_argument("--no-pdf", action="store_true")
    ap.add_argument("--crop", default=None, help="x0,y0,x1,y1 as 0..1 fractions")
    args = ap.parse_args()

    soffice = find_soffice()
    outdir = args.out or os.path.dirname(os.path.abspath(args.pptx)) or "."
    os.makedirs(outdir, exist_ok=True)

    png = convert(soffice, args.pptx, outdir, "png")
    print("PNG:", png)
    if not args.no_pdf:
        print("PDF:", convert(soffice, args.pptx, outdir, "pdf"))

    if args.crop:
        from PIL import Image
        x0, y0, x1, y1 = (float(v) for v in args.crop.split(","))
        im = Image.open(png); W, H = im.size
        cp = os.path.join(outdir, "crop.png")
        im.crop((int(W * x0), int(H * y0), int(W * x1), int(H * y1))).save(cp)
        print("CROP:", cp)


if __name__ == "__main__":
    main()
