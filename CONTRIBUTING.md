# Contributing to academic-poster

Thanks for taking the time to contribute! Bug reports, new block types, more
cross-platform font paths, non-macOS equivalents for the helper scripts, and
docs improvements are all welcome.

## Ways to help
- **Report a bug** — open an issue with a minimal config that reproduces it and
  (ideally) a screenshot of the rendered PNG.
- **Request a feature** — describe the poster outcome you want, not just the API.
- **Send a PR** — see the workflow below.
- **Improve docs** — the `reference/` files are meant to teach the *why*; fixes
  and clarifications are valuable.

## Development setup
```bash
git clone https://github.com/YJlang/academic-poster.git
cd academic-poster
pip install -r requirements.txt      # python-pptx, Pillow, PyYAML
# LibreOffice is also needed for the render/preview step (soffice on PATH)
```

## The golden rule: render and look
This project's whole point is a layout that doesn't overflow. **Any change that
touches layout must be verified visually**, not just by the code running:

```bash
cd academic-poster
python examples/demo_poster.py examples/config.example.yaml out.pptx
python scripts/render_preview.py out.pptx        # -> out.png (+ out.pdf)
# then actually open out.png and check for overflow, clipping, balance
```

If you change the measuring logic, also test with a long paragraph and with CJK
text, since those are the cases the safety factor protects.

## PR workflow
1. Fork and create a branch off `main` (e.g. `fix/equation-overflow`).
2. Make your change. Keep it focused — one logical change per PR.
3. Verify with the render-and-look loop above. Attach a before/after PNG when
   the change is visual.
4. Open a PR describing **what** you changed and **why**.

## Project principles (please keep these)
- **Content stays out of the engine.** All text, colors, figures, tables, and
  charts live in the config. The bundled example config must stay
  placeholder-only — never hard-code real author/venue content in `poster.py`.
- **Flow, never fixed `y`.** New block renderers must measure their content and
  return the height they consumed. See
  [reference/layout-engine.md](academic-poster/reference/layout-engine.md).
- **Editable output.** Prefer native PowerPoint charts/tables over images so the
  result stays editable.
- **Mark platform-specific code as such.** macOS-only helpers (quarantine,
  clipboard, app-bundled fonts) should note the cross-platform equivalent.

## Code style
- Plain, readable Python; match the surrounding style in `scripts/poster.py`.
- No new hard dependencies without discussion — the install footprint
  (python-pptx, Pillow, PyYAML) is intentionally small.

## Reporting security issues
For anything sensitive, please avoid filing a public issue and contact the
maintainer directly via the email on their GitHub profile.

By contributing, you agree that your contributions are licensed under the
project's [MIT License](LICENSE).
