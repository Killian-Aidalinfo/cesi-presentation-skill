---
name: cesi-presentations
description: Generate indicative CESI-inspired internal presentations in PPTX and PDF.
---

# CESI-inspired presentations

Use this skill to create editable 16:9 PPTX slides and/or independently drawn PDF slides from a small JSON specification. The styling is **indicative, provisional and not official CESI branding**. No CESI logo, trademark asset, institutional identity or endorsement is supplied. When official brand compliance is required, obtain an authorized current template and visual guidelines instead.

## Workflow

1. Collect the audience, objective, language, slide content and any approved brand assets. Do not invent facts or claim approval. Use brief slide copy.
2. Copy `examples/internal-update.json` or `examples/designs.json` as a starting point and edit it in UTF-8. Available slide types: `title`, `section`, `content`, `quote` and `cards`. Each requires `title`; `subtitle` is optional. `content` supports up to five `bullets`; `cards` requires one to three. For `quote`, use `title` for the message and `subtitle` for attribution. `footer` is a deck-level string. Constraints and limitations are in README.md.
3. Install Python dependencies in a virtual environment as shown in README.md; then run:

```bash
python scripts/generate.py examples/internal-update.json --pptx output/deck.pptx --pdf output/deck.pdf
python -m unittest discover -s tests -v
```

4. Read back slide and page counts/text, and visually inspect every rendered page before delivery. If rendering is unavailable, state that visual verification was not performed. Review line breaks, overflow and glyphs, especially accented/non-Latin characters.
5. Explain that the PDF is **drawn independently** from the same input, not exported from PPTX; line wrapping, fonts and layout can differ. For actual PPTX-to-PDF conversion, use LibreOffice only if installed: `soffice --headless --convert-to pdf --outdir output output/deck.pptx`, then review the converted PDF. Avoid overwriting the independent PDF inadvertently.

Do not claim that the provisional colors, typography, or visual motif are official. No network access is needed for generation.
