# CESI-inspired presentation generator

Offline generator for editable 16:9 PowerPoint and independently rendered PDF slides from a UTF-8 JSON deck. Its black `#000000`, yellow `#FBE216` and white `#FFFFFF` palette and high-contrast headings are inspired by observations of [the CESI homepage](https://www.cesi.fr/) and its [website stylesheet](https://www.cesi.fr/wp-content/themes/cesi-2022/css/base.min.css?ver=1.4.25). This is **not an official CESI brand guide**; no official logo or endorsement is included. The generator uses Arial and Helvetica, not CESI's web fonts. Obtain authorized materials for official-branded deliverables.

## Setup

Python 3.10+:

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/generate.py examples/internal-update.json --pptx output/deck.pptx --pdf output/deck.pdf
python -m unittest discover -s tests -v
```

On Windows use `.venv\Scripts\activate`. Outputs can be requested separately with `--pptx` or `--pdf`. The script creates parent directories and overwrites existing output files. Input and output paths must differ. It emits JSON on success and exits nonzero on error.

## Input

`slides` is a non-empty array. Each slide has `type` (`title`, `section`, or `content`; default `content`), required single-line `title` (max 80 characters), optional single-line `subtitle` (max 110). Content slides may include up to five nonempty, single-line `bullets` (max 120 characters each); bullets are not allowed on other types. Optional deck-level `footer` has at most 70 characters. Use concise text; character limits alone do not guarantee a fit. The PDF generator rejects text it detects exceeding boxes; PowerPoint can wrap differently and must be visually inspected. Unknown properties are ignored. Images, speaker notes, charts and official templates are not supported.

PDF is **not a conversion of the PPTX**. ReportLab draws its own slide pages using Helvetica, while PPTX uses Arial; rendering and line breaks may differ. Standard PDF fonts are not guaranteed to cover arbitrary Unicode; check glyphs or use an approved font workflow. PPTX-to-PDF fidelity requires a real converter such as LibreOffice, installed separately; for example, convert to a different directory using `soffice --headless --convert-to pdf --outdir converted output/deck.pptx`. Inspect the result. This generator does not invoke LibreOffice.

Tests use unittest and check CLI errors, deck counts and extracted text. Visual verification is separate. Any sample content is fictitious and must not be treated as CESI facts.
