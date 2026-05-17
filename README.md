# tex_to_docx

A small LaTeX-to-DOCX pipeline built around Pandoc, with a post-processing pass
that makes the generated Word document look closer to the LaTeX PDF.

Pandoc does the heavy conversion. The helper scripts then patch the DOCX OOXML
to improve Word styles, fonts, margins, line numbers, page breaks, figure
sizing, and PDF figure handling.

## Requirements

- `pandoc`
- `python3`
- Poppler tools: `pdftocairo` and `pdfinfo`
- Optional, but useful: the same fonts you use in the PDF, such as Latin Modern
  Roman

On macOS, the command-line dependencies can be installed with Homebrew:

```bash
brew install pandoc poppler
```

## Usage

```bash
./tex_to_word.sh INPUT.tex [OUTPUT.docx]
```

Examples:

```bash
./tex_to_word.sh paper.tex
DOCX_FONT="Times New Roman" ./tex_to_word.sh paper.tex paper.docx
DOCX_BIBLIOGRAPHY=refs.bib ./tex_to_word.sh paper.tex -- --csl nature.csl
```

By default, the script looks for figure assets in these directories beside the
input `.tex` file:

- `images-pdf-cropped`
- `images-pdf`
- `images`

PDF figures are rendered into temporary PNG previews under `build/docx-figures`
before Pandoc writes the DOCX.

## Configuration

Set these environment variables when your project layout differs:

```bash
DOCX_FONT="Latin Modern Roman"
DOCX_BUILD_DIR="build"
DOCX_REFERENCE_DOC="build/latex-like-reference.docx"
DOCX_BIBLIOGRAPHY="refs.bib,more_refs.bib"
DOCX_FIGURE_DIRS="figures:images"
DOCX_FIGURE_DPI=300
```

Extra Pandoc arguments can be passed after `--`:

```bash
./tex_to_word.sh paper.tex paper.docx -- --csl journal.csl --metadata lang=en-US
```

## Files

- `tex_to_word.sh`: main conversion command
- `tools/polish_pandoc_docx.py`: post-processes styles, page layout, figures,
  and page breaks inside the generated DOCX
- `tools/docx_pdf_images_to_png.lua`: points DOCX images at PNG previews instead
  of PDFs
- `tools/move_refs_before_backmatter.lua`: moves Pandoc citeproc references
  before common manuscript backmatter headings
- `build/latex-like-reference.docx`: reference DOCX template used by Pandoc

## Notes

This was originally tuned for a scientific manuscript, so it is intentionally
plain: letter paper, 1 inch margins, justified body text, centered figures, and
LaTeX-like heading sizes. If your journal or thesis style differs, adjust
`build/latex-like-reference.docx` or `tools/polish_pandoc_docx.py`.

## Reddit blurb

I made a small LaTeX-to-DOCX pipeline that uses Pandoc for the conversion, then
post-processes the generated DOCX so it matches the native LaTeX PDF much more
closely. It was the most robust option I found. You can ask Codex to adjust the
reference DOCX or the post-processing script to match your own style.
