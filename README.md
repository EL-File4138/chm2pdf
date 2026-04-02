# chm2pdf

A simple CHM to PDF conversion script.

It's surprising that after all these years a simple search of `chm2pdf` still yields the ancestor project on [Google Code Archive](https://code.google.com/archive/p/chm2pdf/). Some other similar projects on GitHub are almost all LLM slops, broken to an extent and contain questionable design decisions. So off we go to dedicate some hours of effort.

This rewrite removes the original runtime dependency on `pychm` and `htmldoc`.
It now uses:

- `pylibmspack` to read and extract CHM archives
- `beautifulsoup4` to rebuild a single printable HTML document
- headless Chromium to render PDF output

## Requirements

- Python 3.10+
- `uv` (Use it for your own good, or use another Python package manager with `pyproject.toml` support)
- `chromium-browser` or another Chromium-compatible browser on `PATH` (Better than using `playwright`, since if you are anti-chromium, then you won't use that route anyway, and if you do not, it's much better than installing another Chromium instance)

## Setup

Download the project's source, then 

```bash
uv sync
```

## Usage
```
usage: chm2pdf [-h] [--extract-only] [--dontextract] [--keep-temp] [--work-dir WORK_DIR] [--title TITLE]
               [--titlefile TITLEFILE] [--paper-size PAPER_SIZE] [--landscape] [--inline-toc] [--pdf-header-footer] [-v]
               [--version]
               input [output]

Convert CHM files into PDF.

positional arguments:
  input                 Input CHM file
  output                Output PDF file

options:
  -h, --help            show this help message and exit
  --extract-only        Extract the CHM and stop
  --dontextract         Reuse previously extracted files
  --keep-temp           Keep temporary build files
  --work-dir WORK_DIR   Override the temporary work root
  --title TITLE         Override the document title
  --titlefile TITLEFILE
                        Promote a specific extracted HTML file to the front
  --paper-size PAPER_SIZE
                        CSS page size, e.g. A4 or Letter
  --landscape           Render in landscape orientation
  --inline-toc          Insert a generated table-of-contents page into the PDF
  --pdf-header-footer   Keep Chromium's printed header/footer block (date, title, URL, page numbers)
  -v, --verbose         Print progress information
  --version             show program's version number and exit
```


## Examples

```bash
uv run chm2pdf input.chm
uv run chm2pdf input.chm output.pdf
uv run chm2pdf --extract-only input.chm
uv run chm2pdf --title "Manual" --paper-size Letter input.chm output.pdf
```

## Note
- Yes, it doesn't generate a sophisticatedly typeset document. I'm solving my own problem, and my limited samples (2) are good enough for me. If you have samples that pose a typographic feature that is important enough not to ditch, open an issue or (if you code) a PR.
- CHM, as an obsolete artifact that exists predominantly in the pre-Microslop proprietary era, contains many quirks and malformed structures. Parsing can be done only to a limited extent. If you have samples that require specific parsing to be readable and you think it would be beneficial for others, do that thingy too.

## License
Unlicensed.
