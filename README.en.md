# office-docs-skill

**Read, write, convert and generate Word / Excel / PowerPoint / PDF documents — as an Agent Skill
or as a plain CLI.** It wraps LibreOffice headless plus Python libraries, so it works on machines
without Microsoft Office (and without WPS).

[中文](README.md) ｜ MIT

---

## Requirements

| Kind | What you need |
| --- | --- |
| System commands | `soffice` (LibreOffice), `pdftoppm` / `pdftotext` (poppler-utils), `tesseract` with `chi_sim` for Chinese OCR |
| Python packages | `python-docx`, `openpyxl`, `python-pptx`, `xlrd`, `odfpy`, `pymupdf`, `pypdf`, `pdf2docx`, `cryptography` (PDF AES only) |
| Optional | `pandoc` (Markdown ↔ docx and friends) |
| Chinese layout (optional) | CJK fonts plus fontconfig aliases; `scripts/install_fonts.sh` installs and checks them |
| Common Latin fonts (optional) | The same alias file maps Arial / Helvetica / Times New Roman / Georgia / Calibri / Cambria / Verdana / Tahoma / Courier New onto metric-compatible local families |

Run `python3 scripts/office_cli.py doctor` first — it lists every item above as `ok` or missing.
Nothing else is assumed: there is no install step, no service and no configuration file.

## Quick start

```bash
python3 scripts/office_cli.py doctor          # environment check; {"ok": true} means ready
python3 scripts/selftest.py                   # end-to-end self-test; prints N/N and exits 0 only when all pass
bash    scripts/install_fonts.sh --check      # CJK font/alias check (read-only)

python3 scripts/office_cli.py info  a.xlsx    # real type and structure, not the extension
python3 scripts/office_cli.py text  a.pdf     # extract text
python3 scripts/office_cli.py convert a.docx --to pdf
python3 scripts/office_cli.py pptx-new --spec s.json -o out.pptx
python3 scripts/office_cli.py thumb out.pptx -o thumbs/
python3 scripts/office_cli.py pdf-ocr scan.pdf -l chi_sim+eng
```

**Contract:** a command succeeded when the **exit code is 0 and stdout is `{"ok": true, ...}`**.
Errors come back in the same envelope with a message and a non-zero exit code.

## The four gates before delivery

Converting successfully is not the same as being correct. These four checks catch the failures that
do not show up when you open the file locally:

```bash
python3 scripts/office_cli.py ooxml-check out.docx    # ① package: content types / rels / r:id / XML well-formedness
python3 scripts/office_cli.py form-verify spec.json   # ② geometry: render, measure table lines against spec.tol_cm
python3 scripts/office_cli.py xlsx-recalc out.xlsx    # ③ formulas: recalculate with LibreOffice, scan #REF!/#NAME?
python3 scripts/office_cli.py pdf-boxcheck spec.json  # ④ boxes: overflow / overlap / off-page before filling a PDF
python3 scripts/office_cli.py pdf-fill a.pdf spec.json -o out.pdf --preview p.png
```

Rationale, failure modes and spec formats: [`references/quality-gates.md`](references/quality-gates.md).

## What it does

| Area | Highlights |
| --- | --- |
| Read | Text and tables from docx/xlsx/pptx/pdf/odt/ods and legacy doc/xls/ppt; structure info (pages, sheets, slide count); real format detection by content |
| Convert | Any-to-any through LibreOffice headless, with an isolated profile per call (safe to run concurrently) |
| Generate | `docx-new`, `xlsx-new`, `pptx-new` from a JSON spec; Chinese official-document (`gongwen`), contract, report and briefing profiles |
| Edit | Find-and-replace across body, tables, text boxes, headers and footers — run-aware, so matches split across runs still hit |
| Excel | Recalculate formulas to real values; add charts; export UTF-8 CSV |
| PDF | Merge, split, extract images, rotate, watermark, AES-256 encrypt/decrypt, OCR, PDF → editable Word, render to PNG |
| Forms | List and fill AcroForm fields; for static PDFs, extract the structure, self-check box positions, then fill by box with a red-box preview |
| Reverse-engineer forms | Learn a profile from a scan/photo (`form-learn`), render a clean blank template from a spec (`form-render`), verify geometry (`form-verify`), or cut the handwriting layer out (`form-extract` / `form-blank`) |
| Review | Inventory Word revisions and comments (`docx-review`, read-only) |

Every subcommand is listed in [`SKILL.md`](SKILL.md); formats, filters and spec JSON are documented
in [`references/formats.md`](references/formats.md).

## As an Agent Skill

`SKILL.md` is the skill body: its YAML front matter (`name`, `description`) decides when an agent
loads it. Put the directory in your agent's skills folder and enable it there — for example
`~/.claude/skills/office-docs-skill/` in Claude Code or `~/.dsh/skills/office-docs-skill/` in DSH.

`examples/gongwen-demo.json` is a runnable Chinese official-document spec (title, recipient, four
heading levels, attachment, signature, note). `examples/gongwen-demo.docx`, `gongwen-demo-portable.docx`,
`gongwen-demo-portable.pdf` and `gongwen-demo-portable-1.png` / `-2.png` are its outputs, kept for
visual layout comparison.

`templates/` holds form specs in JSON: a blank template plus a matching scan profile, used by
`form-render` / `form-verify`.

## Limits

- **Microsoft Works/WPS proprietary formats** (`.wps`/`.et`/`.dps`) are not supported. LibreOffice's
  `MS_Works` filter is Microsoft Works, unrelated to Kingsoft WPS.
- **PDF → Word is reconstruction, not conversion.** Text usually survives; tables are often flattened
  into text lines and multi-column layouts can interleave. Check important documents by hand, or ask
  for the original docx.
- **CJK glyphs use substitutes** unless the Microsoft fonts are installed: SimSun → AR PL SungtiL GB,
  FangSong → cwTeX FangSong, KaiTi → LXGW WenKai, SimHei → Noto Sans CJK. The requirement is only
  that each name lands in the **right class** (SimSun → a serif Song, SimHei → a sans black); glyphs
  and metrics differ, so pagination is not pixel-identical to Word on Windows. Write the normal font
  names (`宋体`, `仿宋_GB2312`, `Arial`, `Times New Roman`) in the document — both ends resolve them.
  See `font_aliases` in `doctor` output for the per-name verdict.
- **Legacy binary formats** (doc/xls/ppt) are readable only through LibreOffice; convert to OOXML
  before processing them programmatically.
- **PDF encryption needs `cryptography`.** Form filling covers AcroForm fields only; XFA (`/XFA`)
  forms are not supported.
- **Automated OCR templates need a human pass.** On faint or phone-photographed forms, `form-learn`
  misreads characters, so the profile must be reviewed. Where handwriting overlaps printed text,
  rebuild from a spec instead of trying to separate the scan.
- **Layout changes are verified by measurement, not by eye** — run `form-verify`.

## Layout

```
SKILL.md                  # Skill entry: rules, decision table, command reference
README.md / README.en.md  # Chinese / English documentation
scripts/
  office_cli.py           # Single entry point (JSON envelope with an `ok` field)
  convert.py              # LibreOffice headless wrapper + doctor (concurrency, timeouts, artifact checks)
  readers.py              # Text / structure / tables, typed by content
  writers.py              # Generate and modify docx/xlsx/pptx
  cn_profiles.py          # Chinese official-document / contract / report / briefing layouts
  pdf_ops.py              # PDF merge, split, images, render, OCR, to Word
  extras.py               # PDF rotate/watermark/AES/forms, Excel charts, Word revision inventory
  docx_runs.py            # Run merging + replace by rendered text
  xlsx_recalc.py          # LibreOffice macro recalculation + error scan
  pdf_forms.py            # Static PDF: structure, box check, fill by box, preview
  ooxml_check.py          # OOXML package gate
  form_template.py        # Form reverse engineering: grid/OCR profile, spec rendering, geometry verify
  install_fonts.sh        # Install CJK fonts + Windows font-name aliases (--check for read-only)
  selftest.py             # End-to-end self-test, real assertions
assets/                   # fontconfig alias file installed by install_fonts.sh
templates/                # Form specs (JSON) and matching scan profiles
references/
  formats.md              # Formats, filter names, font substitution, spec JSON
  quality-gates.md        # The four delivery gates in depth
  cn-doc-formats.md       # Chinese document typography quick reference
  form-reverse.md         # Form reverse engineering: profile fields, fill regions, font sizing
examples/                 # Chinese official-document demo spec and rendered outputs
```

## License

MIT

---

> **Note**: this project was accumulated by an agent through day-to-day use — the commands, the
> pitfalls and the checks all come from tasks actually run, not from a spec written up front.
