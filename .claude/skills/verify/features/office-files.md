# Office files

`.pptx`, `.docx` and `.xlsx` files are published as they are and drawn on the server from the file's own masters, styles and theme. Presentations become decks, documents become pages of selectable text, and workbooks get a tab per sheet. Pictures are served from inside the file under `/~office/…`. The original bytes are always one click away.

## Sub-features

- `pptx`: `.slide.active` sections with real text. Pictures come from `<content>/a/<id>/v/<n>/~office/<file>/ppt/media/…`. Speaker notes go in `.slide-notes`. Arrow keys turn the slides.
- `docx`: headings and paragraphs as DOM text.
- `xlsx`: `a.xl-tab` per sheet, with `href="#sheet-N"`. Clicking a tab shows that sheet.
- `download`: `#export-pptx` in the viewer, labelled PPTX/DOCX/XLSX, serves the original file.
- `read/export`: `artifact read <id> --out FILE` returns the original. `artifact export <id>` returns a published `.pptx` unchanged.
- `no editor`: without ONLYOFFICE configured there is no 編輯 button (`#edit` hidden).

## How to get to it (user POV)

- Agent: `artifact publish deck.pptx` (or `.docx`/`.xlsx`) and shares the URL.
- Reader: opens the viewer, flips slides or sheet tabs, and clicks the PPTX/DOCX/XLSX button to get the file.

## Driving it with the artifact CLI and Playwright Chromium

Preconditions: instance ready. Optionally `VERIFY_OFFICE_FILE=/path/to/real.pptx` to add a real file to the run.

```bash
node .claude/skills/verify/scripts/make-office.mjs .verify/run/work/office   # just the files
node .claude/skills/verify/scripts/journeys/office-files.mjs                 # full journey
VERIFY_OFFICE_FILE=~/Downloads/real.pptx node .claude/skills/verify/scripts/journeys/office-files.mjs
```

- `make-office.mjs` writes three files. `verify.pptx` (pptxgenjs) has 2 slides; slide 2 has a 64×64 PNG, and both slides have notes. `verify.docx` has a heading and a paragraph. `verify.xlsx` has sheets "Totals" and "Details". The last two are zipped with `test/helpers/zip.ts`.
- Each file is published with `--new`.
- pptx: `.slide.active` shows "Verify deck slide one". After ArrowRight it shows "Slide two with a picture". Its `img` has `naturalWidth` 64 and `currentSrc` matching `/~office/verify.pptx/ppt/media/`. `.slide-notes` holds the notes and the frame origin is `'null'`.
- docx: the frame body text contains the heading and the paragraph.
- xlsx: there are 2 `a.xl-tab` tabs. Clicking `a.xl-tab[href="#sheet-2"]` shows "verify-second-sheet" and hides the first sheet's text.
- For each file:
  - `#export-pptx` download has the same sha256 as the source.
  - `artifact read <id> --out` has the same sha256.
  - `#edit` is hidden.
- pptx only: `artifact export <id> --out` has the same sha256.
- With `VERIFY_OFFICE_FILE`: it publishes that file and requires no `.pp-missing` boxes. Then it runs the same download and read checks.

Evidence: `01-pptx-slide1.png` … `05-xlsx-sheet2.png` (`06-real-file.png` with a real file), `cli.log`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to production and reads the real token.
- The generated files are only a few shapes. They prove the publish → render → download pipeline, not that any uploaded file renders. Feed real files through `VERIFY_OFFICE_FILE` for that, and note that `.pp-missing` only catches whole-slide failures, so look at the screenshot too.
- ONLYOFFICE online editing (the 編輯 button, signed callbacks) needs an external Document Server and is not covered. Launch leaves `ONLYOFFICE_*` unset on purpose.
- `#sheet-2` is a hash the tab links to, not an element id. Click `a.xl-tab[href="#sheet-2"]` inside the frame.
- The download button's id is `#export-pptx` for every Office type (and for Markdown decks, where it exports). Its label tells which.
- A picture may still be loading when the slide becomes active. Wait for `load` before reading `naturalWidth`.
