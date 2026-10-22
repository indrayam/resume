# Resume site

This repository contains the static resume site published at [resume.indrayam.com](https://resume.indrayam.com/).

## Project map

- `references/resume.md` — factual resume content and source referenced by the HTML page.
- `index.html` and `style.css` — primary resume page and its styling.
- `architect.html` — the supporting software-architect page linked from the resume.
- `build_resume_docx.py` — builds the editable DOCX at `output/docx/anand-sharma-resume.docx`. Its resume text is currently maintained separately in the script; it does not read `references/resume.md`.
- `output/pdf/` and `output/docx/` — downloadable resume documents served by the site.
- `images/` — site and resume images.

The Markdown, HTML, DOCX builder, and downloadable files are not all generated from one source. When resume facts change, review the affected representations and keep them consistent. The PDF generation workflow is not documented in this repository; inspect it before attempting to rebuild the PDF.

## Working on the repository

Read [`AGENTS.md`](AGENTS.md) before making changes. It covers source-of-truth handling, Git synchronization, and preservation and verification expectations. This repository is part of the wider `indrayam.com` site; the parent folder's `AGENTS.md` applies when work crosses repository boundaries.

No general site build or test command is documented here. Avoid assuming one exists; identify the relevant workflow before regenerating deliverables.
