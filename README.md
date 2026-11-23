# Resume site

This repository contains the static resume site published at [resume.indrayam.com](https://resume.indrayam.com/). It uses the shared `indrayam.com` chrome: the site nav, the 10px orange top bar, a centered Source Sans Pro masthead, and the shared copyright footer.

## Project map

- `references/resume.md` — factual resume content and the source the HTML page names (`<meta name="content-source">`).
- `index.html` and `style.css` — the resume page and its styling. It links the downloadable PDF (opens inline) and DOCX in `output/`.
- `architect.html` — the supporting software-architect page linked from the resume.
- `404.html` — custom not-found page with the shared nav and footer.
- `output/docx/anand-sharma-resume.docx` and `output/pdf/anand-sharma-resume.pdf` — the downloadable resume documents served by the site. See "DOCX and PDF" below.
- `references/` — not deployed. `resume-original.docx` and `.pdf` are the original Word resume; `headshot-version/` is the current resume with the original studio headshot instead of the website portrait (its README explains how to switch back); `resume-photo.jpg` is the website portrait used in the DOCX header.
- `build_resume_docx.py` — the older DOCX generator. See below before running it.
- `images/` — site images: company logos, the share image `og-resume.png`, the 404 illustration, and the architect-elevator image.
- `favicon.ico`, `apple-touch-icon.png` — the icon shared by every indrayam.com site.
- `robots.txt`, `sitemap.xml` — crawler files for this host.
- `vercel.json` — the shared security headers, plus `Content-Type: application/pdf` and `Content-Disposition: inline` on the PDF.
- `.vercelignore` — keeps `AGENTS.md`, `README.md`, `build_resume_docx.py`, `references/`, `tmp/`, `.env*`, and `.claude/` out of deployments.
- `tmp/` — ignored scratch space (see `AGENTS.md`).

## DOCX and PDF

The HTML page, the DOCX, and the PDF are not generated from one source. When resume facts change in `references/resume.md`, update each representation and keep them consistent.

- **DOCX (current workflow).** Since November 2026 the DOCX is edited in place from the Word file, so its fonts (Montserrat), layout, and three-page structure come from the original document. The editing scripts live in the ignored `tmp/docx-update/` folder and are not version-controlled: `edit.py` applies text edits to a copy of the DOCX by cloning existing paragraphs and runs, `italic.py` italicizes a word inside a run, and `move_overview.py` moves Career Overview to the top of page 2. Each takes `in.docx out.docx`; always write to a copy and check the rendered pages before replacing `output/docx/`.
- **PDF.** Exported from the DOCX by Microsoft Word, so it matches the DOCX exactly: `tmp/docx-update/export.sh in.docx out.pdf` drives Word through AppleScript.
- **`build_resume_docx.py` (older generator).** Builds a DOCX from text hard-coded in the script, which was last updated in September 2026. Its output does not match the published DOCX, and running it overwrites `output/docx/anand-sharma-resume.docx`. Use it only after bringing its content up to date, or restore the published file with `git checkout -- output/docx/` afterwards. It reads the header portrait from `references/resume-photo.jpg` and writes the circular crop to `tmp/`. Run it with `uv run --no-project --with python-docx --with pillow python build_resume_docx.py`.

## Working on the repository

Read [`AGENTS.md`](AGENTS.md) before making changes. It covers source-of-truth handling, Git synchronization, and preservation and verification expectations. This repository is part of the wider `indrayam.com` site; the parent folder's `AGENTS.md` applies when work crosses repository boundaries.

The site has no build step. Preview it with any static server (the workspace's `.claude/launch.json` has a `resume` entry on port 8083). Vercel (Git-connected, project `resume`) deploys on every push to `main`.
