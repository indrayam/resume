# Resume site

This repository contains the static resume site published at [resume.indrayam.com](https://resume.indrayam.com/). It uses the shared `indrayam.com` chrome: the site nav, the 10px orange top bar, a centered Source Sans Pro masthead, and the shared copyright footer.

## Project map

- `references/resume.md` — factual resume content and the source the HTML page names (`<meta name="content-source">`).
- `index.html` and `style.css` — the resume page and its styling. It links the downloadable PDF (opens inline) in `output/`.
- `architect.html` — the supporting software-architect page linked from the resume.
- `404.html` — custom not-found page with the shared nav and footer.
- `output/docx/anand-sharma-resume.docx` and `output/pdf/anand-sharma-resume.pdf` — the resume documents. The page links only the PDF; the DOCX is its Word source and stays reachable by direct URL. See "DOCX and PDF" below.
- `output/docx/anand-sharma-resume-2page.docx`, `output/pdf/anand-sharma-resume-2page.pdf`, and the `-fed` pair — two 2-page resumes styled after the main DOCX (Montserrat, cyan accent), single column with no photo. Neither pair is linked from the site. The federal pair (VA and IRS applications) adds hours per week, supervisor, and citizenship and is not linked from the site. `tools/federal-resume/build.py` generates both DOCX files (`--short` for the 2-page version) and `tools/docx-update/export.sh` exports each PDF. Their content lives in the script, not in `references/resume.md`.
- `references/` — not deployed. `resume-original.docx` and `.pdf` are the original Word resume; `headshot-version/` is the current resume with the original studio headshot instead of the website portrait (its README explains how to switch back); `resume-photo.jpg` is the website portrait used in the DOCX header.
- `build_resume_docx.py` — the older DOCX generator. See below before running it.
- `tools/docx-update/` — the scripts behind the current DOCX and PDF (not deployed). See below.
- `images/` — site images: company logos, the dash background tile `blue-dash.gif`, the share image `og-resume.png`, the 404 illustration, and the architect-elevator image.
- `favicon.ico`, `apple-touch-icon.png` — the icon shared by every indrayam.com site.
- `robots.txt`, `sitemap.xml` — crawler files for this host.
- `vercel.json` — the shared security headers, plus `Content-Type: application/pdf` and `Content-Disposition: inline` on the PDF.
- `.vercelignore` — keeps `AGENTS.md`, `README.md`, `build_resume_docx.py`, `references/`, `tmp/`, `tools/`, `.env*`, and `.claude/` out of deployments.
- `tmp/` — ignored scratch space (see `AGENTS.md`).

## DOCX and PDF

The HTML page, the DOCX, and the PDF are not generated from one source. When resume facts change in `references/resume.md`, update each representation and keep them consistent.

- **DOCX (current workflow).** Since November 2026 the DOCX is edited in place from the Word file, so its fonts (Montserrat), layout, and three-page structure come from the original document. The scripts are in `tools/docx-update/`, and each takes `in.docx out.docx`:
  - `edit.py` — applies the November 2026 `resume.md` edits to a copy of the Word resume by cloning existing paragraphs and runs, so formatting is unchanged. `--photo=FILE.jpg` swaps the header portrait (the current one is `references/resume-photo.jpg`) and `--move` re-flows paragraphs. It adds the sparkles icon from `icon/`.
  - `move_overview.py` — moves Career Overview and the first role to the top of page 2.
  - `skills_nov23.py` — the November 23 sidebar edit: replaces the Dev Tools line and adds an AI Coding Harness row at the top of the EXPERIENCE (skills) list.
  - `italic.py in.docx out.docx WORD` — italicizes every occurrence of a word inside its run.

  `edit.py` and `move_overview.py` assert exact paragraph positions in the document version they were written for, so they record how the November edits were made rather than replay from scratch; copy and adapt them for the next round of edits. `italic.py` and `export.sh` work on any version. Always write to a copy and check the rendered pages before replacing `output/docx/`.
- **PDF.** Exported from the DOCX by Microsoft Word, so it matches the DOCX exactly: `tools/docx-update/export.sh in.docx out.pdf` drives Word through AppleScript. Word needs the Montserrat fonts installed (they are in `~/Library/Fonts` on this Mac).
- **`build_resume_docx.py` (older generator).** Builds a DOCX from text hard-coded in the script, which was last updated in September 2026. Its output does not match the published DOCX, and running it overwrites `output/docx/anand-sharma-resume.docx`. Use it only after bringing its content up to date, or restore the published file with `git checkout -- output/docx/` afterwards. It reads the header portrait from `references/resume-photo.jpg` and writes the circular crop to `tmp/`. Run it with `uv run --no-project --with python-docx --with pillow python build_resume_docx.py`.

## Working on the repository

Read [`AGENTS.md`](AGENTS.md) before making changes. It covers source-of-truth handling, Git synchronization, and preservation and verification expectations. This repository is part of the wider `indrayam.com` site; the parent folder's `AGENTS.md` applies when work crosses repository boundaries.

The site has no build step. Preview it with any static server (the workspace's `.claude/launch.json` has a `resume` entry on port 8083). Vercel (Git-connected, project `resume`) deploys on every push to `main`.
