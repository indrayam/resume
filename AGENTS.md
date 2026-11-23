# Working in this repository

## Before editing

- Check the current branch, working-tree status, and upstream before changing files. Preserve any existing user changes.
- Fetch the upstream branch when network access is available. On a clean private branch, rebase onto its upstream when that is the established workflow. Do not rebase a shared or default branch. If the branch is behind or diverged, or the working tree is dirty, explain the state and use a safe project workflow before editing; never discard changes to make the update fit.
- Read the relevant files and keep changes limited to the requested work.

## Resume content and site behavior

- Treat `references/resume.md` as the factual source for resume content. Do not invent, embellish, or silently alter career facts, dates, titles, metrics, or claims.
- The HTML in `index.html` identifies `references/resume.md` as its content source. The downloadable DOCX and PDF in `output/` are separate representations; keep them consistent with `references/resume.md` when changing resume content, while respecting their format-specific layout.
- The published DOCX is edited in place from the Word file and the PDF is exported from it by Microsoft Word; `README.md` describes the scripts. Do not run `build_resume_docx.py` to refresh `output/`: its hard-coded content dates from September 2026 and would overwrite the published DOCX with an older resume. If you run it for another reason, restore `output/docx/` afterwards.
- Nothing in `output/` is regenerated automatically. Do not claim to have rebuilt an output you did not rebuild and render.
- Preserve the site's existing visual design, responsive behavior, accessibility, working links, and shared navigation unless the request explicitly calls for a change.
- Avoid unrelated dependencies, tooling, or redesigns. Do not overwrite original reference documents without an explicit request.
- Use `tmp/` as scratch space for content generation: renders, intermediate DOCX/PDF builds, diffs. Git and Vercel both ignore it, so never rely on anything there surviving or shipping. Move anything worth keeping, such as an alternate version of a document, into `references/`. The DOCX editing scripts currently live in the ignored `tmp/docx-update/`; do not delete them during cleanup.

## Verification

- For content changes, review the affected HTML and document representations for consistency and check the links they expose.
- For layout or styling changes, inspect the affected pages at desktop and mobile sizes. A syntax check alone does not verify visual behavior.
- Report what you inspected or rendered; do not claim visual or output verification that you did not perform.
- Follow the parent `AGENTS.md` when work crosses into the wider site repository.
