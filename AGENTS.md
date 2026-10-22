# Working in this repository

## Before editing

- Check the current branch, working-tree status, and upstream before changing files. Preserve any existing user changes.
- Fetch the upstream branch when network access is available. On a clean private branch, rebase onto its upstream when that is the established workflow. Do not rebase a shared or default branch. If the branch is behind or diverged, or the working tree is dirty, explain the state and use a safe project workflow before editing; never discard changes to make the update fit.
- Read the relevant files and keep changes limited to the requested work.

## Resume content and site behavior

- Treat `references/resume.md` as the factual source for resume content. Do not invent, embellish, or silently alter career facts, dates, titles, metrics, or claims.
- The HTML in `index.html` identifies `references/resume.md` as its content source. `build_resume_docx.py` contains a separate hard-coded DOCX representation; it does not read that Markdown file. Keep representations consistent when changing resume content, while respecting their format-specific layout.
- Do not assume the PDF or other files in `output/` are regenerated automatically. Inspect the available workflow before changing or claiming to have rebuilt an output.
- Preserve the site's existing visual design, responsive behavior, accessibility, working links, and shared navigation unless the request explicitly calls for a change.
- Avoid unrelated dependencies, tooling, or redesigns. Do not overwrite original reference documents without an explicit request.

## Verification

- For content changes, review the affected HTML and document representations for consistency and check the links they expose.
- For layout or styling changes, inspect the affected pages at desktop and mobile sizes. A syntax check alone does not verify visual behavior.
- Report what you inspected or rendered; do not claim visual or output verification that you did not perform.
- Follow the parent `AGENTS.md` when work crosses into the wider site repository.
