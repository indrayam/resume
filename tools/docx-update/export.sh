#!/bin/bash
# Usage: export.sh in.docx out.pdf  — has Microsoft Word save the DOCX as PDF.
IN="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"; OUT="$(cd "$(dirname "$2")" && pwd)/$(basename "$2")"
osascript <<EOF
tell application "Microsoft Word"
  open file name (POSIX file "$IN" as text)
  delay 3
  save as active document file name (POSIX file "$OUT" as text) file format format PDF
  close active document saving no
end tell
EOF
