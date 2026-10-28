#!/bin/bash
# render.sh <pptx> <prefix> : PDF + PNG (110 dpi) + phone-width PNG
set -e
SK=/root/.claude/skills/synced/33660d22-ca00-4e96-93fd-5108ab32a209_639bc7d5-e5bd-4aa4-8625-1eb0b1c12b0c/pptx
f="$1"; pre="$2"; d=$(dirname "$f"); b=$(basename "$f" .pptx)
python3 $SK/scripts/office/soffice.py --headless --convert-to pdf --outdir "$d" "$f" >/dev/null 2>&1
rm -f "$d/$pre"-*.png
pdftoppm -png -r 110 "$d/$b.pdf" "$d/$pre"
for p in "$d/$pre"-*.png; do convert "$p" -resize 400x "${p%.png}-phone.png"; done
ls -1 "$d/$pre"-*.png
