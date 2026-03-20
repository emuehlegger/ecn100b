# Renders a .qmd file to both HTML and RevealJS formats.
# Run from the repo root. Example:
#   python scripts/render_qmd.py "s2026/lectures/Class 1 - Introduction.qmd"

import subprocess
import sys
from pathlib import Path

if len(sys.argv) < 2:
    print("Usage: python render_qmd.py <file.qmd>")
    sys.exit(1)

qmd_path = Path(sys.argv[1]).resolve()
stem = qmd_path.stem
cwd = qmd_path.parent

html_output = f"{stem}.html"
deck_output = f"{stem}_deck.html"

subprocess.run(["quarto", "render", qmd_path.name, "--to", "html", "--output", html_output], check=True, cwd=cwd)
subprocess.run(["quarto", "render", qmd_path.name, "--to", "revealjs", "--output", deck_output], check=True, cwd=cwd)
