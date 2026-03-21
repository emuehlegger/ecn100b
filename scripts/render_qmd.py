# Renders .qmd file(s) to both HTML and RevealJS formats.
# Run from the repo root. Examples:
#   python scripts/render_qmd.py "s2026/lectures/Class 1 - Introduction.qmd"
#   python scripts/render_qmd.py "s2026/lectures/*.qmd"

import subprocess
import sys
from pathlib import Path

if len(sys.argv) < 2:
    print("Usage: python render_qmd.py <file.qmd or glob pattern>")
    sys.exit(1)

pattern = sys.argv[1]
path = Path(pattern)

# Resolve glob or single file
if "*" in pattern or "?" in pattern:
    qmd_files = sorted(path.parent.glob(path.name))
else:
    qmd_files = [path]

if not qmd_files:
    print(f"No files matched: {pattern}")
    sys.exit(1)

for qmd_path in qmd_files:
    qmd_path = qmd_path.resolve()
    stem = qmd_path.stem
    cwd = qmd_path.parent

    html_output = f"{stem}.html"
    deck_output = f"{stem}_deck.html"

    print(f"\nRendering: {qmd_path.name}")
    subprocess.run(["quarto", "render", qmd_path.name, "--to", "html", "--output", html_output], check=True, cwd=cwd)
    subprocess.run(["quarto", "render", qmd_path.name, "--to", "revealjs", "--output", deck_output], check=True, cwd=cwd)
