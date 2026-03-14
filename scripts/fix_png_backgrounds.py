# This script sets the background for .png files exported from PowerPoint to be white not transparent.
# To run from repo root: py scripts/fix_png_backgrounds.py <folder>
# Example: py scripts/fix_png_backgrounds.py lectures/lecture01/images

import argparse
from pathlib import Path
from PIL import Image

# =========================================================
# SETTINGS
# =========================================================

parser = argparse.ArgumentParser(description="Fix transparent PNG backgrounds to white.")
parser.add_argument("folder", nargs="?", default=None, help="Folder to scan (default: current directory)")
args = parser.parse_args()

ROOT_FOLDER = Path(args.folder) if args.folder else Path.cwd()

# Scan subfolders too
RECURSIVE = True

# Optional resizing for slide decks
RESIZE_LARGE_IMAGES = False
MAX_WIDTH = 2400
MAX_HEIGHT = 1800

# =========================================================
# HELPERS
# =========================================================

def has_transparency(img: Image.Image) -> bool:
    """Return True if the image has any transparent pixels."""
    if img.mode in ("RGBA", "LA"):
        alpha = img.getchannel("A")
        return alpha.getextrema()[0] < 255
    if img.mode == "P" and "transparency" in img.info:
        return True
    return False


def resize_if_needed(img: Image.Image) -> Image.Image:
    """Resize only if enabled and image exceeds max dimensions."""
    if not RESIZE_LARGE_IMAGES:
        return img

    if img.width <= MAX_WIDTH and img.height <= MAX_HEIGHT:
        return img

    out = img.copy()
    out.thumbnail((MAX_WIDTH, MAX_HEIGHT), Image.Resampling.LANCZOS)
    return out


def flatten_to_white(img: Image.Image) -> Image.Image:
    """Replace transparency with a white background."""
    rgba = img.convert("RGBA")
    bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
    return Image.alpha_composite(bg, rgba).convert("RGB")


def process_png(path: Path) -> str:
    """Process one PNG file and return a status message."""
    try:
        with Image.open(path) as img:
            transparent = has_transparency(img)

            if not transparent and not RESIZE_LARGE_IMAGES:
                return f"Skipped: {path}"

            out = flatten_to_white(img) if transparent else img.convert("RGB")
            out = resize_if_needed(out)
            out.save(path, format="PNG", optimize=True)

            if transparent:
                return f"Fixed transparency: {path}"
            return f"Resaved/resized: {path}"

    except Exception as e:
        return f"Error: {path} -> {e}"


def main():
    if not ROOT_FOLDER.exists():
        print(f"Folder does not exist: {ROOT_FOLDER}")
        return

    pattern = "**/*.png" if RECURSIVE else "*.png"
    png_files = list(ROOT_FOLDER.glob(pattern))

    if not png_files:
        print(f"No PNG files found in {ROOT_FOLDER}")
        return

    fixed = 0
    skipped = 0
    errors = 0

    print(f"Scanning folder: {ROOT_FOLDER}\n")

    for path in png_files:
        result = process_png(path)
        print(result)

        if result.startswith("Fixed transparency") or result.startswith("Resaved/resized"):
            fixed += 1
        elif result.startswith("Skipped"):
            skipped += 1
        else:
            errors += 1

    print("\nSummary")
    print(f"Updated: {fixed}")
    print(f"Skipped: {skipped}")
    print(f"Errors:  {errors}")


if __name__ == "__main__":
    main()
