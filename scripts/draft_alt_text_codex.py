# Quick start (PowerShell):
#   $env:OPENAI_API_KEY="your_api_key_here"
#   py -3 scripts/draft_alt_text_codex.py "Class 1 - Introduction.qmd"
#   py -3 scripts/draft_alt_text_codex.py "Class 1 - Introduction.qmd" --dry-run
#   py -3 scripts/draft_alt_text_codex.py "Class 1 - Introduction.qmd" --model gpt-4.1-mini
#!/usr/bin/env python3
"""
Draft alt-text for figures in a Quarto (.qmd) file.

Behavior:
1) Reads a .qmd file.
2) Finds markdown image tags that already include a `fig-alt=...` attribute.
3) Resolves each linked image file.
4) Uses a vision model to generate:
   - 2-4 sentence alt-text emphasizing economic significance.
5) Writes the .qmd back with blank image captions and updated fig-alt values.

Nothing else in the file is modified.
"""

from __future__ import annotations

import argparse
import base64
import io
import mimetypes
import json
import re
import sys
import tempfile
from pathlib import Path
from typing import List, Tuple


IMAGE_WITH_ATTRS_RE = re.compile(
    r"!\[(?P<caption>[^\]]*)\]\((?P<path>[^)]+)\)\{(?P<attrs>[^}]*)\}",
    flags=re.MULTILINE | re.DOTALL,
)
FIG_ALT_RE = re.compile(r"""fig-alt\s*=\s*(?P<q>["'])(?P<val>.*?)(?P=q)""", flags=re.DOTALL)
HEADING_RE = re.compile(r"(?m)^#\s+(.+?)\s*(?:\{.*\})?\s*$")


def _clean_caption(caption: str) -> str:
    c = caption.strip()
    c = re.sub(r"\s+[—-]\s*figure\s*$", "", c, flags=re.IGNORECASE)
    c = re.sub(r"\s+figure\s*$", "", c, flags=re.IGNORECASE)
    return c.strip() or "Untitled figure"


def _find_slide_block(text: str, pos: int) -> Tuple[str, List[str]]:
    headings = list(HEADING_RE.finditer(text))
    if not headings:
        return "Untitled slide", []

    current_idx = 0
    for i, h in enumerate(headings):
        if h.start() <= pos:
            current_idx = i
        else:
            break
    current = headings[current_idx]
    slide_title = current.group(1).strip()

    start = current.end()
    end = headings[current_idx + 1].start() if current_idx + 1 < len(headings) else len(text)
    block = text[start:end]

    bullets = []
    for line in block.splitlines():
        s = line.strip()
        if s.startswith("- "):
            bullets.append(s[2:].strip())
        if len(bullets) >= 6:
            break
    return slide_title, bullets


def _image_to_data_url(img_path: Path) -> str:
    mime, _ = mimetypes.guess_type(str(img_path))
    if not mime:
        mime = "image/png"
    blob = img_path.read_bytes()
    encoded = base64.b64encode(blob).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def _resolve_image_for_vision(img_path: Path) -> Tuple[Path, bool]:
    """
    Return a vision-compatible image path and whether it's a temporary file.
    Supported by API: jpeg/png/gif/webp.
    """
    # Prefer PNG first when a same-basename PNG exists (case-insensitive).
    stem = img_path.stem
    parent = img_path.parent
    for sibling in parent.glob(stem + ".*"):
        if sibling.suffix.lower() == ".png" and sibling.exists():
            return sibling, False

    supported = {".jpg", ".jpeg", ".png", ".gif", ".webp"}
    ext = img_path.suffix.lower()
    if ext in supported:
        return img_path, False

    if ext == ".svg":
        # Case-insensitive sibling fallback.
        for sibling in parent.glob(stem + ".*"):
            if sibling.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".gif"} and sibling.exists():
                return sibling, False

        try:
            import cairosvg  # type: ignore
        except Exception as exc:
            raise RuntimeError(
                f"Cannot send SVG directly to vision model and no raster sibling found for {img_path.name}. "
                "Install cairosvg or provide PNG/JPG exports alongside SVG files."
            ) from exc

        fd, tmp_name = tempfile.mkstemp(suffix=".png", prefix="alt_vision_")
        # Close descriptor immediately; cairo will write by filename.
        try:
            import os
            os.close(fd)
        except Exception:
            pass
        cairosvg.svg2png(url=str(img_path), write_to=tmp_name)
        return Path(tmp_name), True

    raise RuntimeError(
        f"Unsupported image format for vision request: {img_path.suffix}. "
        "Use JPG, PNG, GIF, WEBP, or SVG (with raster sibling/conversion)."
    )


def _to_clean_png_data_url(img_path: Path) -> str:
    """
    Re-encode input image into a normalized PNG payload to avoid invalid image errors.
    """
    try:
        from PIL import Image  # type: ignore
    except Exception:
        # Fallback to raw bytes if Pillow is unavailable.
        return _image_to_data_url(img_path)

    with Image.open(img_path) as im:
        # Normalize mode so save() always succeeds consistently.
        if im.mode not in ("RGB", "RGBA", "L"):
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        encoded = base64.b64encode(buf.getvalue()).decode("ascii")
        return f"data:image/png;base64,{encoded}"


def _draft_alt_text_with_vision(
    client,
    model: str,
    slide_title: str,
    bullets: List[str],
    caption: str,
    img_path: Path,
) -> str:
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found for alt-text generation: {img_path}")

    bullet_text = "\n".join(f"- {b}" for b in bullets[:6]) if bullets else "- (no bullets extracted)"
    prompt = (
        "You are writing accessibility text for an intermediate microeconomics lecture slide figure.\n"
        "Return ONLY plain alt-text as a single string.\n"
        "Do NOT return JSON.\n"
        "Do NOT return markdown code fences.\n"
        "Do NOT prefix with labels like 'alt_text:' or 'caption:'.\n"
        "Do NOT add surrounding quotes.\n"
        "Requirements for alt_text:\n"
        "1) Must be 2 to 4 sentences.\n"
        "2) Describe what is visually present (axes, curves, labels, shaded regions, arrows, key points).\n"
        "3) State the economic meaning in plain language.\n"
        "4) Mention notable values/equilibrium points if visible.\n"
        "5) Do not mention file names, pixels, or that this is an image.\n\n"
        f"Slide title: {slide_title}\n"
        f"Current caption: {caption}\n"
        f"Nearby slide bullets:\n{bullet_text}\n"
    )

    vision_img, is_temp = _resolve_image_for_vision(img_path)
    try:
        data_url = _to_clean_png_data_url(vision_img)
        resp = client.responses.create(
            model=model,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {"type": "input_image", "image_url": data_url},
                    ],
                }
            ],
            temperature=0.2,
        )
    finally:
        if is_temp:
            vision_img.unlink(missing_ok=True)
    text = (resp.output_text or "").strip()
    if not text:
        raise RuntimeError("Vision model returned empty response.")

    # Plain-text-first parsing.
    alt_out = _clean_model_alt_text_response(text)

    # Robust fallback: if model still returned JSON, use alt_text key when present.
    json_candidate = _strip_code_fences(text).strip()
    try:
        obj = json.loads(json_candidate)
        if isinstance(obj, dict) and obj.get("alt_text"):
            alt_out = _clean_model_alt_text_response(str(obj.get("alt_text")))
    except Exception:
        pass

    alt_out = _normalize_alt_text(alt_out)
    return alt_out


def _strip_code_fences(raw: str) -> str:
    txt = (raw or "").strip()
    m = re.match(r"^\s*```(?:json|text|markdown)?\s*(.*?)\s*```\s*$", txt, flags=re.IGNORECASE | re.DOTALL)
    if m:
        return m.group(1).strip()
    return txt


def _clean_model_alt_text_response(raw: str) -> str:
    txt = _strip_code_fences(raw)
    txt = txt.strip()
    txt = re.sub(r"^(?:alt[_\s-]*text|caption)\s*:\s*", "", txt, flags=re.IGNORECASE)
    txt = txt.strip().strip('"').strip("'").strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def _normalize_alt_text(raw_alt: str) -> str:
    txt = re.sub(r"\s+", " ", (raw_alt or "").strip())
    txt = txt.strip('"').strip()
    # Ensure 2-4 sentences by soft clipping.
    sentence_parts = re.split(r"(?<=[.!?])\s+", txt)
    sentence_parts = [s.strip() for s in sentence_parts if s.strip()]
    if len(sentence_parts) > 4:
        sentence_parts = sentence_parts[:4]
    if len(sentence_parts) < 2 and sentence_parts:
        sentence_parts.append("Economically, the figure illustrates the central mechanism discussed on this slide.")
    if not sentence_parts:
        sentence_parts = [
            "The figure shows the core variables and relationships highlighted on this slide.",
            "Economically, it summarizes the key mechanism and implication for equilibrium or welfare.",
        ]
    return " ".join(sentence_parts)


def _extract_image_path(raw_path: str) -> str:
    p = raw_path.strip()
    if p.startswith("<") and p.endswith(">"):
        p = p[1:-1].strip()
    if " " in p and p.count('"') >= 2:
        # Markdown image destination may include an optional title: (path "title")
        # Keep only path component.
        m = re.match(r'^(?P<dest>[^"]+?)\s+"[^"]*"\s*$', p)
        if m:
            p = m.group("dest").strip()
    return p


def update_qmd_alt_text(qmd_path: Path, client, model: str, dry_run: bool = False) -> int:
    text = qmd_path.read_text(encoding="utf-8")
    updates = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal updates
        caption = _clean_caption(match.group("caption"))
        attrs = match.group("attrs")
        path_raw = _extract_image_path(match.group("path"))
        alt_match = FIG_ALT_RE.search(attrs)
        if not alt_match:
            return match.group(0)

        img_abs = (qmd_path.parent / path_raw).resolve()
        slide_title, bullets = _find_slide_block(text, match.start())
        try:
            new_alt = _draft_alt_text_with_vision(
                client=client,
                model=model,
                slide_title=slide_title,
                bullets=bullets,
                caption=caption,
                img_path=img_abs,
            )
            new_alt = new_alt.replace('"', "'")
        except Exception as exc:
            print(
                f"[warn] Skipping alt-text update for image '{path_raw}': {exc}",
                file=sys.stderr,
            )
            return match.group(0)

        old_attr = alt_match.group(0)
        quote = alt_match.group("q")
        new_attr = f"fig-alt={quote}{new_alt}{quote}"
        new_attrs = attrs.replace(old_attr, new_attr, 1)
        updates += 1
        return f"![]({match.group('path')}){{{new_attrs}}}"

    updated_text = IMAGE_WITH_ATTRS_RE.sub(repl, text)

    if not dry_run and updates > 0:
        qmd_path.write_text(updated_text, encoding="utf-8")
    return updates


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Draft and inject fig-alt text for figures in a .qmd file using a vision model."
    )
    parser.add_argument("qmd", type=Path, help="Path to the .qmd file.")
    parser.add_argument(
        "--model",
        default="gpt-4.1-mini",
        help="Vision-capable model name to use for alt-text generation.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report updates without writing file.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    qmd_path = args.qmd.resolve()
    if not qmd_path.exists():
        raise FileNotFoundError(f"QMD not found: {qmd_path}")
    if qmd_path.suffix.lower() != ".qmd":
        raise ValueError("Input file must have .qmd extension.")

    try:
        from openai import OpenAI  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency 'openai'. Install with: pip install openai"
        ) from exc

    client = OpenAI()
    updates = update_qmd_alt_text(qmd_path, client=client, model=args.model, dry_run=args.dry_run)
    action = "Would update" if args.dry_run else "Updated"
    print(f"{action} {updates} figure alt-text entr{'y' if updates == 1 else 'ies'} in {qmd_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
