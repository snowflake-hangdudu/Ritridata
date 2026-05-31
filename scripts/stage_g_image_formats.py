#!/usr/bin/env python3
"""Generate G-drive image preview samples in all product-supported formats.

Reads source photos from scripts/user_resources/图片/ (or --sources), writes to
G:\\图片\\ with naming G_<ext>_<seq>_<yyMMdd_HHmm>.<ext>.

Supported via conversion: jpg, jpeg, png, gif, bmp, tiff, webp, heic, svg, psd
Supported via template copy: raw, cr2, nef, arw (place real files under
scripts/user_resources/raw/ or anywhere under user_resources/)
"""

from __future__ import annotations

import argparse
import base64
import io
import struct
import sys
from datetime import datetime
from pathlib import Path

from PIL import Image

try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    HAS_HEIF = True
except ImportError:
    HAS_HEIF = False

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE_DIR = SCRIPT_DIR / "user_resources" / "图片"
DEFAULT_DRIVE = Path("G:/")
IMAGE_FOLDER = "\u56fe\u7247"  # 图片

# Matches product image type filters (two menu pages).
CONVERT_EXTENSIONS = (
    "jpg",
    "jpeg",
    "png",
    "gif",
    "bmp",
    "tiff",
    "webp",
    "heic",
    "svg",
    "psd",
)
TEMPLATE_EXTENSIONS = ("raw", "cr2", "nef", "arw")
ALL_EXTENSIONS = CONVERT_EXTENSIONS + TEMPLATE_EXTENSIONS

SOURCE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Stage G-drive image format preview samples.")
    parser.add_argument("--drive", default=str(DEFAULT_DRIVE), help="Target drive root (default: G:\\)")
    parser.add_argument(
        "--sources",
        nargs="*",
        default=[],
        help="Source image files or directories (default: scripts/user_resources/图片/)",
    )
    parser.add_argument("--per-format", type=int, default=3, help="Files per extension (default: 3)")
    parser.add_argument(
        "--clear-image-folder",
        action="store_true",
        help="Remove existing files under <drive>/图片/ before generating",
    )
    parser.add_argument(
        "--formats",
        nargs="*",
        default=[],
        help="Subset of extensions to generate (default: all product image types)",
    )
    return parser.parse_args()


def collect_sources(explicit: list[str]) -> list[Path]:
    paths: list[Path] = []
    if explicit:
        for item in explicit:
            p = Path(item)
            if p.is_dir():
                paths.extend(sorted(p.rglob("*")))
            elif p.is_file():
                paths.append(p)
    else:
        if not DEFAULT_SOURCE_DIR.is_dir():
            raise SystemExit(f"Source folder not found: {DEFAULT_SOURCE_DIR}")
        paths.extend(sorted(DEFAULT_SOURCE_DIR.rglob("*")))

    sources = [p for p in paths if p.is_file() and p.suffix.lower() in SOURCE_SUFFIXES]
    if not sources:
        raise SystemExit("No source images found (.jpg/.jpeg/.png/.webp/.bmp/.gif/.tif/.tiff).")
    return sources


def collect_templates(user_resource_dir: Path) -> dict[str, list[Path]]:
    by_ext: dict[str, list[Path]] = {ext: [] for ext in TEMPLATE_EXTENSIONS}
    if not user_resource_dir.is_dir():
        return by_ext
    for path in sorted(user_resource_dir.rglob("*")):
        if not path.is_file():
            continue
        ext = path.suffix.lstrip(".").lower()
        if ext in by_ext:
            by_ext[ext].append(path)
    return by_ext


def unique_name(ext: str, index: int, run_time: str) -> str:
    return f"G_{ext}_{index:03d}_{run_time}.{ext}"


def save_minimal_psd(path: Path, image: Image.Image) -> None:
    rgb = image.convert("RGB")
    width, height = rgb.size
    r, g, b = rgb.split()
    planes = [plane.tobytes() for plane in (r, g, b)]

    def section(data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + data

    header = (
        b"8BPS"
        + struct.pack(">H", 1)
        + b"\0" * 6
        + struct.pack(">H", 3)
        + struct.pack(">I", height)
        + struct.pack(">I", width)
        + struct.pack(">H", 8)
        + struct.pack(">H", 3)
    )
    color_mode = section(b"")
    resources = section(b"")
    layers = section(b"")
    compression = struct.pack(">H", 0) + b"".join(planes)
    image_data = section(compression)
    path.write_bytes(header + color_mode + resources + layers + image_data)


def save_svg_embedded(path: Path, image: Image.Image) -> None:
    rgb = image.convert("RGB")
    width, height = rgb.size
    buf = io.BytesIO()
    rgb.save(buf, format="JPEG", quality=90, optimize=True)
    encoded = base64.b64encode(buf.getvalue()).decode("ascii")
    svg = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'width="{width}" height="{height}" viewBox="0 0 {width} {height}">\n'
        f'  <title>Ritridata image preview sample</title>\n'
        f'  <image width="{width}" height="{height}" '
        f'xlink:href="data:image/jpeg;base64,{encoded}"/>\n'
        "</svg>\n"
    )
    path.write_text(svg, encoding="utf-8")


def save_converted(path: Path, ext: str, image: Image.Image) -> None:
    rgb = image.convert("RGB")
    if ext in ("jpg", "jpeg"):
        rgb.save(path, format="JPEG", quality=90, optimize=True)
        return
    if ext == "png":
        rgb.save(path, format="PNG", optimize=True)
        return
    if ext == "gif":
        rgb.save(path, format="GIF", optimize=True)
        return
    if ext == "bmp":
        rgb.save(path, format="BMP")
        return
    if ext == "tiff":
        rgb.save(path, format="TIFF", compression="tiff_deflate")
        return
    if ext == "webp":
        rgb.save(path, format="WEBP", quality=85, method=4)
        return
    if ext == "heic":
        if not HAS_HEIF:
            raise RuntimeError("pillow-heif not installed")
        pillow_heif.from_pillow(rgb).save(path, quality=85)
        return
    if ext == "svg":
        save_svg_embedded(path, rgb)
        return
    if ext == "psd":
        save_minimal_psd(path, rgb)
        return
    raise ValueError(f"unsupported extension: {ext}")


def clear_image_folder(out_dir: Path) -> None:
    if not out_dir.is_dir():
        return
    for item in out_dir.iterdir():
        if item.is_file():
            item.unlink()
            print(f"Removed: {item}")


def main() -> int:
    args = parse_args()
    drive_root = Path(args.drive)
    if not drive_root.exists():
        print(f"Target drive not found: {drive_root}", file=sys.stderr)
        return 1

    extensions = [e.lower().lstrip(".") for e in args.formats] if args.formats else list(ALL_EXTENSIONS)
    unknown = [e for e in extensions if e not in ALL_EXTENSIONS]
    if unknown:
        print(f"Unknown extensions: {', '.join(unknown)}", file=sys.stderr)
        return 1

    sources = collect_sources(args.sources)
    templates = collect_templates(SCRIPT_DIR / "user_resources")
    run_time = datetime.now().strftime("%y%m%d_%H%M")
    out_dir = drive_root / IMAGE_FOLDER
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.clear_image_folder:
        clear_image_folder(out_dir)

    print(f"Sources ({len(sources)}):")
    for src in sources:
        print(f"  {src}")
    print(f"Output: {out_dir}")
    print(f"Formats: {', '.join(extensions)}")
    print(f"Per format: {args.per_format}")

    global_index = 1
    generated: set[str] = set()
    skipped: dict[str, str] = {}

    convert_exts = [e for e in extensions if e in CONVERT_EXTENSIONS]
    for ext in convert_exts:
        if ext == "heic" and not HAS_HEIF:
            skipped[ext] = "install pillow-heif"
            continue
        try:
            for i in range(args.per_format):
                source = sources[(i) % len(sources)]
                image = Image.open(source)
                name = unique_name(ext, global_index, run_time)
                target = out_dir / name
                save_converted(target, ext, image)
                print(f"Created ({ext}): {target}")
                global_index += 1
            generated.add(ext)
        except Exception as exc:
            skipped[ext] = str(exc)
            print(f"Warning: failed .{ext}: {exc}", file=sys.stderr)

    template_exts = [e for e in extensions if e in TEMPLATE_EXTENSIONS]
    for ext in template_exts:
        pool = templates.get(ext, [])
        if not pool:
            skipped[ext] = f"add real .{ext} under scripts/user_resources/raw/"
            continue
        for i in range(args.per_format):
            template = pool[i % len(pool)]
            name = unique_name(ext, global_index, run_time)
            target = out_dir / name
            target.write_bytes(template.read_bytes())
            print(f"Created (template {ext}): {target}")
            global_index += 1
        generated.add(ext)

    missing = [e for e in extensions if e not in generated]
    if missing:
        print("Warning: not generated:", ", ".join(missing))
        for ext in missing:
            reason = skipped.get(ext, "unknown")
            print(f"  .{ext}: {reason}")

    total = global_index - 1
    print(f"Done. Total files: {total}")
    print("Naming: G_<ext>_<seq>_<yyMMdd_HHmm>.<ext>")
    return 0 if total > 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
