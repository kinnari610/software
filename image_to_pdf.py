"""Convert one or more images into a PDF file.

Usage:
    python image_to_pdf.py input.png output.pdf
    python image_to_pdf.py page1.png page2.png output.pdf

This script uses Pillow (PIL). It preserves the image dimensions and embeds each image as a PDF page.
"""

import sys
from pathlib import Path

from PIL import Image


def _die(msg: str, code: int = 1):
    print(msg, file=sys.stderr)
    sys.exit(code)


def images_to_pdf(image_paths, out_pdf_path, dpi=300):
    # Load images and convert to RGB (PDF doesn't support transparency well).
    images = []
    for p in image_paths:
        img = Image.open(p)
        if img.mode in ("RGBA", "LA"):
            # Convert transparency -> white background
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)

    if not images:
        _die("No images to convert.")

    first, rest = images[0], images[1:]

    # Save as PDF. Pillow will use the provided DPI as the page resolution.
    first.save(out_pdf_path, "PDF", resolution=dpi, save_all=True, append_images=rest)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) < 2:
        _die("Usage: python image_to_pdf.py <input-image> [<more-images>...] <output.pdf>")

    out_pdf = Path(argv[-1])
    input_paths = [Path(p) for p in argv[:-1]]

    for p in input_paths:
        if not p.exists():
            _die(f"Input file not found: {p}")

    images_to_pdf(input_paths, out_pdf)
    print(f"Written PDF: {out_pdf}")


if __name__ == "__main__":
    main()
