"""Make a contact sheet of slides exported by verify_powerpoint.ps1."""

import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", type=Path)
    parser.add_argument("--output", type=Path, default=Path("output/preview.png"))
    parser.add_argument("--columns", type=int, choices=(1, 2, 3, 4), default=2)
    parser.add_argument("--slides", help="Номера слайдов через запятую; по умолчанию все")
    args = parser.parse_args()
    files = sorted(args.directory.glob("slide-*.png"))
    if not files:
        parser.error("Нет slide-*.png; сначала экспортируйте слайды из PowerPoint")
    if args.slides:
        try:
            selected = [int(item) for item in args.slides.split(",")]
            if not selected or any(number < 1 or number > len(files) for number in selected):
                raise ValueError
            files = [files[number - 1] for number in selected]
        except ValueError:
            parser.error("--slides должен содержать существующие номера слайдов через запятую")
    rows = (len(files) + args.columns - 1) // args.columns
    preview = Image.new("RGB", (20 + args.columns * 620, rows * 375 + 20), "#e8edf3")
    draw = ImageDraw.Draw(preview)
    try:
        font = ImageFont.truetype("arial.ttf", 16)
    except OSError:
        font = ImageFont.load_default()
    for i, path in enumerate(files):
        x, y = 20 + i % args.columns * 620, 20 + i // args.columns * 375
        with Image.open(path) as source:
            thumb = source.convert("RGB")
            thumb.thumbnail((600, 338))
            preview.paste(thumb, (x, y))
        draw.text((x, y + 343), path.stem.removeprefix("slide-"), font=font, fill="#002f6c")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    preview.save(args.output)
    print(args.output.resolve())


if __name__ == "__main__":
    main()
