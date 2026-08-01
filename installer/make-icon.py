from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"


def build_icon() -> Image.Image:
    scale = 4
    size = 256 * scale
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def box(coords: tuple[int, int, int, int], radius: int, fill: str, outline: str | None = None, width: int = 1) -> None:
        draw.rounded_rectangle(tuple(value * scale for value in coords), radius=radius * scale, fill=fill, outline=outline, width=width * scale)

    box((8, 8, 248, 248), 48, "#0b1220")
    box((38, 45, 218, 102), 14, "#111a2b", "#38bdf8", 6)
    box((38, 105, 218, 162), 14, "#162033", "#38bdf8", 6)
    box((38, 165, 218, 222), 14, "#1d2a42", "#38bdf8", 6)

    for y, accent in ((74, "#f8fafc"), (134, "#f8fafc"), (194, "#f8fafc")):
        draw.ellipse(((57 * scale), ((y - 7) * scale), (71 * scale), ((y + 7) * scale)), fill=accent)
        draw.rounded_rectangle(((84 * scale), ((y - 5) * scale), (194 * scale), ((y + 5) * scale)), radius=5 * scale, fill="#9aa9bf")

    return image.resize((256, 256), Image.Resampling.LANCZOS)


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    icon = build_icon()
    icon.save(ASSETS / "contextpack-icon.png", optimize=True)
    icon.save(ASSETS / "contextpack.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    main()
