from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

PNG_PATH = ASSETS / "lobster-icon.png"
ICO_PATH = ASSETS / "lobster-icon.ico"


def draw_lobster_icon(draw: ImageDraw.ImageDraw, size: int) -> None:
    scale = size / 512.0

    def sx(value: float) -> float:
        return value * scale

    sand = "#F5E8D9"
    sand_soft = "#FCF6EF"
    ink = "#161616"
    coral = "#FF6B35"
    coral_soft = "#FF9A73"
    line = "#FFF8EF"

    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((sx(52), sx(60), sx(460), sx(468)), radius=sx(128), fill=(0, 0, 0, 40))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=sx(18)))
    draw.bitmap((0, 0), shadow)

    draw.rounded_rectangle((sx(44), sx(44), sx(468), sx(468)), radius=sx(136), fill=sand)
    draw.rounded_rectangle((sx(70), sx(70), sx(442), sx(442)), radius=sx(112), fill=sand_soft)

    draw.line((sx(228), sx(112), sx(196), sx(74)), fill=ink, width=int(sx(8)))
    draw.line((sx(284), sx(112), sx(316), sx(74)), fill=ink, width=int(sx(8)))

    draw.polygon(
        [
            (sx(122), sx(154)),
            (sx(86), sx(126)),
            (sx(100), sx(94)),
            (sx(152), sx(112)),
            (sx(166), sx(146)),
            (sx(146), sx(172)),
        ],
        fill=ink,
    )
    draw.polygon(
        [
            (sx(390), sx(154)),
            (sx(426), sx(126)),
            (sx(412), sx(94)),
            (sx(360), sx(112)),
            (sx(346), sx(146)),
            (sx(366), sx(172)),
        ],
        fill=ink,
    )

    draw.line((sx(176), sx(172), sx(130), sx(198)), fill=ink, width=int(sx(11)))
    draw.line((sx(336), sx(172), sx(382), sx(198)), fill=ink, width=int(sx(11)))
    draw.line((sx(186), sx(262), sx(130), sx(286)), fill=ink, width=int(sx(9)))
    draw.line((sx(194), sx(304), sx(140), sx(332)), fill=ink, width=int(sx(9)))
    draw.line((sx(326), sx(262), sx(382), sx(286)), fill=ink, width=int(sx(9)))
    draw.line((sx(318), sx(304), sx(372), sx(332)), fill=ink, width=int(sx(9)))

    draw.ellipse((sx(184), sx(138), sx(328), sx(306)), fill=coral)
    draw.ellipse((sx(206), sx(162), sx(306), sx(282)), fill=coral_soft)

    draw.arc((sx(198), sx(156), sx(314), sx(228)), start=196, end=340, fill=line, width=int(sx(7)))
    draw.arc((sx(192), sx(186), sx(320), sx(258)), start=196, end=340, fill=line, width=int(sx(7)))
    draw.arc((sx(200), sx(216), sx(312), sx(286)), start=196, end=340, fill=line, width=int(sx(7)))

    draw.rounded_rectangle((sx(208), sx(292), sx(304), sx(332)), radius=sx(18), fill=coral)
    draw.rounded_rectangle((sx(220), sx(324), sx(292), sx(362)), radius=sx(18), fill=coral_soft)
    draw.polygon(
        [
            (sx(228), sx(362)),
            (sx(194), sx(412)),
            (sx(242), sx(400)),
            (sx(254), sx(366)),
        ],
        fill=ink,
    )
    draw.polygon(
        [
            (sx(284), sx(362)),
            (sx(318), sx(412)),
            (sx(270), sx(400)),
            (sx(258), sx(366)),
        ],
        fill=ink,
    )


def main() -> None:
    size = 512
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw_lobster_icon(draw, size)
    image.save(PNG_PATH)
    image.save(ICO_PATH, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])


if __name__ == "__main__":
    main()
