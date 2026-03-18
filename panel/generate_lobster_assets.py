from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

PNG_PATH = ASSETS / "lobster-icon.png"
ICO_PATH = ASSETS / "lobster-icon.ico"


def draw_rounded_backplate(image: Image.Image, size: int) -> None:
    scale = size / 512.0

    def sx(value: float) -> float:
        return value * scale

    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle(
        (sx(52), sx(60), sx(460), sx(468)),
        radius=sx(132),
        fill=(0, 0, 0, 36),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=sx(18)))
    image.alpha_composite(shadow)

    face = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    face_draw = ImageDraw.Draw(face)
    face_draw.rounded_rectangle((sx(44), sx(44), sx(468), sx(468)), radius=sx(136), fill="#FCFCFD")
    face_draw.rounded_rectangle((sx(60), sx(60), sx(452), sx(452)), radius=sx(118), fill="#FFFFFF")
    image.alpha_composite(face)


def draw_sparkle(draw: ImageDraw.ImageDraw, size: int) -> None:
    scale = size / 512.0

    def sx(value: float) -> float:
        return value * scale

    center_x = sx(120)
    center_y = sx(110)
    long_r = sx(52)
    short_r = sx(24)
    gold = "#FFCC3D"
    gold_core = "#FFE486"

    points = [
        (center_x, center_y - long_r),
        (center_x + sx(12), center_y - sx(12)),
        (center_x + long_r, center_y),
        (center_x + sx(12), center_y + sx(12)),
        (center_x, center_y + long_r),
        (center_x - sx(12), center_y + sx(12)),
        (center_x - long_r, center_y),
        (center_x - sx(12), center_y - sx(12)),
    ]
    draw.polygon(points, fill=gold)
    draw.polygon(
        [
            (center_x, center_y - short_r),
            (center_x + sx(8), center_y - sx(8)),
            (center_x + short_r, center_y),
            (center_x + sx(8), center_y + sx(8)),
            (center_x, center_y + short_r),
            (center_x - sx(8), center_y + sx(8)),
            (center_x - short_r, center_y),
            (center_x - sx(8), center_y - sx(8)),
        ],
        fill=gold_core,
    )


def draw_claw_icon(image: Image.Image, size: int) -> None:
    scale = size / 512.0

    def sx(value: float) -> float:
        return value * scale

    mask = Image.new("L", (size, size), 0)
    mask_draw = ImageDraw.Draw(mask)

    red = "#F34A45"
    red_dark = "#D72E30"
    red_light = "#FF8F88"
    red_soft = "#FFC4BC"
    white = "#FFF7F4"

    arm_mask = Image.new("L", (size, size), 0)
    arm_draw = ImageDraw.Draw(arm_mask)
    arm_draw.rounded_rectangle((sx(184), sx(240), sx(278), sx(446)), radius=sx(44), fill=255)
    arm_mask = arm_mask.rotate(29, resample=Image.Resampling.BICUBIC, center=(sx(230), sx(344)))
    mask = ImageChops.lighter(mask, arm_mask)
    mask_draw = ImageDraw.Draw(mask)

    mask_draw.ellipse((sx(246), sx(150), sx(430), sx(364)), fill=255)
    mask_draw.ellipse((sx(144), sx(134), sx(304), sx(286)), fill=255)
    mask_draw.ellipse((sx(192), sx(176), sx(302), sx(298)), fill=255)

    # Cut the jaw opening so the silhouette matches the uploaded claw more closely.
    mask_draw.ellipse((sx(230), sx(140), sx(338), sx(284)), fill=0)
    mask_draw.polygon(
        [(sx(216), sx(170)), (sx(274), sx(156)), (sx(244), sx(248)), (sx(198), sx(246))],
        fill=0,
    )

    base = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    base_draw = ImageDraw.Draw(base)
    base_draw.rectangle((0, 0, size, size), fill=red)
    base.putalpha(mask)

    shadow_alpha = mask.filter(ImageFilter.GaussianBlur(radius=sx(10)))
    shadow = Image.new("RGBA", (size, size), (187, 37, 40, 0))
    shadow.putalpha(shadow_alpha)
    image.alpha_composite(shadow, dest=(0, 0))

    image.alpha_composite(base)

    detail = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    detail_draw = ImageDraw.Draw(detail)
    detail_draw.ellipse((sx(276), sx(188), sx(384), sx(314)), fill=red_light)
    detail_draw.ellipse((sx(302), sx(212), sx(352), sx(270)), fill=red_soft)
    detail_draw.ellipse((sx(164), sx(156), sx(272), sx(256)), fill=red_light)
    detail_draw.ellipse((sx(186), sx(180), sx(232), sx(224)), fill=red_soft)
    detail_draw.ellipse((sx(214), sx(216), sx(270), sx(278)), fill=white)
    detail_draw.ellipse((sx(282), sx(238), sx(328), sx(292)), fill=white)
    detail_draw.polygon(
        [(sx(192), sx(156)), (sx(246), sx(108)), (sx(220), sx(220))],
        fill=red_dark,
    )
    detail_draw.polygon(
        [(sx(278), sx(150)), (sx(334), sx(114)), (sx(296), sx(226))],
        fill=red_dark,
    )
    detail_alpha = ImageChops.multiply(detail.getchannel("A"), mask)
    detail.putalpha(detail_alpha)
    image.alpha_composite(detail)


def draw_lobster_icon(image: Image.Image, size: int) -> None:
    draw_rounded_backplate(image, size)
    draw = ImageDraw.Draw(image)
    draw_sparkle(draw, size)
    draw_claw_icon(image, size)


def main() -> None:
    size = 512
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw_lobster_icon(image, size)
    image.save(PNG_PATH)
    image.save(ICO_PATH, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])


if __name__ == "__main__":
    main()
