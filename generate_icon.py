from PIL import Image, ImageDraw
from pathlib import Path


BG       = (0, 114, 189, 255)   # Microsoft blue
DOC_BACK = (150, 200, 235, 255) # light blue-white
DOC_FRONT = (255, 255, 255, 255)
LINE     = (0, 80, 155, 210)
ARROW_BG = (16, 124, 16, 255)   # Microsoft green
ARROW_FG = (255, 255, 255, 255)


def draw_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size

    # Rounded square background
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=max(3, s // 5), fill=BG)

    m   = s * 0.13   # outer margin
    dw  = s * 0.44   # document width
    dh  = s * 0.54   # document height
    cr  = max(2, int(s * 0.05))
    off = s * 0.11   # offset of back doc

    # Back document
    d.rounded_rectangle(
        [m + off, m + off, m + off + dw, m + off + dh],
        radius=cr, fill=DOC_BACK,
    )

    # Front document
    d.rounded_rectangle(
        [m, m, m + dw, m + dh],
        radius=cr, fill=DOC_FRONT,
    )

    # Horizontal lines on front document (spreadsheet rows)
    lx1  = m + s * 0.07
    lh   = max(1, int(s * 0.042))
    gap  = dh * 0.185
    y0   = m + dh * 0.24
    lpad = s * 0.065
    for i, w_frac in enumerate([0.80, 0.80, 0.55]):
        ly  = y0 + i * gap
        lx2 = lx1 + (dw - lpad * 1.4) * w_frac
        d.rounded_rectangle([lx1, ly, lx2, ly + lh], radius=max(1, lh // 2), fill=LINE)

    # Arrow badge — green circle with "→" in bottom-right
    badge_r = s * 0.20
    bx = s - m * 0.4 - badge_r
    by = s - m * 0.4 - badge_r
    d.ellipse([bx - badge_r, by - badge_r, bx + badge_r, by + badge_r], fill=ARROW_BG)

    # Arrow head inside badge
    aw  = badge_r * 0.90
    ah  = badge_r * 0.48
    abh = ah * 0.40          # body half-height
    ax0 = bx - aw * 0.48
    ax1 = bx + aw * 0.48
    body_end = ax1 - ah * 0.80
    # body
    d.rectangle([ax0, by - abh, body_end, by + abh], fill=ARROW_FG)
    # head
    d.polygon([(body_end, by - ah), (ax1, by), (body_end, by + ah)], fill=ARROW_FG)

    return img


def main():
    sizes  = [256, 64, 48, 32, 16]
    images = [draw_icon(s) for s in sizes]
    out    = Path(__file__).parent / "icon.ico"
    images[0].save(
        out,
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=images[1:],
    )
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
