from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFilter  # type: ignore
except Exception as e:  # pragma: no cover
    raise SystemExit("Pillow is required to generate the icons: pip install pillow")

W = 1024


def hex_(c: str):
    c = c.lstrip("#")
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))


def make_variant(mode: str, path: Path):
    img = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    if mode == "dark":
        bg = hex_("#0F172A"); ring = hex_("#C3B5FD"); bolt = (255, 255, 255)
        glow = (167, 139, 250, 90)
    else:
        bg = hex_("#FFFFFF"); ring = hex_("#A78BFA"); bolt = hex_("#6D28D9")
        glow = (167, 139, 250, 70)

    cx = cy = W // 2
    R = int(W * 0.42)
    ring_w = int(W * 0.08)

    # Shadow
    shadow = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.ellipse((cx - R, cy - R, cx + R, cy + R), fill=(0, 0, 0, 160))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=int(W * 0.03)))
    img.alpha_composite(shadow)

    # Ring + highlight
    draw.ellipse((cx - R, cy - R, cx + R, cy + R), outline=ring + (255,), width=ring_w)
    inset = int(ring_w * 0.35)
    draw.ellipse((cx - R + inset, cy - R + inset, cx + R - inset, cy + R - inset),
                 outline=(255, 255, 255, 60), width=max(1, int(ring_w * 0.25)))

    # Inner disc
    r2 = R - int(ring_w * 0.75)
    draw.ellipse((cx - r2, cy - r2, cx + r2, cy + r2), fill=bg + (255,))

    # Glow
    glow_im = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow_im)
    gd.ellipse((cx - int(r2 * 0.95), cy - int(r2 * 0.95), cx + int(r2 * 0.95), cy + int(r2 * 0.95)), fill=glow)
    glow_im = glow_im.filter(ImageFilter.GaussianBlur(radius=int(W * 0.05)))
    img.alpha_composite(glow_im)

    # Bolt
    bolt_scale = r2 * 0.9
    pts = [(0.05, -0.55), (-0.10, 0.05), (0.10, 0.05), (-0.05, 0.55), (0.20, -0.05), (0.02, -0.05)]
    poly = [(cx + int(x * bolt_scale), cy + int(y * bolt_scale)) for (x, y) in pts]
    stroke = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    sd = ImageDraw.Draw(stroke)
    sd.polygon(poly, fill=(0, 0, 0, 120))
    stroke = stroke.filter(ImageFilter.GaussianBlur(radius=int(W * 0.01)))
    img.alpha_composite(stroke)
    draw.polygon(poly, fill=bolt + (255,))

    out = img.resize((256, 256), Image.LANCZOS)
    out.save(path)
    return out


def main():
    icons_dir = Path(__file__).resolve().parent
    icons_dir.mkdir(exist_ok=True)
    import sys
    force = "--force" in sys.argv
    light = icons_dir / "cyme_logo_light.png"
    dark = icons_dir / "cyme_logo_dark.png"
    ico = icons_dir / "cyme_logo.ico"
    if force or not light.exists():
        make_variant("light", light)
    if force or not dark.exists():
        make_variant("dark", dark)
    if force or not ico.exists():
        im = Image.open(light)
        im.save(ico, sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
    print("Generated icons in:", icons_dir)


if __name__ == "__main__":
    main()
