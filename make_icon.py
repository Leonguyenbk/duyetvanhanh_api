# -*- coding: utf-8 -*-
"""Sinh icon.ico cho app Duyệt vận hành MPLIS: tài liệu xanh + dấu tick."""
from PIL import Image, ImageDraw


def draw_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size

    # --- nền tròn gradient-ish xanh dương ---
    pad = int(s * 0.04)
    circle_box = [pad, pad, s - pad, s - pad]
    top = (37, 99, 235)      # blue-600
    bottom = (29, 78, 216)   # blue-700
    steps = circle_box[3] - circle_box[1]
    for i in range(steps):
        t = i / max(steps - 1, 1)
        r = int(top[0] + (bottom[0] - top[0]) * t)
        g = int(top[1] + (bottom[1] - top[1]) * t)
        b = int(top[2] + (bottom[2] - top[2]) * t)
        y = circle_box[1] + i
        d.line([(circle_box[0], y), (circle_box[2], y)], fill=(r, g, b, 255))
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).ellipse(circle_box, fill=255)
    bg = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    bg.paste(img, (0, 0), mask)
    img = bg
    d = ImageDraw.Draw(img)

    # --- tờ giấy trắng, góc gấp ---
    doc_w = s * 0.40
    doc_h = s * 0.52
    doc_x = s * 0.32
    doc_y = s * 0.22
    fold = doc_w * 0.32

    paper = [
        (doc_x, doc_y),
        (doc_x + doc_w - fold, doc_y),
        (doc_x + doc_w, doc_y + fold),
        (doc_x + doc_w, doc_y + doc_h),
        (doc_x, doc_y + doc_h),
    ]
    d.polygon(paper, fill=(255, 255, 255, 255))
    fold_tri = [
        (doc_x + doc_w - fold, doc_y),
        (doc_x + doc_w, doc_y + fold),
        (doc_x + doc_w - fold, doc_y + fold),
    ]
    d.polygon(fold_tri, fill=(219, 234, 254, 255))  # blue-100

    # --- các dòng văn bản trên giấy ---
    line_color = (147, 197, 253, 255)  # blue-300
    lx0 = doc_x + doc_w * 0.16
    lx1 = doc_x + doc_w * 0.84
    for i, ly_ratio in enumerate((0.30, 0.44, 0.58)):
        ly = doc_y + doc_h * ly_ratio
        lw = max(int(s * 0.012), 1)
        x1 = lx1 if i != 2 else doc_x + doc_w * 0.62
        d.line([(lx0, ly), (x1, ly)], fill=line_color, width=lw)

    # --- huy hiệu tick tròn xanh lá, góc dưới phải ---
    badge_r = s * 0.20
    badge_cx = doc_x + doc_w * 0.98
    badge_cy = doc_y + doc_h * 0.98
    badge_box = [badge_cx - badge_r, badge_cy - badge_r, badge_cx + badge_r, badge_cy + badge_r]
    d.ellipse(
        [b + (s * 0.012 if i < 2 else -s * 0.012) for i, b in enumerate(badge_box)]
        if False else badge_box,
        fill=(22, 163, 74, 255),  # green-600
        outline=(255, 255, 255, 255),
        width=max(int(s * 0.02), 1),
    )
    # dấu tick
    tw = max(int(s * 0.022), 2)
    p1 = (badge_cx - badge_r * 0.45, badge_cy)
    p2 = (badge_cx - badge_r * 0.12, badge_cy + badge_r * 0.35)
    p3 = (badge_cx + badge_r * 0.48, badge_cy - badge_r * 0.38)
    d.line([p1, p2], fill=(255, 255, 255, 255), width=tw)
    d.line([p2, p3], fill=(255, 255, 255, 255), width=tw)

    return img


sizes = [16, 24, 32, 48, 64, 128, 256]
base = draw_icon(256)
imgs = [base.resize((s, s), Image.LANCZOS) if s != 256 else base for s in sizes]
base.save("icon.ico", sizes=[(s, s) for s in sizes])
print("Da tao icon.ico voi cac kich thuoc:", sizes)
