# -*- coding: utf-8 -*-
from io import BytesIO
from PIL import Image

import gb5_config as C


def make_qr_png_bytes(qr_text, box_px=240):
    if C.QR_OK:
        import qrcode
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=0
        )
        qr.add_data(qr_text)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        img = img.resize((box_px, box_px))
        bio = BytesIO()
        img.save(bio, format="PNG", optimize=True)
        return bio.getvalue()

    img = Image.new("RGB", (box_px, box_px), (255, 255, 255))
    if C.PIL_DRAW_OK:
        from PIL import ImageDraw
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, box_px - 1, box_px - 1], outline=(0, 0, 0), width=3)
        d.text((box_px * 0.28, box_px * 0.40), "QR", fill=(0, 0, 0))
    bio = BytesIO()
    img.save(bio, format="PNG", optimize=True)
    return bio.getvalue()