"""Optional: make the poster QR image from the command line.

    pip install "qrcode[pil]"
    python make_qr.py https://your-site.example.com
"""
import sys

import qrcode

base = sys.argv[1].rstrip("/") if len(sys.argv) > 1 else sys.exit("usage: python make_qr.py <site-url>")
img = qrcode.make(base + "/register", error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=20, border=4)
img.save("registration-qr.png")
print("Saved registration-qr.png ->", base + "/register")
