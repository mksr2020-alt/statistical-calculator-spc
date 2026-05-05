"""
generate_icon.py  —  Converts assets/icon.png  →  assets/icon.ico
Run this once before building the .exe.

Usage:
    python generate_icon.py

Requirements:
    pip install Pillow
"""

import sys
import os

def convert():
    try:
        from PIL import Image
    except ImportError:
        print("[generate_icon] Pillow not found — installing...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow", "--quiet"])
        from PIL import Image

    src = os.path.join("assets", "icon.png")
    dst = os.path.join("assets", "icon.ico")

    if not os.path.exists(src):
        print(f"[generate_icon] ERROR: '{src}' not found. Make sure assets/icon.png exists.")
        sys.exit(1)

    img = Image.open(src).convert("RGBA")

    # Windows ICO standard: include all sizes for crisp rendering at every scale
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    icons = []
    for size in sizes:
        resized = img.resize(size, Image.LANCZOS)
        icons.append(resized)

    icons[0].save(dst, format="ICO", sizes=sizes, append_images=icons[1:])
    print(f"[generate_icon] ✓  Saved {dst}  ({os.path.getsize(dst) // 1024} KB, {len(sizes)} sizes)")


if __name__ == "__main__":
    convert()
