#!/usr/bin/env python3
"""Generate a 48×48 RGB565 BMP placeholder for pet avatar testing."""

import struct
import sys
from pathlib import Path

W, H = 48, 48


def rgb565(r: int, g: int, b: int) -> int:
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)


def write_bmp(path: Path) -> None:
    row_bytes = W * 2
    pixel_data_size = row_bytes * H
    file_size = 54 + pixel_data_size

    with path.open("wb") as f:
        # BMP header
        f.write(b"BM")
        f.write(struct.pack("<I", file_size))
        f.write(struct.pack("<HH", 0, 0))
        f.write(struct.pack("<I", 54))
        # DIB header
        f.write(struct.pack("<I", 40))
        f.write(struct.pack("<ii", W, H))
        f.write(struct.pack("<HH", 1, 16))
        f.write(struct.pack("<I", 0))
        f.write(struct.pack("<I", pixel_data_size))
        f.write(struct.pack("<ii", 2835, 2835))
        f.write(struct.pack("<II", 0, 0))

        # Pixels bottom-up
        for y in range(H - 1, -1, -1):
            for x in range(W):
                # Simple paw-like circle pattern
                cx, cy = W // 2, H // 2
                dist = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5
                if dist < 18:
                    color = rgb565(255, 180, 100)
                elif dist < 22:
                    color = rgb565(200, 120, 60)
                else:
                    color = rgb565(40, 40, 80)
                f.write(struct.pack("<H", color))


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent.parent / "data" / "pet.bmp"
    out.parent.mkdir(parents=True, exist_ok=True)
    write_bmp(out)
    print(f"Wrote {out} ({W}x{H} RGB565 BMP)")
