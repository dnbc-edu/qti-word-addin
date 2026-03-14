import struct
import zlib
from pathlib import Path


def chunk(tag: bytes, data: bytes) -> bytes:
    return (
        struct.pack('>I', len(data))
        + tag
        + data
        + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)
    )


def write_png(path: Path, width: int, height: int, pixel_fn):
    raw = bytearray()
    for y in range(height):
        raw.append(0)
        for x in range(width):
            r, g, b, a = pixel_fn(x, y, width, height)
            raw.extend([r, g, b, a])

    png = bytearray(b'\x89PNG\r\n\x1a\n')
    ihdr = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    png.extend(chunk(b'IHDR', ihdr))
    png.extend(chunk(b'IDAT', zlib.compress(bytes(raw), 9)))
    png.extend(chunk(b'IEND', b''))
    path.write_bytes(png)


def icon_pixel(x: int, y: int, w: int, h: int):
    bg = (37, 99, 235, 255)
    fg = (255, 255, 255, 255)

    cx = (w - 1) / 2
    cy = (h - 1) / 2
    r_outer = min(w, h) * 0.32
    r_inner = min(w, h) * 0.20

    dx = x - cx
    dy = y - cy
    d2 = dx * dx + dy * dy

    color = bg

    if r_inner * r_inner <= d2 <= r_outer * r_outer:
        color = fg

    if x > cx and y > cy and abs((y - cy) - 0.8 * (x - cx)) < max(1.0, w * 0.04):
        color = fg

    border = max(1, int(round(w * 0.04)))
    if x < border or y < border or x >= w - border or y >= h - border:
        color = (30, 64, 175, 255)

    return color


if __name__ == '__main__':
    out = Path('addin/assets')
    out.mkdir(parents=True, exist_ok=True)

    for size in (16, 32, 80):
        write_png(out / f'icon-{size}.png', size, size, icon_pixel)
        print(f'Wrote icon-{size}.png ({size}x{size})')
