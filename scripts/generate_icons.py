"""Generate minimal PNG icons for the Office Add-in manifest."""
import os, struct, zlib

def make_png(size, r, g, b):
    """Create a solid-color PNG of given size."""
    def chunk(name, data):
        c = struct.pack(">I", len(data)) + name + data
        crc = zlib.crc32(name + data) & 0xFFFFFFFF
        return c + struct.pack(">I", crc)

    sig = b'\x89PNG\r\n\x1a\n'
    ihdr_data = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    ihdr = chunk(b'IHDR', ihdr_data)

    raw = b''
    for _ in range(size):
        row = b'\x00' + bytes([r, g, b] * size)
        raw += row
    compressed = zlib.compress(raw, 9)
    idat = chunk(b'IDAT', compressed)
    iend = chunk(b'IEND', b'')
    return sig + ihdr + idat + iend

assets_dir = os.path.join(os.path.dirname(__file__), "..", "assets")
os.makedirs(assets_dir, exist_ok=True)

for size in [16, 32, 64, 80]:
    data = make_png(size, 16, 110, 190)  # #106ebe blue
    path = os.path.join(assets_dir, f"icon-{size}.png")
    with open(path, "wb") as f:
        f.write(data)
    print(f"Created: {path}")
