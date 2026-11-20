# Make a NEW plotter configuration "<base> - <suffix>.pc3" (+ its .pmp) with user-defined paper sizes, next to
# the standard ones in the AutoCAD profile's Plotters folder. The base .pc3 is only read, never changed.
# PC3/PMP file = 48-byte header + adler32(compressed) + uncompressed length + compressed length + zlib(text).
#   Scripts\python.exe tools\pc3_custom_size.py "DWG To PDF.pc3" "Langformat" 3950x914 [WxH ...]
#   Scripts\python.exe tools\pc3_custom_size.py --dump "DWG To PDF.pc3"
import os
import re
import struct
import sys
import zlib

PLOTTERS = os.path.expandvars(r"%APPDATA%\Autodesk\AutoCAD 2027\R26.0\enu\Plotters")
HEADER = b"PIAFILEVERSION_2.0,PC3VER1,compress\r\npmzlibcodec"


def read(path: str) -> str:
    """The text of a .pc3 / .pmp / .ctb / .stb (their headers differ only in the 'PC3VER1' part)."""
    raw = open(path, "rb").read()
    assert raw[:19] == HEADER[:19] and raw[35:48] == HEADER[35:], f"{path}: unknown header"
    chk, n, c = struct.unpack("<III", raw[48:60])
    comp = raw[60:60 + c]
    assert zlib.adler32(comp) == chk, f"{path}: checksum"
    return zlib.decompress(comp).decode("latin-1")


def write(path: str, text: str, header: bytes = HEADER) -> None:
    data = text.encode("latin-1")
    comp = zlib.compress(data, 9)
    open(path, "wb").write(header + struct.pack("<III", zlib.adler32(comp), len(data), len(comp)) + comp)


MEDIA = ('  abilities="500005500500505555000005550000000500000500000500000\n'
         '  caps_state="000000000000000000000000000000000000000000000000000\n'
         '  ui_owner="11111111111111111111110\n'
         '  size_max_x=5080.0\n  size_max_y=5080.0\n  max_roll_height=5080.0\n')


def pmp_text(meta: str, sizes: list[tuple[float, float]]) -> str:
    size, desc = [], []
    for i, (w, h) in enumerate(sizes):
        land = w >= h
        name = f"UserDefinedMetric ({w:.2f} x {h:.2f}MM)"
        mdn = (f"UserDefinedMetric {'Landscape' if land else 'Portrait'} {w:.2f}W x {h:.2f}H - (0, 0) x "
               f"({w:.0f}, {h:.0f}) ={w * h:.0f} MM")
        size.append(f"   {i}{{\n    caps_type=2\n    name=\"{name}\n    localized_name=\"User {i + 1} "
                    f"({w:.2f} x {h:.2f} MM)\n    media_description_name=\"{mdn}\n    media_group=15\n"
                    f"    landscape_mode={'TRUE' if land else 'FALSE'}\n   }}\n")
        desc.append(f"   {i}{{\n    caps_type=2\n    name=\"{mdn}\n    media_bounds_urx={w:.1f}\n"
                    f"    media_bounds_ury={h:.1f}\n    printable_bounds_llx=0.0\n    printable_bounds_lly=0.0\n"
                    f"    printable_bounds_urx={w:.1f}\n    printable_bounds_ury={h:.1f}\n"
                    f"    printable_area={w * h:.1f}\n    dimensional=TRUE\n   }}\n")
    block = " media{\n" + MEDIA + " }\n"
    return (meta + "mod{\n" + block + "}\ndel{\n" + block + "}\nudm{\n calibration{\n  _x=1.0\n  _y=1.0\n }\n"
            " media{\n" + MEDIA + "  size{\n" + "".join(size) + "  }\n  description{\n" + "".join(desc)
            + "  }\n }\n}\nhidden{\n" + block + "}\n")


def main(base: str, suffix: str, sizes: list[tuple[float, float]]) -> None:
    src = read(os.path.join(PLOTTERS, base))
    stem = os.path.splitext(base)[0] + f" - {suffix}"
    pmp = os.path.join(PLOTTERS, "PMP Files", stem + ".pmp")
    pc3 = os.path.join(PLOTTERS, stem + ".pc3")
    for p in (pmp, pc3):
        if os.path.exists(p):
            sys.exit(f"exists, not overwritten: {p}")
    path = 'user_defined_model_pathname="' + pmp.replace("/", "\\")
    text = re.sub(r'user_defined_model_pathname="[^\r\n]*', lambda m: path, src)
    meta = text[:text.index("}\n") + 2]
    write(pmp, pmp_text(meta.replace(" default_system_cfg=FALSE\n", ""), sizes))
    write(pc3, text)
    assert read(pmp) and read(pc3)
    print("written:", pc3, "|", pmp)


if __name__ == "__main__":
    if sys.argv[1] == "--dump":
        print(read(os.path.join(PLOTTERS, sys.argv[2])))
    else:
        main(sys.argv[1], sys.argv[2], [tuple(map(float, s.lower().split("x"))) for s in sys.argv[3:]])
