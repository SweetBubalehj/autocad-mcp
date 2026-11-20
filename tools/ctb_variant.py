# Make a NEW plot style table from an existing .ctb in which the given ACI colours plot in their own (object)
# colour instead of the table's assignment - e.g. monochrome.ctb with 255 kept white (text masks of elevation
# blocks: monochrome prints them as black boxes) and the greys 8, 9, 250-254 kept grey. The base is only read.
# Same file format as .pc3: 48-byte header + adler32(compressed) + length + compressed length + zlib(text).
#   Scripts\python.exe tools\ctb_variant.py monochrome.ctb monochrome_grau.ctb 8 9 250 251 252 253 254 255
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pc3_custom_size import read, write  # noqa: E402

STYLES = os.path.expandvars(r"%APPDATA%\Autodesk\AutoCAD 2027\R26.0\enu\Plotters\Plot Styles")


def main(base: str, new: str, aci: list[int]) -> None:
    dst = os.path.join(STYLES, new)
    if os.path.exists(dst):
        sys.exit(f"exists, not overwritten: {dst}")
    text = read(os.path.join(STYLES, base))
    for n in aci:
        pat = re.compile(r'( name="Color_%d\n.*?color=)(-?\d+)(\n  color_policy=)(\d+)' % n, re.S)
        text, k = pat.subn(lambda m: m.group(1) + "-1" + m.group(3) + "1", text, count=1)
        if not k:
            sys.exit(f"Color_{n} not found in {base}")
    text = re.sub(r'^description="[^\n]*', f'description="{base} with object colour for ACI {" ".join(map(str, aci))}',
                  text, count=1, flags=re.M)
    write(dst, text, open(os.path.join(STYLES, base), "rb").read()[:48])
    assert read(dst) == text
    print("written:", dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], [int(a) for a in sys.argv[3:]])
