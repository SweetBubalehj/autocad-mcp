"""Where the texts of one dimension chain row go, so that no two overlap - pure geometry, no AutoCAD.

A row is a run of adjacent segments [a, b] along one axis, each with the width its text needs (text plus free
space on both sides). Extension lines are positions along that axis: the row's own (its segment ends) must
never be crossed by a text; those of the farther rows on the same side (they run through this row on their way
out) are avoided where there is room and crossed where a dense stretch leaves none.

- A text that fits its segment stays inside it, at the place nearest the middle that crosses no extension
  line, else in the middle.
- A text that does not fit (a 120 / 250 / 400 wall next to an opening) goes beside its segment: left, else
  right, else farther out. A fitting neighbour in the way slides aside within its own segment if it has room.
"""
from __future__ import annotations

EPS = 1e-6


def _crosses(lo: float, hi: float, lines) -> bool:
    return any(lo + EPS < x < hi - EPS for x in lines)


def layout_row(segments: list[tuple[float, float, float]], own_lines, strict_lines) -> list[tuple[float, bool]]:
    """segments: (a, b, width) with a < b; own_lines: the row's extension lines; strict_lines: own plus the
    farther rows'. Returns (text centre, moved off the middle) for each segment, in the order given."""
    segs = []
    for i, (a, b, w) in enumerate(segments):
        stops = [a] + [x for x in sorted(strict_lines) if a + EPS < x < b - EPS] + [b]
        mid, best = (a + b) / 2, None
        for g0, g1 in zip(stops, stops[1:]):
            if g1 - g0 >= w:
                c = min(max(mid, g0 + w / 2), g1 - w / 2)
                if best is None or abs(c - mid) < abs(best - mid):
                    best = c
        c = best if best is not None else mid
        segs.append({"i": i, "a": a, "b": b, "w": w, "fits": best is not None or b - a >= w, "c": c,
                     "moved": abs(c - mid) > EPS})
    placed = [s for s in segs if s["fits"]]

    def clashes(lo: float, hi: float, skip=()) -> list:
        return [s for s in placed if s not in skip and hi > s["c"] - s["w"] / 2 and lo < s["c"] + s["w"] / 2]

    def try_place(s: dict, c: float, lines) -> bool:
        """Put s at centre c if that crosses none of `lines` and every clashing text can slide clear."""
        lo, hi = c - s["w"] / 2, c + s["w"] / 2
        if _crosses(lo, hi, lines):
            return False
        slides = {}
        for o in clashes(lo, hi):
            if not o["fits"]:
                return False
            half = o["w"] / 2
            nc = hi + half if o["c"] >= c else lo - half                # push it away from s
            if not o["a"] + half - EPS <= nc <= o["b"] - half + EPS:
                return False
            if clashes(nc - half, nc + half, skip=(o,)) or _crosses(nc - half, nc + half, lines):
                return False
            slides[id(o)] = (o, nc)
        for o, nc in slides.values():
            o["c"], o["moved"] = nc, True
        s["c"], s["moved"] = c, True
        placed.append(s)
        return True

    for s in sorted(segs, key=lambda s: s["a"]):
        if s["fits"]:
            continue
        a, b, w = s["a"], s["b"], s["w"]
        options = [c for k in range(4) for c in (a - w / 2 - k * w / 2, b + w / 2 + k * w / 2)]
        if not any(try_place(s, c, lines) for lines in (strict_lines, own_lines) for c in options):
            s["c"], s["moved"] = b + w / 2, True                        # nowhere free: beside it on the right
            placed.append(s)
    return [(s["c"], s["moved"]) for s in sorted(segs, key=lambda s: s["i"])]


ROW_ORDER = {"openings": 0, "walls": 1, "overall": 2}   # arch chains, nearest the building first


def farther_lines(bounds: dict[str, set[float]], row: str) -> set[float]:
    """Extension lines of the rows on the same side as `row` ("bottom:openings") that lie farther out."""
    side, _, kind = row.partition(":")
    rank = ROW_ORDER.get(kind, len(ROW_ORDER))
    return set().union(*(b for r, b in bounds.items()
                         if r.partition(":")[0] == side and ROW_ORDER.get(r.partition(":")[2], -1) > rank))
