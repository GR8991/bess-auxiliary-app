"""Minimal, explicit reproduction of the Excel cell/function semantics used by the v1.1 workbook.

Value model
-----------
None   -> empty cell (never entered)
""     -> empty text (formula returning "")
int/float -> number (bool is NOT a number here)
str    -> text
XLError -> Excel error value (#VALUE!, #N/A, #DIV/0!)

Only the behaviours the workbook relies on are implemented; each is unit-tested in
tests/test_excel_semantics.py.  Text comparisons are case-insensitive (Excel default).
"""
from __future__ import annotations

import math
import re

SQRT3 = math.sqrt(3)  # workbook Name SQRT3 = SQRT(3)


class XLError(Exception):
    def __init__(self, code="#VALUE!"):
        super().__init__(code)
        self.code = code

    def __repr__(self):
        return self.code


def isnumber(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def isblank_text(v) -> bool:
    """True for empty cell or empty text (Excel  v=""  is TRUE for both)."""
    return v is None or v == ""


def N(v):
    """Excel N(): number -> number, text/blank -> 0, error -> propagate."""
    if isinstance(v, XLError):
        raise v
    return v if isnumber(v) else 0


def num(v):
    """Value used in arithmetic: blank -> 0, number -> number, numeric text -> number, other text -> #VALUE!."""
    if v is None:
        return 0
    if isnumber(v):
        return v
    if isinstance(v, XLError):
        raise v
    if isinstance(v, str):
        try:
            return float(v.strip()) if v.strip() != "" else _raise()
        except ValueError:
            raise XLError("#VALUE!")
    raise XLError("#VALUE!")


def _raise():
    raise XLError("#VALUE!")


def text(v) -> str:
    """Excel  v&""  conversion."""
    if v is None:
        return ""
    if isnumber(v):
        return num_to_text(v)
    return str(v)


def num_to_text(x) -> str:
    """Number -> text as Excel's General format (up to 15 significant digits)."""
    if isinstance(x, int) or (isinstance(x, float) and x.is_integer() and abs(x) < 1e15):
        return str(int(x))
    s = format(x, ".15g")
    if "e" in s:
        m, e = s.split("e")
        return f"{m}E{int(e):+03d}"
    return s


def eq(a, b) -> bool:
    """Excel '=' operator."""
    if isinstance(a, XLError):
        raise a
    if isinstance(b, XLError):
        raise b
    if a is None and b is None:
        return True
    if a is None:
        a = 0 if isnumber(b) else ""
    if b is None:
        b = 0 if isnumber(a) else ""
    if isnumber(a) and isnumber(b):
        return a == b
    if isinstance(a, str) and isinstance(b, str):
        return a.lower() == b.lower()
    return False


def ne(a, b) -> bool:
    return not eq(a, b)


def _rank(v):
    # Excel ordering: numbers < text < logicals
    return 0 if isnumber(v) else 1


def cmp(a, b) -> int:
    """Excel comparison for <, >, <=, >= ; returns -1/0/1."""
    if isinstance(a, XLError):
        raise a
    if isinstance(b, XLError):
        raise b
    if a is None:
        a = 0 if isnumber(b) or b is None else ""
    if b is None:
        b = 0 if isnumber(a) else ""
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 0:
        return (a > b) - (a < b)
    al, bl = a.lower(), b.lower()
    return (al > bl) - (al < bl)


def gt(a, b): return cmp(a, b) > 0
def ge(a, b): return cmp(a, b) >= 0
def lt(a, b): return cmp(a, b) < 0
def le(a, b): return cmp(a, b) <= 0


def left(v, n) -> str:
    return text(v)[:n]


# ------------------------------------------------------------------ criteria (COUNTIF / SUMIFS)
def _wild_to_re(p: str):
    out = []
    i = 0
    while i < len(p):
        ch = p[i]
        if ch == "~" and i + 1 < len(p) and p[i + 1] in "*?~":
            out.append(re.escape(p[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("^" + "".join(out) + "$", re.IGNORECASE | re.DOTALL)


def _parse_number(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


class Criterion:
    """Compiled Excel criterion (as used in COUNTIF/COUNTIFS/SUMIFS)."""

    def __init__(self, crit):
        self.raw = crit
        if isnumber(crit):
            self.op, self.val, self.isnum = "=", crit, True
            return
        s = "" if crit is None else str(crit)
        op = "="
        for o in ("<=", ">=", "<>", "<", ">", "="):
            if s.startswith(o):
                op, s = o, s[len(o):]
                break
        self.op = op
        n = _parse_number(s) if s.strip() != "" else None
        self.isnum = n is not None
        self.val = n if self.isnum else s
        self.rx = _wild_to_re(s) if (not self.isnum and op in ("=", "<>") and any(c in s for c in "*?~")) else None

    def match(self, v) -> bool:
        op, val = self.op, self.val
        if isinstance(v, XLError):
            return False
        if self.isnum:
            if op == "=":
                return (isnumber(v) and v == val) or (isinstance(v, str) and _parse_number(v) == val)
            if op == "<>":
                return not ((isnumber(v) and v == val) or (isinstance(v, str) and _parse_number(v) == val))
            if not isnumber(v):
                return False
            return {"<": v < val, ">": v > val, "<=": v <= val, ">=": v >= val}[op]
        # text criterion
        if op == "=":
            if val == "":
                return v is None or v == ""
            if not isinstance(v, str):
                return False
            return bool(self.rx.match(v)) if self.rx else v.lower() == val.lower()
        if op == "<>":
            if val == "":
                return not (v is None or v == "")
            if not isinstance(v, str):
                return True
            return not (bool(self.rx.match(v)) if self.rx else v.lower() == val.lower())
        if not isinstance(v, str):
            return False
        c = (v.lower() > val.lower()) - (v.lower() < val.lower())
        return {"<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]


def countifs(*pairs) -> int:
    rngs = pairs[0::2]
    crits = [Criterion(c) for c in pairs[1::2]]
    n = len(rngs[0])
    return sum(1 for i in range(n) if all(cr.match(r[i]) for r, cr in zip(rngs, crits)))


def countif(rng, crit) -> int:
    return countifs(rng, crit)


def sumifs(sum_rng, *pairs) -> float:
    rngs = pairs[0::2]
    crits = [Criterion(c) for c in pairs[1::2]]
    tot = 0
    for i, v in enumerate(sum_rng):
        if all(cr.match(r[i]) for r, cr in zip(rngs, crits)) and isnumber(v):
            tot += v
    return tot


def match_exact(x, rng):
    """MATCH(x, rng, 0) -> 1-based index or None (#N/A)."""
    if x is None or x == "":
        return None
    if isnumber(x):
        for i, v in enumerate(rng):
            if isnumber(v) and v == x:
                return i + 1
        return None
    s = str(x)
    rx = _wild_to_re(s) if any(c in s for c in "*?~") else None
    for i, v in enumerate(rng):
        if isinstance(v, str) and (rx.match(v) if rx else v.lower() == s.lower()):
            return i + 1
    return None


def xmax(vals):
    """MAX over numbers only (text/blank ignored); MAX of nothing = 0."""
    ns = [v for v in vals if isnumber(v)]
    return max(ns) if ns else 0


def xmin(vals):
    ns = [v for v in vals if isnumber(v)]
    return min(ns) if ns else 0


def count(vals) -> int:
    return sum(1 for v in vals if isnumber(v))


def small(vals, k):
    ns = sorted(v for v in vals if isnumber(v))
    if k < 1 or k > len(ns):
        raise XLError("#NUM!")
    return ns[k - 1]


def rank_desc(x, vals) -> int:
    return 1 + sum(1 for v in vals if isnumber(v) and v > x)


def safe(fn, *a):
    """Evaluate a cell formula; Excel errors become their error code string."""
    try:
        return fn(*a)
    except XLError as e:
        return e
    except ZeroDivisionError:
        return XLError("#DIV/0!")
