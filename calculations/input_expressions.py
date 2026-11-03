"""Input-cell formulas (approval A-6).

In the v1.1 workbook any yellow input cell may contain an Excel formula, e.g. the worked example's
quantity cell ``=DB_N_Encl/2`` (enclosure count from the Design Basis).  This module reproduces that
behaviour: an input given as a string starting with ``=`` is evaluated with Excel arithmetic, where
``DB_<key>`` refers to the effective Design Basis value (blank -> 0, as Excel does for a reference to
an empty cell).  The original expression is preserved for traceability; nothing is replaced by a
plain number.

Supported grammar (a strict subset of Excel): numbers, DB_<key> names, + - * / ^, unary +/-,
parentheses, and the functions ROUND, ROUNDUP, ROUNDDOWN, INT, MAX, MIN, ABS.
Anything else raises UnsupportedExpression - it is reported, never silently substituted.
"""
from __future__ import annotations

import math
import re

from .excel_semantics import XLError, num, isnumber


class UnsupportedExpression(ValueError):
    pass


_TOKEN = re.compile(r"\s*(?:(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+)|(DB_[A-Za-z0-9_]+)|([A-Z]+)\s*\(|(.))")


def _tokenize(s):
    pos, out = 0, []
    while pos < len(s):
        m = _TOKEN.match(s, pos)
        if not m or m.end() == pos:
            break
        n, name, fn, op = m.groups()
        if n is not None:
            out.append(("num", float(n)))
        elif name is not None:
            out.append(("name", name))
        elif fn is not None:
            out.append(("fn", fn))
        elif op is not None and op.strip():
            out.append(("op", op))
        pos = m.end()
    return out


def _round(x, d, mode):
    f = 10 ** int(d)
    v = abs(x) * f
    if mode == "round":
        r = math.floor(v + 0.5)
    elif mode == "up":
        r = math.ceil(v - 1e-12)
    else:
        r = math.floor(v + 1e-12)
    return math.copysign(r / f, x)


FUNCS = {
    "ROUND": lambda a: _round(a[0], a[1], "round"),
    "ROUNDUP": lambda a: _round(a[0], a[1], "up"),
    "ROUNDDOWN": lambda a: _round(a[0], a[1], "down"),
    "INT": lambda a: math.floor(a[0]),
    "MAX": lambda a: max(a),
    "MIN": lambda a: min(a),
    "ABS": lambda a: abs(a[0]),
}


class _Parser:
    def __init__(self, toks, resolve):
        self.t, self.i, self.resolve, self.refs = toks, 0, resolve, {}

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expr(self):
        v = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            r = self.term()
            v = v + r if op == "+" else v - r
        return v

    def term(self):
        v = self.power()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            r = self.power()
            if op == "/" and r == 0:
                raise XLError("#DIV/0!")
            v = v * r if op == "*" else v / r
        return v

    def power(self):
        v = self.unary()
        while self.peek() == ("op", "^"):
            self.take()
            v = v ** self.unary()
        return v

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return -self.unary()
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        return self.atom()

    def atom(self):
        kind, val = self.take()
        if kind == "num":
            return val
        if kind == "name":
            raw = self.resolve(val)
            self.refs[val] = raw
            return num(raw)
        if kind == "fn":
            if val not in FUNCS:
                raise UnsupportedExpression(f"function {val}() is not supported in input cells")
            args = [self.expr()]
            while self.peek() == ("op", ","):
                self.take()
                args.append(self.expr())
            if self.take() != ("op", ")"):
                raise UnsupportedExpression("missing ')'")
            return FUNCS[val](args)
        if (kind, val) == ("op", "("):
            v = self.expr()
            if self.take() != ("op", ")"):
                raise UnsupportedExpression("missing ')'")
            return v
        raise UnsupportedExpression(f"unexpected token {val!r}")


def is_expression(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def evaluate(expr: str, resolve):
    """Evaluate an input-cell formula. Returns (value, references)."""
    body = expr[1:]
    toks = _tokenize(body)
    if "".join(str(t[1]) for t in toks).strip() == "" or not toks:
        raise UnsupportedExpression("empty formula")
    p = _Parser(toks, resolve)
    v = p.expr()
    if p.i != len(toks):
        raise UnsupportedExpression("unparsed trailing input")
    if isinstance(v, float) and v.is_integer():
        v = int(v) if abs(v) < 2 ** 53 else v
    return v, p.refs


def resolve_inputs(record: dict, db_resolve, trace: list, where: str):
    """Return a copy of `record` with every '=...' input evaluated. Unsupported formulas become an
    XLError('#UNSUPPORTED') value and are recorded in `trace` (never silently replaced)."""
    out = dict(record)
    for k, v in record.items():
        if is_expression(v):
            try:
                val, refs = evaluate(v, db_resolve)
                out[k] = val
                trace.append(dict(where=where, field=k, expression=v, value=val, references=refs, status="OK"))
            except XLError as e:
                out[k] = e
                trace.append(dict(where=where, field=k, expression=v, value=e.code, references={}, status="EXCEL ERROR"))
            except UnsupportedExpression as e:
                out[k] = XLError("#UNSUPPORTED")
                trace.append(dict(where=where, field=k, expression=v, value=None, references={}, status=f"UNSUPPORTED: {e}"))
    return out
