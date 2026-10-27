"""Helpers shared by the format handlers: addresses, run-aware text replacement, safe text checks."""

import copy
import re

from ..errors import Refusal

ORDINAL = re.compile(r"(\w+)\[(\d+)\]$")


def steps(address: str) -> list[tuple[str, int | str]]:
    """'body/table[1]/r[2]' -> [('body', None), ('table', 1), ('r', 2)]; 'shape[7]' keeps 7."""
    out = []
    for part in address.strip().strip("/").split("/"):
        match = ORDINAL.match(part)
        if match:
            out.append((match.group(1), int(match.group(2))))
        elif re.fullmatch(r"\w+\[[^\]]+\]", part):
            name, _, rest = part.partition("[")
            out.append((name, rest[:-1]))
        else:
            out.append((part, None))
    return out


def pick(items: list, ordinal, address: str, what: str):
    if not isinstance(ordinal, int) or ordinal < 1 or ordinal > len(items):
        raise Refusal("target_not_found", f"{address}: there is no {what} {ordinal} (there are {len(items)})",
                      hint="Use the addresses office_inspect or office_read returns")
    return items[ordinal - 1]


def expect(actual: str, expected, address: str) -> None:
    if expected is not None and actual != expected:
        raise Refusal("expected_mismatch", f"{address} does not hold the expected text",
                      hint="Read the target again and use its current text", actual=actual[:2000])


def excerpt(text: str, limit: int = 200) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"


def replace_in_runs(runs: list, find: str, replace: str, match_case: bool = True) -> int:
    """Replace every occurrence of `find` across a paragraph's runs, keeping each run's formatting.

    `runs` are objects with a writable `.text`. A match that spans runs is written into the run
    where it starts; the following runs lose only the matched part. Returns the number replaced.
    """
    texts = [r.text or "" for r in runs]
    whole = "".join(texts)
    haystack, needle = (whole, find) if match_case else (whole.lower(), find.lower())
    starts, i = [], haystack.find(needle)
    while needle and i >= 0:
        starts.append(i)
        i = haystack.find(needle, i + len(needle))
    if not starts:
        return 0
    bounds, offset = [], 0
    for t in texts:
        bounds.append((offset, offset + len(t)))
        offset += len(t)
    new = list(texts)
    for start in reversed(starts):
        end = start + len(find)
        first = next(k for k, (a, b) in enumerate(bounds) if a <= start < b or (a == b == start))
        for k in range(len(runs) - 1, first - 1, -1):
            a, b = bounds[k]
            if b <= start or a >= end:
                continue
            lo, hi = max(start, a) - a, min(end, b) - a
            piece = new[k]
            # Positions shift only inside this run, and later matches were already handled.
            new[k] = piece[:lo] + (replace if k == first else "") + piece[hi:]
    for r, t in zip(runs, new):
        if (r.text or "") != t:
            r.text = t
    return len(starts)


def clone(element):
    return copy.deepcopy(element)
