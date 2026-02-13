import hashlib
from pathlib import Path


def quality(items) -> float:
    """Расчет качества."""

    if not items:
        return 0.0
    t = len(items)
    s = 0.0

    if t >= 10:
        s += 0.2
    elif t >= 3:
        s += 0.1

    pr = sum(1 for i in items if i.get("price", 0) > 0)
    if pr > t * 0.5:
        s += 0.3
    elif pr > t * 0.1:
        s += 0.15

    sk = sum(1 for i in items if i.get("sku"))
    if sk > t * 0.5:
        s += 0.2
    elif sk > t * 0.1:
        s += 0.1

    avg = sum(len(i.get("name", "")) for i in items) / t
    if avg > 20:
        s += 0.2
    elif avg > 10:
        s += 0.1

    units = sum(
        1 for i in items
        if i.get("unit") and i["unit"] not in ("", "шт")
    )
    if units > t * 0.3:
        s += 0.1

    return min(s, 1.0)


def hash_file(p: Path) -> str:
    """Хеширование файла."""

    h = hashlib.sha256()
    with open(p, "rb") as f:
        for ch in iter(lambda: f.read(8192), b""):
            h.update(ch)
    return h.hexdigest()
