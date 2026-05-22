"""命令行进度条（优先 tqdm，无则文本进度）。"""

from __future__ import annotations

import sys
from typing import Iterable, Iterator, TypeVar

T = TypeVar("T")


def iter_progress(
    items: Iterable[T],
    desc: str = "处理中",
    total: int | None = None,
    unit: str = "项",
) -> Iterator[T]:
    if total is None:
        items_list = list(items)
        total = len(items_list)
        items = items_list

    try:
        from tqdm import tqdm

        yield from tqdm(
            items,
            desc=desc,
            total=total,
            unit=unit,
            file=sys.stdout,
            dynamic_ncols=True,
        )
        return
    except ImportError:
        pass

    n = total or 0
    i = 0
    for item in items:
        i += 1
        if n > 0:
            pct = int(i * 100 / n)
            bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
            print(f"\r{desc} [{bar}] {i}/{n} ({pct}%)", end="", flush=True)
        else:
            print(f"\r{desc} {i} {unit}", end="", flush=True)
        yield item
    if n > 0 or i > 0:
        print(flush=True)
