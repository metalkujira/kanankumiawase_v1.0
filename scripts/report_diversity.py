from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path

import openpyxl


def summarize(path: Path) -> None:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["対戦分散チェック"]

    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    data_rows = rows[1:]

    idx_name = header.index("ペア名")
    idx_unique = header.index("対戦グループ数")
    idx_level = header.index("レベル")

    overall = Counter()
    per_level: dict[str, Counter[int]] = defaultdict(Counter)

    for row in data_rows:
        name = row[idx_name]
        if not name:
            continue
        unique_groups = row[idx_unique]
        level = row[idx_level]

        overall[unique_groups] += 1
        per_level[level][unique_groups] += 1

    print("Overall unique-group distribution:")
    for unique_groups in sorted(overall):
        print(f"  {unique_groups} groups: {overall[unique_groups]}")

    print("\nBy level:")
    for level in sorted(per_level):
        print(f"  Level {level}:")
        for unique_groups in sorted(per_level[level]):
            count = per_level[level][unique_groups]
            print(f"    {unique_groups} groups: {count}")


if __name__ == "__main__":
    summarize(Path("graph_schedule_20251127_143529.xlsx"))
