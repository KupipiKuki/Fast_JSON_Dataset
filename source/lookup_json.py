#!/usr/bin/env python3
"""Perform exact indexed lookups without loading a complete dataset."""
from __future__ import annotations
import argparse, json, re, unicodedata
from pathlib import Path
from typing import Any


def normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value))).strip().casefold()


def parse_filters(values: list[str]) -> dict[str, str]:
    filters: dict[str, str] = {}
    for value in values:
        if "=" not in value:
            raise SystemExit(f"Filter must be FIELD=VALUE: {value}")
        field, expected = value.split("=", 1)
        if not field.strip():
            raise SystemExit(f"Filter must be FIELD=VALUE: {value}")
        filters[field.strip()] = normalize(expected)
    return filters


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dataset", required=True)
    p.add_argument("--index", required=True)
    p.add_argument("--key", required=True)
    p.add_argument("--where", action="append", default=[], metavar="FIELD=VALUE",
                   help="Filter matches by an exact normalized data field; repeat as needed")
    p.add_argument("--limit", type=int, default=100)
    args = p.parse_args()
    root = Path(args.dataset)
    indexes = json.loads((root / "indexes.json").read_text(encoding="utf-8"))
    field = indexes.get("fields", {}).get(args.index)
    if field is None:
        raise SystemExit(f"Unknown index: {args.index}")
    ids = field.get(normalize(args.key), [])
    filters = parse_filters(args.where)
    by_shard: dict[str, list[tuple[str, int]]] = {}
    for rid in ids:
        loc = indexes["record_locations"][rid]
        by_shard.setdefault(loc["shard"], []).append((rid, loc["offset"]))
    found = []
    limit = max(0, args.limit)
    for shard in sorted(by_shard):
        records = json.loads((root / shard).read_text(encoding="utf-8"))
        for _, offset in by_shard[shard]:
            record = records[offset]
            if all(normalize(record["data"].get(name, "")) == expected
                   for name, expected in filters.items()):
                found.append(record)
                if len(found) >= limit:
                    break
        if len(found) >= limit:
            break
    print(json.dumps({"index": args.index, "key": args.key, "where": args.where,
                      "matches": len(found), "records": found},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
