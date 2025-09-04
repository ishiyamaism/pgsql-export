from __future__ import annotations

import argparse
import sys

from .db import connect_via_env_or_dsn
from .exporter import write_excel


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Export PostgreSQL schema metadata to Excel (.xlsx)")
    p.add_argument("--env-key", help="Key in .env containing database URL/DSN (e.g., INVENTORY_DB)")
    p.add_argument("--dsn", help="Explicit DSN/URI (overrides --env-key)")
    p.add_argument("--out", default="pg_metadata.xlsx", help="Output .xlsx path")
    p.add_argument("--with-counts", action="store_true", help="Add row_count to Tables (may be slow)")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        conn = connect_via_env_or_dsn(args.env_key, args.dsn)
    except Exception as e:
        print(f"Connection error: {e}", file=sys.stderr)
        return 2

    try:
        write_excel(conn, args.out, with_counts=args.with_counts)
    finally:
        conn.close()

    print(f"OK: {args.out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

