# -*- coding: utf-8 -*-
"""
Create stoxx_data.json.gz for upload limits (e.g. 25 MB).

The portfolio HTML tries stoxx_data.json.gz first, then falls back to stoxx_data.json.
gzip typically shrinks this dataset to a few MB without changing the JSON contents.

Usage:
  python compress_stoxx_data_for_web.py
  python compress_stoxx_data_for_web.py --input path/to/stoxx_data.json --output path/to/stoxx_data.json.gz
"""

from __future__ import annotations

import argparse
import gzip
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main() -> None:
    ap = argparse.ArgumentParser(description="Gzip stoxx_data.json for the ESG portfolio HTML tool.")
    ap.add_argument("--input", type=Path, default=HERE / "stoxx_data.json", help="Source JSON path")
    ap.add_argument(
        "--output",
        type=Path,
        default=HERE / "stoxx_data.json.gz",
        help="Output .gz path (upload this instead of the raw JSON when needed)",
    )
    ap.add_argument(
        "--level",
        type=int,
        default=9,
        choices=range(1, 10),
        metavar="1-9",
        help="gzip compression level (default: 9)",
    )
    args = ap.parse_args()
    inp: Path = args.input
    if not inp.is_file():
        raise SystemExit(f"Missing input file: {inp}")
    raw = inp.read_bytes()
    out_bytes = gzip.compress(raw, compresslevel=args.level)
    args.output.write_bytes(out_bytes)
    print(
        f"Wrote {args.output} ({len(out_bytes) / 1_048_576:.2f} MiB compressed) "
        f"from {inp} ({len(raw) / 1_048_576:.2f} MiB raw)"
    )


if __name__ == "__main__":
    main()
