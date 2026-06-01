"""Entry point: fetch Star Wars data from swapi.info and write an Excel workbook.

Usage:
    python main.py                         # writes ./star_wars.xlsx
    python main.py --output data/sw.xlsx   # custom path
    python main.py --verbose               # debug logging (see every API call)

Set a breakpoint anywhere in src/ and launch "Build Star Wars Spreadsheet" from
the VS Code Run & Debug panel to step through fetching and sheet generation.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from src.spreadsheet import build_workbook
from src.swapi_client import BASE_URL, SwapiClient

log = logging.getLogger("starwars")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a Star Wars Excel workbook from swapi.info.")
    parser.add_argument(
        "-o",
        "--output",
        default="star_wars.xlsx",
        help="Path to write the .xlsx file (default: star_wars.xlsx).",
    )
    parser.add_argument(
        "--base-url",
        default=BASE_URL,
        help=f"swapi.info API base URL (default: {BASE_URL}).",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable debug logging (logs every HTTP request).",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    log.info("Connecting to %s", args.base_url)
    client = SwapiClient(base_url=args.base_url)
    try:
        data = client.fetch_all()
    except Exception as exc:  # noqa: BLE001 - surface any network/parse failure clearly
        log.error("Failed to fetch data from swapi.info: %s", exc)
        return 1

    workbook = build_workbook(client, data)

    output_path = Path(args.output)
    if output_path.parent != Path("."):
        output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)

    log.info("Done. Wrote %s (%d sheets).", output_path.resolve(), len(workbook.sheetnames))
    return 0


if __name__ == "__main__":
    sys.exit(main())
