#!/bin/env python3
import sys
import argparse

from . import cue2pdf


def get_opt() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Convert BRM Cue to PDF')
    # parser.add_argument("population_csv_dir", action="store")
    # parser.add_argument("jigyosha_csv_file", action="store")
    # parser.add_argument("-o", "--opacity", action="store", type=float, help="0.0 - 1.0")
    # parser.add_argument("--topojson", action="store")

    args = parser.parse_args()

    return args


def main() -> int:
    args = get_opt()

    cue2pdf.convert()

    return 0


if __name__ == "__main__":
    sys.exit(main())