"""
    Write CSV file to EXCEL file
"""

from csv import reader
from argparse import ArgumentParser
from pathlib import Path


try:
    from openpyxl import Workbook
except ImportError:
    print("openpyxl not found. Try `pip install openpyxl`")
    exit()


DESCRIPTION = """
    Writing csv file to xls/xlsx excel file
"""
CSV_SUFFIXES = ".csv"
EXCEL_SUFFIXES = (".xls", ".xlsx")

necessary_log = print
optional_log = print


def read_and_write(inp: str, out: str, header: bool, append: bool) -> None:
    in_fp = Path(inp)
    out_fp = Path(out)
    mode = "a" if append else "w"

    if not in_fp.exists():
        raise FileNotFoundError(".csv file not found")
    if in_fp.suffix not in CSV_SUFFIXES:
        raise ValueError(f"{in_fp.name} is not a csv file")

    if not out_fp.exists():
        necessary_log(f"{out_fp.name} is not exist. Will be created")
    if not out_fp.suffix:
        optional_log(f"{out_fp.name} file without extension. Will implicitly add .xls")
        out_fp = Path(out + ".xls")
    if out_fp.suffix and out_fp.suffix not in EXCEL_SUFFIXES:
        raise ValueError(f"{in_fp.name} is not a excel file")

    csv_file = open(in_fp, "r")
    excel_file = open(out_fp, mode)

    wb = Workbook()
    ws = wb.active

    optional_log("Reading CSV...")
    csv_rows = list(reader(csv_file))[::] if header else list(reader(csv_file))[1::]
    rows_len = len(csv_rows)

    if rows_len == 0:
        csv_file.close()
        excel_file.close()
        raise ValueError(f"{in_fp.name} file is empty")

    for i, row in enumerate(csv_rows):
        optional_log(f"[{(i+1) / rows_len * 100:.0f}%] Writing {i+1}/{rows_len}...")
        for j, val in enumerate(row):
            ws.cell(row=i + 1, column=j + 1, value=val)

    optional_log("Done!")
    wb.save(out_fp)
    csv_file.close()
    excel_file.close()


def main() -> None:
    global optional_log
    global necessary_log

    parser = ArgumentParser(description=DESCRIPTION)

    parser.add_argument(
        "-f",
        "-from",
        "-file",
        "-csv",
        metavar="input .csv filename",
        required=True,
        help=".csv file path",
    )

    parser.add_argument(
        "-o",
        "-output",
        "-out",
        metavar="output file",
        required=True,
        help="output file path",
    )

    parser.add_argument(
        "-a",
        action="append",
        required=False,
        help="Append instead of rewrite",
    )

    parser.add_argument(
        "-debug",
        "-d",
        choices=["all", "necessary", "no"],
        default="necessary",
        help="Debug info (default: necessary)",
    )

    parser.add_argument(
        "-header",
        choices=["include", "no"],
        default="include",
        help="Include header(first row of csv file) (default: include)",
    )

    args = parser.parse_args()

    if args.debug == "necessary":
        optional_log = lambda x: None

    if args.debug == "no":
        optional_log, necessary_log = lambda x: None

    include_header = True if args.header == "include" else False

    try:
        read_and_write(args.f, args.o, include_header, args.a)
    except Exception as e:
        necessary_log(e)


if __name__ == "__main__":
    main()
