"""Command-line interface for xlsx-diff."""
import argparse
import sys
from pathlib import Path

from .compare import compare_workbooks
from .report import summary, write_csv, write_html


def parser():
    result = argparse.ArgumentParser(
        prog="xlsx-diff", description="Compare .xlsx files by cell address without evaluating formulas.",
        epilog="Inserting rows or columns can produce many reported cell differences.")
    result.add_argument("old", type=Path, help="original .xlsx file")
    result.add_argument("new", type=Path, help="updated .xlsx file")
    result.add_argument("--ignore-formatting", action="store_true",
                        help="ignore fonts, fills, borders, alignment and number formats")
    result.add_argument("--ignore-blank", action="store_true",
                        help="skip cell changes where both values are blank (None or empty string)")
    result.add_argument("--values-only", action="store_true",
                        help="compare cell values/formula expressions and worksheet presence/order only")
    result.add_argument("--output", type=Path, default=Path("report.html"), metavar="REPORT.html",
                        help="HTML destination (default: report.html); CSV uses the same stem")
    return result


def main(argv=None):
    args = parser().parse_args(argv)
    html_path = args.output
    csv_path = html_path.with_suffix(".csv")
    try:
        for path in (args.old, args.new):
            if path.suffix.lower() != ".xlsx":
                raise ValueError(f"Expected an .xlsx input: {path}")
            if not path.is_file():
                raise ValueError(f"Input file does not exist or is not a file: {path}")
        inputs = (args.old, args.new)
        outputs = (html_path, csv_path)
        if html_path.resolve() == csv_path.resolve():
            raise ValueError("HTML and CSV output paths must be different; use --output report.html")
        for output in outputs:
            for source in inputs:
                if output.resolve() == source.resolve() or (output.exists() and output.samefile(source)):
                    raise ValueError(f"Output would overwrite an input file: {output}")
        if html_path.exists() and csv_path.exists() and html_path.samefile(csv_path):
            raise ValueError("HTML and CSV outputs must not refer to the same file")
        changes = compare_workbooks(args.old, args.new, ignore_formatting=args.ignore_formatting,
                                    ignore_blank=args.ignore_blank, values_only=args.values_only)
        options = [key.replace('_', '-') for key in ('ignore_formatting', 'ignore_blank', 'values_only')
                   if getattr(args, key)]
        write_html(html_path, changes, args.old, args.new, options)
        write_csv(csv_path, changes)
    except Exception as exc:
        print(f"xlsx-diff: {exc}", file=sys.stderr)
        return 2
    print(summary(changes))
    print(f"HTML: {html_path}\nCSV:  {csv_path}")
    return 0
