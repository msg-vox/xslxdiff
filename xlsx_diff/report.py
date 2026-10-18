"""Self-contained, escaped HTML and machine-readable CSV output."""
import csv
from collections import Counter
from html import escape
from itertools import groupby

FIELDS = ("sheet", "reference", "change_type", "property", "old_value", "new_value")


def summary(changes):
    counts = Counter(c.change_type for c in changes)
    return (f"{len(changes)} changes: {counts['added']} added, "
            f"{counts['removed']} removed, {counts['modified']} modified "
            f"across {len({c.sheet for c in changes})} worksheets.")


def write_csv(path, changes):
    with open(path, "w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(c.as_dict() for c in changes)


def write_html(path, changes, old_path, new_path, options):
    esc = lambda value: escape(str(value), quote=True)
    with open(path, "w", encoding="utf-8") as stream:
        stream.write('''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel comparison report</title><style>
body{font:15px system-ui,sans-serif;margin:2rem;color:#18212b;background:#fff}
h1,h2{line-height:1.3}table{border-collapse:collapse;width:100%;margin-bottom:2rem}
th,td{border:1px solid #cbd2d9;padding:.55rem;text-align:left;vertical-align:top}
th{background:#edf0f3}td{white-space:pre-wrap;overflow-wrap:anywhere;max-width:36rem}
.added td{background:#e0f3e5}.removed td{background:#fce1e1}.modified td{background:#fff0c2}
nav a{margin-right:1rem}small{color:#45515e}.table-wrap{overflow-x:auto}
</style></head><body><h1>Excel comparison report</h1>''')
        stream.write(f"<p><b>Old:</b> {esc(old_path)}<br><b>New:</b> {esc(new_path)}</p>")
        stream.write(f"<p><strong>{esc(summary(changes))}</strong></p>")
        stream.write(f"<p>Options: {esc(', '.join(options) or 'default comparison')}</p>")
        stream.write("<p>Cells are compared by address. Inserting rows or columns can produce many cell differences. "
                     "Formulas are compared as expressions and are never evaluated. "
                     "Counts represent individual property changes, not distinct cells.</p>")
        stream.write("<p>Colours: <span class='added'>added (green)</span>, "
                     "<span class='removed'>removed (red)</span>, "
                     "<span class='modified'>modified (yellow)</span>. The Change column also labels each change.</p>")
        groups = [(name, list(items)) for name, items in groupby(changes, key=lambda c: c.sheet)]
        stream.write("<nav>" + " ".join(f'<a href="#sheet-{i}">{esc(name)}</a>'
                                      for i, (name, _) in enumerate(groups)) + "</nav>")
        if not changes:
            stream.write("<p>No differences found.</p>")
        for i, (name, items) in enumerate(groups):
            stream.write(f'<h2 id="sheet-{i}">{esc(name)} ({len(items)} changes)</h2>')
            stream.write('<div class="table-wrap"><table><thead><tr><th>Reference</th><th>Change</th>'
                         '<th>Property</th><th>Old value</th><th>New value</th></tr></thead><tbody>')
            for c in items:
                stream.write(f'<tr class="{c.change_type}">' + ''.join(
                    f"<td>{esc(v)}</td>" for v in (c.reference or '(worksheet)', c.change_type,
                                                   c.property, c.old_value, c.new_value)) + '</tr>')
            stream.write('</tbody></table></div>')
        stream.write('</body></html>\n')
