#!/usr/bin/env python3
# Check the rendered .docx from the CI smoke test. Usage: check_render.py OUT.docx
import sys, re, html, zipfile

# Force UTF-8 output so printing Cyrillic does not crash on a cp1252 console
# (e.g. Python on Windows/MSYS2).
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

xml = zipfile.ZipFile(sys.argv[1]).read("word/document.xml").decode("utf-8")
# Visible text = all <w:t> contents (the regex avoids matching <w:tbl>, <w:tcPr>…).
vis = html.unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", xml, re.S)))
print("visible:", vis)

assert "Client: Test" in vis, "variable not substituted: %r" % vis
assert "Сотрудник" in vis, "fixed table \\tablewage not inserted: %r" % vis
assert "<w:tbl" in xml, "no table element generated"
print("smoke OK")
