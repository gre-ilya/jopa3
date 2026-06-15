#!/usr/bin/env python3
"""Self-checking test suite for docxform.

It builds a set of .docx templates that exercise the tricky parts of the engine
(text before/after a variable, several variables per paragraph, placeholders
split across runs, formatted runs, intervening Word noise such as proofErr /
bookmarks, repeated and adjacent variables, and \\table{...}), then drives the
docxform binary in its headless modes and checks the results.

Usage:
    python3 tests/run_tests.py [path-to-docxform]   # default: ./docxform

Exit code 0 means every check passed.
"""
import os, re, subprocess, sys, tempfile, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BIN = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "docxform")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def run(text):
    return f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def run_fmt(text, rpr):
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def para(*runs):
    return "<w:p>" + "".join(runs) + "</w:p>"


def make_docx(path, paras):
    body = "".join(paras)
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def docx_text_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


PASS, FAIL = 0, 0


def check(cond, msg):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  ok  ", msg)
    else:
        FAIL += 1
        print("  FAIL", msg)


def vars_of(docx):
    out = subprocess.run([BIN, "--vars", docx], capture_output=True, text=True)
    rows = []
    for line in out.stdout.splitlines():
        parts = line.split("\t")
        if parts and parts[0] in ("VAR", "TABLE"):
            rows.append(parts)
    return rows


def render(docx, *args):
    out = os.path.join(tempfile.gettempdir(), "docxform_test_out.docx")
    r = subprocess.run([BIN, "--render", docx, out, *args],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    return docx_text_xml(out)


# Build the templates next to this script so they can be opened in Word too.
def p(name):
    return os.path.join(HERE, name)


def build_all():
    make_docx(p("simple.docx"), [
        para(run("Уважаемый {{client_name}}, ваш заказ {{order_id}} имеет статус {{status}}.")),
    ])
    make_docx(p("split.docx"), [
        para(run("Привет, "), run("{{na"), run("me"), run("}}"), run("! Рады видеть.")),
    ])
    make_docx(p("multi.docx"), [
        para(run("Договор №123 (без переменных).")),
        para(run("{{greeting}}, {{client_name}}!")),
        para(run("Повторно: {{client_name}} — ваш номер {{order_id}}.")),
    ])
    make_docx(p("table.docx"), [
        para(run("Клиент: {{client_name}}. Ниже график:")),
        para(run("\\table{schedule}")),
    ])
    make_docx(p("edge.docx"), [
        para(run("{{ spaced }} в начале, и ещё {{tail}} в конце")),
    ])
    make_docx(p("formatted.docx"), [
        para(run("Договор с "), run_fmt("{{client_name}}", "<w:b/>"), run(" заключён.")),
    ])
    make_docx(p("repeat.docx"), [
        para(run("Привет, {{name}}! Ещё раз: {{name}}.")),
        para(run("{{a}}{{b}} склеены.")),
    ])
    make_docx(p("realword.docx"), [
        para('<w:pPr><w:jc w:val="both"/></w:pPr>',
             run("Уважаемый "), run("{{cli"),
             '<w:proofErr w:type="spellStart"/>',
             '<w:bookmarkStart w:id="0" w:name="bm"/>',
             run("ent_name}}"),
             '<w:bookmarkEnd w:id="0"/>',
             '<w:proofErr w:type="spellEnd"/>',
             run(", добрый день.")),
    ])


def main():
    if not os.path.exists(BIN):
        print(f"docxform binary not found at {BIN}; build it first (make docxform)")
        return 1
    build_all()

    print("simple.docx — text before/between/after, 3 vars in one paragraph")
    rows = vars_of(p("simple.docx"))
    names = [r[1] for r in rows]
    check(names == ["client_name", "order_id", "status"], f"vars order {names}")
    ctx = rows[0][2]
    check(ctx.startswith("Уважаемый {{client_name}}"), "context keeps text BEFORE the variable")
    check("имеет статус {{status}}." in ctx, "context keeps text AFTER the variable")
    xml = render(p("simple.docx"), "client_name=Иван", "order_id=A-1", "status=ОК")
    check("Уважаемый " in xml and "Иван" in xml and " имеет статус " in xml,
          "render keeps surrounding text and substitutes values")
    check(xml.count('w:highlight w:val="yellow"') == 3, "each value highlighted yellow")

    print("split.docx — placeholder split across runs")
    rows = vars_of(p("split.docx"))
    check([r[1] for r in rows] == ["name"], "split placeholder detected as one var")
    xml = render(p("split.docx"), "name=Мир")
    check("Мир" in xml and "{{" not in xml, "split placeholder substituted, no leftover braces")

    print("multi.docx — repeated var grouped under first paragraph")
    rows = vars_of(p("multi.docx"))
    check([r[1] for r in rows] == ["greeting", "client_name", "order_id"],
          "each var listed once, in first-seen order")
    check(rows[2][2].startswith("Повторно:"), "order_id context is its own paragraph")

    print("table.docx — \\table{...} placeholder")
    rows = vars_of(p("table.docx"))
    kinds = [r for r in rows if r[0] == "TABLE"]
    check(kinds and kinds[0][1] == "schedule", "table placeholder reported")
    xml = render(p("table.docx"), "client_name=ООО Заря", "+schedule=schedule")
    check("<w:tbl>" in xml, "selected table inserted as <w:tbl>")
    check("Ниже график:" in xml, "text in the paragraph before the table preserved")

    print("edge.docx — spaces in braces, var at paragraph start")
    rows = vars_of(p("edge.docx"))
    check([r[1] for r in rows] == ["spaced", "tail"], "spaced/edge vars detected")
    xml = render(p("edge.docx"), "spaced=НАЧАЛО", "tail=КОНЕЦ")
    check(xml.find("НАЧАЛО") < xml.find("КОНЕЦ"), "order preserved")

    print("formatted.docx — bold placeholder keeps its formatting")
    xml = render(p("formatted.docx"), "client_name=Заря")
    check(re.search(r"<w:b/>\s*<w:highlight", xml) is not None,
          "bold kept and highlight added")

    print("repeat.docx — repeated and adjacent placeholders")
    xml = render(p("repeat.docx"), "name=Анна", "a=X", "b=Y")
    check(xml.count("Анна") == 2, "both occurrences of {{name}} replaced")
    check("X" in xml and "Y" in xml, "adjacent {{a}}{{b}} both replaced")

    print("realword.docx — pPr + proofErr/bookmark between split halves")
    rows = vars_of(p("realword.docx"))
    check([r[1] for r in rows] == ["client_name"], "placeholder across Word noise detected")
    xml = render(p("realword.docx"), "client_name=Иван")
    check("Иван" in xml and "{{cli" not in xml, "substituted despite intervening tags")
    check("<w:bookmarkStart" in xml and '<w:jc w:val="both"/>' in xml,
          "bookmark and paragraph properties preserved")

    print(f"\n{PASS} passed, {FAIL} failed")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
