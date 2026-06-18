#!/usr/bin/env python3
# Generate a tiny .docx used by the CI smoke test. Usage: make_sample.py OUT.docx
import sys, zipfile

HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body>')
TAIL = '<w:sectPr/></w:body></w:document>'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
      'content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
      'package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.'
      'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="rId1" Type="http://schemas.'
        'openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>')


def para(text):
    return ('<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "sample.docx"
    body = para("Client: \\var{name}") + para("\\tablewage")
    doc = HEAD + body + TAIL
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
    print("wrote", out)


if __name__ == "__main__":
    main()
