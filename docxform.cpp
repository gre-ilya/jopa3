// docxform - a tag-driven templater for .docx files (Qt5).
//
// The document is filled entirely from YOUR OWN tags (defined in tablekinds.cpp);
// there is no interactive per-variable form. Two kinds of tag are expanded:
//
//   Tables: a fixed tag like \tablewage is replaced by a generated <w:tbl>. Each
//   table has its own bare tag bound to a builder in tablekinds.cpp - that is
//   where you add new ones. Fixed tags need no user input; they are always
//   inserted automatically.
//
//   Text tags: a bare tag like \company is replaced INLINE by text a function in
//   tablekinds.cpp returns (fixedTexts()). Same idea as a fixed table, but it
//   substitutes text in place (keeping surrounding text/formatting) instead of
//   inserting a whole table.
//
// Workflow (GUI):
//   1. On launch the program asks for the .docx template to fill, then
//      immediately asks where to save the generated document.
//   2. It reads the template, expands every fixed table/text tag, and writes the
//      result to the chosen path. Whether the inserted text (and table contents)
//      is highlighted yellow is decided by a bool passed to fillTemplate().
//
// Tags split by the editor across several runs (a very common thing in Word) are
// handled: a paragraph's text is matched as a whole and only the matched span is
// rebuilt. Images, tables, tabs and other formatting are preserved; tags inside
// tables are expanded too.
//
// Build:  make docxform
//   or:   g++ -O2 -std=c++17 -fPIC docxform.cpp tablekinds.cpp -o docxform
//             $(pkg-config --cflags --libs Qt5Widgets) -lz
// Usage:  ./docxform            (asks for the template, then the output path)
//         ./docxform template.docx    (skips the chooser; asks only where to save)
//         ./docxform --render <in.docx> <out.docx> [--no-highlight]  (headless)

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <map>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include <zlib.h>

#include <QApplication>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QMessageBox>
#include <QWidget>

#include "docxform.h"
#include "tablekinds.h"

namespace {

// ---- File I/O -------------------------------------------------------------
//
// All file access goes through QFile (not std::fstream), so non-ASCII paths —
// e.g. Cyrillic file names — open correctly on every platform, including
// Windows, where the narrow std::fstream would use the ANSI code page. For
// command-line paths use QFile::decodeName(argv[i]) to get the QString.

bool readWholeFile(const QString& path, std::string& out) {
    QFile f(path);
    if (!f.open(QIODevice::ReadOnly)) return false;
    const QByteArray data = f.readAll();
    out.assign(data.constData(), static_cast<size_t>(data.size()));
    return true;
}

// Read a whole file straight from a QFile the caller already holds (so it can be
// filled from an existing handle without going through a path string). Opens it
// read-only if it is not open yet, reads from the beginning, and restores the
// original open/closed state afterwards.
bool readWholeQFile(QFile& f, std::string& out) {
    const bool wasOpen = f.isOpen();
    if (!wasOpen && !f.open(QIODevice::ReadOnly)) return false;
    if (!(f.openMode() & QIODevice::ReadOnly)) {  // opened, but not readable
        if (!wasOpen) f.close();
        return false;
    }
    if (!f.seek(0)) { if (!wasOpen) f.close(); return false; }
    const QByteArray data = f.readAll();
    out.assign(data.constData(), static_cast<size_t>(data.size()));
    if (!wasOpen) f.close();
    return true;
}

bool writeWholeFile(const QString& path, const std::string& bytes) {
    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) return false;
    const qint64 n = f.write(bytes.data(), static_cast<qint64>(bytes.size()));
    f.close();
    return n == static_cast<qint64>(bytes.size());
}

// ---- ZIP reader -----------------------------------------------------------

uint16_t rd16(const std::string& b, size_t p) {
    return static_cast<uint16_t>(static_cast<uint8_t>(b[p]) |
                                 (static_cast<uint8_t>(b[p + 1]) << 8));
}
uint32_t rd32(const std::string& b, size_t p) {
    return static_cast<uint32_t>(static_cast<uint8_t>(b[p])) |
           (static_cast<uint32_t>(static_cast<uint8_t>(b[p + 1])) << 8) |
           (static_cast<uint32_t>(static_cast<uint8_t>(b[p + 2])) << 16) |
           (static_cast<uint32_t>(static_cast<uint8_t>(b[p + 3])) << 24);
}

// Inflate a raw DEFLATE stream (no zlib/gzip wrapper).
bool inflateRaw(const char* data, size_t size, std::string& out) {
    z_stream zs{};
    if (inflateInit2(&zs, -MAX_WBITS) != Z_OK) return false;
    zs.next_in = reinterpret_cast<Bytef*>(const_cast<char*>(data));
    zs.avail_in = static_cast<uInt>(size);
    char buf[65536];
    int ret;
    do {
        zs.next_out = reinterpret_cast<Bytef*>(buf);
        zs.avail_out = sizeof(buf);
        ret = inflate(&zs, Z_NO_FLUSH);
        if (ret != Z_OK && ret != Z_STREAM_END && ret != Z_BUF_ERROR) {
            inflateEnd(&zs);
            return false;
        }
        out.append(buf, sizeof(buf) - zs.avail_out);
    } while (ret != Z_STREAM_END && zs.avail_out == 0);
    inflateEnd(&zs);
    return ret == Z_STREAM_END;
}

// Extract a single member from a ZIP archive held in memory.
bool extractZipMember(const std::string& zip, const std::string& member,
                      std::string& out) {
    const uint32_t kEocdSig = 0x06054b50;
    if (zip.size() < 22) return false;
    size_t eocd = std::string::npos;
    size_t lowest = zip.size() > 65557 ? zip.size() - 65557 : 0;
    for (size_t i = zip.size() - 22 + 1; i-- > lowest;)
        if (rd32(zip, i) == kEocdSig) { eocd = i; break; }
    if (eocd == std::string::npos) return false;

    uint16_t entries = rd16(zip, eocd + 10);
    uint32_t cdOffset = rd32(zip, eocd + 16);
    const uint32_t kCenSig = 0x02014b50;
    size_t p = cdOffset;
    for (uint16_t i = 0; i < entries; ++i) {
        if (p + 46 > zip.size() || rd32(zip, p) != kCenSig) return false;
        uint16_t method = rd16(zip, p + 10);
        uint32_t compSize = rd32(zip, p + 20);
        uint16_t nameLen = rd16(zip, p + 28);
        uint16_t extraLen = rd16(zip, p + 30);
        uint16_t commentLen = rd16(zip, p + 32);
        uint32_t localOffset = rd32(zip, p + 42);
        std::string name = zip.substr(p + 46, nameLen);
        if (name == member) {
            const uint32_t kLocSig = 0x04034b50;
            if (localOffset + 30 > zip.size() ||
                rd32(zip, localOffset) != kLocSig)
                return false;
            uint16_t lNameLen = rd16(zip, localOffset + 26);
            uint16_t lExtraLen = rd16(zip, localOffset + 28);
            size_t dataStart = localOffset + 30 + lNameLen + lExtraLen;
            if (dataStart + compSize > zip.size()) return false;
            const char* data = zip.data() + dataStart;
            if (method == 0) { out.assign(data, compSize); return true; }
            if (method == 8) return inflateRaw(data, compSize, out);
            return false;
        }
        p += 46 + nameLen + extraLen + commentLen;
    }
    return false;
}

// Enumerate every member of a ZIP archive as (name, uncompressed bytes).
bool listZipMembers(const std::string& zip,
                    std::vector<std::pair<std::string, std::string>>& out) {
    const uint32_t kEocdSig = 0x06054b50;
    if (zip.size() < 22) return false;
    size_t eocd = std::string::npos;
    size_t lowest = zip.size() > 65557 ? zip.size() - 65557 : 0;
    for (size_t i = zip.size() - 22 + 1; i-- > lowest;)
        if (rd32(zip, i) == kEocdSig) { eocd = i; break; }
    if (eocd == std::string::npos) return false;

    uint16_t entries = rd16(zip, eocd + 10);
    uint32_t cdOffset = rd32(zip, eocd + 16);
    const uint32_t kCenSig = 0x02014b50;
    size_t p = cdOffset;
    for (uint16_t i = 0; i < entries; ++i) {
        if (p + 46 > zip.size() || rd32(zip, p) != kCenSig) return false;
        uint16_t method = rd16(zip, p + 10);
        uint32_t compSize = rd32(zip, p + 20);
        uint16_t nameLen = rd16(zip, p + 28);
        uint16_t extraLen = rd16(zip, p + 30);
        uint16_t commentLen = rd16(zip, p + 32);
        uint32_t localOffset = rd32(zip, p + 42);
        std::string name = zip.substr(p + 46, nameLen);
        const uint32_t kLocSig = 0x04034b50;
        if (localOffset + 30 > zip.size() || rd32(zip, localOffset) != kLocSig)
            return false;
        uint16_t lNameLen = rd16(zip, localOffset + 26);
        uint16_t lExtraLen = rd16(zip, localOffset + 28);
        size_t dataStart = localOffset + 30 + lNameLen + lExtraLen;
        if (dataStart + compSize > zip.size()) return false;
        const char* data = zip.data() + dataStart;
        std::string content;
        if (method == 0) content.assign(data, compSize);
        else if (method == 8) { if (!inflateRaw(data, compSize, content)) return false; }
        else return false;
        out.emplace_back(std::move(name), std::move(content));
        p += 46 + nameLen + extraLen + commentLen;
    }
    return true;
}

// ---- ZIP writer (STORED) --------------------------------------------------

void put16(std::string& b, uint16_t v) {
    b += static_cast<char>(v & 0xff);
    b += static_cast<char>((v >> 8) & 0xff);
}
void put32(std::string& b, uint32_t v) {
    b += static_cast<char>(v & 0xff);
    b += static_cast<char>((v >> 8) & 0xff);
    b += static_cast<char>((v >> 16) & 0xff);
    b += static_cast<char>((v >> 24) & 0xff);
}

std::string buildZipStored(
    const std::vector<std::pair<std::string, std::string>>& members) {
    std::vector<uint32_t> crcs(members.size()), offsets(members.size());
    std::string out;
    for (size_t k = 0; k < members.size(); ++k) {
        const std::string& name = members[k].first;
        const std::string& data = members[k].second;
        crcs[k] = static_cast<uint32_t>(
            crc32(0, reinterpret_cast<const Bytef*>(data.data()),
                  static_cast<uInt>(data.size())));
        offsets[k] = static_cast<uint32_t>(out.size());
        put32(out, 0x04034b50);
        put16(out, 20); put16(out, 0); put16(out, 0);
        put16(out, 0); put16(out, 0);
        put32(out, crcs[k]);
        put32(out, static_cast<uint32_t>(data.size()));
        put32(out, static_cast<uint32_t>(data.size()));
        put16(out, static_cast<uint16_t>(name.size()));
        put16(out, 0);
        out += name;
        out += data;
    }
    uint32_t cdStart = static_cast<uint32_t>(out.size());
    for (size_t k = 0; k < members.size(); ++k) {
        const std::string& name = members[k].first;
        const std::string& data = members[k].second;
        put32(out, 0x02014b50);
        put16(out, 20); put16(out, 20); put16(out, 0); put16(out, 0);
        put16(out, 0); put16(out, 0);
        put32(out, crcs[k]);
        put32(out, static_cast<uint32_t>(data.size()));
        put32(out, static_cast<uint32_t>(data.size()));
        put16(out, static_cast<uint16_t>(name.size()));
        put16(out, 0); put16(out, 0); put16(out, 0); put16(out, 0);
        put32(out, 0);
        put32(out, offsets[k]);
        out += name;
    }
    uint32_t cdEnd = static_cast<uint32_t>(out.size());
    put32(out, 0x06054b50);
    put16(out, 0); put16(out, 0);
    put16(out, static_cast<uint16_t>(members.size()));
    put16(out, static_cast<uint16_t>(members.size()));
    put32(out, cdEnd - cdStart);
    put32(out, cdStart);
    put16(out, 0);
    return out;
}

// ---- XML helpers ----------------------------------------------------------

std::string xmlUnescape(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size();) {
        if (s[i] == '&') {
            if (s.compare(i, 5, "&amp;") == 0) { out += '&'; i += 5; continue; }
            if (s.compare(i, 4, "&lt;") == 0)  { out += '<'; i += 4; continue; }
            if (s.compare(i, 4, "&gt;") == 0)  { out += '>'; i += 4; continue; }
            if (s.compare(i, 6, "&quot;") == 0){ out += '"'; i += 6; continue; }
            if (s.compare(i, 6, "&apos;") == 0){ out += '\''; i += 6; continue; }
            if (s.compare(i, 2, "&#") == 0) {
                size_t end = s.find(';', i);
                if (end != std::string::npos) {
                    long code = 0;
                    bool ok = true;
                    if (s[i + 2] == 'x' || s[i + 2] == 'X') {
                        for (size_t k = i + 3; k < end; ++k) {
                            char c = s[k];
                            int d = (c >= '0' && c <= '9') ? c - '0'
                                  : (c >= 'a' && c <= 'f') ? c - 'a' + 10
                                  : (c >= 'A' && c <= 'F') ? c - 'A' + 10 : -1;
                            if (d < 0) { ok = false; break; }
                            code = code * 16 + d;
                        }
                    } else {
                        for (size_t k = i + 2; k < end; ++k) {
                            if (s[k] < '0' || s[k] > '9') { ok = false; break; }
                            code = code * 10 + (s[k] - '0');
                        }
                    }
                    if (ok) {
                        if (code < 0x80) {
                            out += static_cast<char>(code);
                        } else if (code < 0x800) {
                            out += static_cast<char>(0xC0 | (code >> 6));
                            out += static_cast<char>(0x80 | (code & 0x3F));
                        } else if (code < 0x10000) {
                            out += static_cast<char>(0xE0 | (code >> 12));
                            out += static_cast<char>(0x80 | ((code >> 6) & 0x3F));
                            out += static_cast<char>(0x80 | (code & 0x3F));
                        } else {
                            out += static_cast<char>(0xF0 | (code >> 18));
                            out += static_cast<char>(0x80 | ((code >> 12) & 0x3F));
                            out += static_cast<char>(0x80 | ((code >> 6) & 0x3F));
                            out += static_cast<char>(0x80 | (code & 0x3F));
                        }
                        i = end + 1;
                        continue;
                    }
                }
            }
        }
        out += s[i++];
    }
    return out;
}

std::string xmlEscape(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (char c : s) {
        switch (c) {
            case '&': out += "&amp;"; break;
            case '<': out += "&lt;"; break;
            case '>': out += "&gt;"; break;
            default:  out += c;       break;
        }
    }
    return out;
}

// Does the tag starting at `pos` (a '<') have local name `name`? (Ignores the
// namespace prefix, e.g. the "w:" in "w:p".)
bool tagIs(const std::string& xml, size_t pos, const char* name) {
    size_t i = pos + 1;
    if (i < xml.size() && xml[i] == '/') ++i;  // closing tag
    size_t colon = xml.find(':', i);
    size_t gt = xml.find_first_of(" \t\r\n/>", i);
    if (colon != std::string::npos && (gt == std::string::npos || colon < gt))
        i = colon + 1;
    size_t len = std::strlen(name);
    if (xml.compare(i, len, name) != 0) return false;
    if (i + len >= xml.size()) return false;
    char after = xml[i + len];
    return after == ' ' || after == '\t' || after == '\r' || after == '\n' ||
           after == '>' || after == '/';
}

// Find the next closing tag "</...local>" at or after `from`.
size_t findClose(const std::string& xml, size_t from, const char* local) {
    size_t p = from;
    while ((p = xml.find("</", p)) != std::string::npos) {
        if (tagIs(xml, p, local)) return p;
        p += 2;
    }
    return std::string::npos;
}

// Concatenate the visible text (<w:t> contents) of a single run body.
std::string runText(const std::string& run) {
    std::string out;
    size_t i = 0;
    while ((i = run.find('<', i)) != std::string::npos) {
        if (tagIs(run, i, "t") && run[i + 1] != '/') {
            size_t open = run.find('>', i);
            if (open == std::string::npos) break;
            if (run[open - 1] == '/') { i = open + 1; continue; }  // <w:t/>
            size_t close = findClose(run, open, "t");
            if (close == std::string::npos) break;
            out += xmlUnescape(run.substr(open + 1, close - open - 1));
            i = close;
            continue;
        }
        size_t e = run.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return out;
}

// Inner content of a run's <w:rPr> (without the wrapping tags), or "".
std::string rPrInner(const std::string& run) {
    size_t i = 0;
    while ((i = run.find('<', i)) != std::string::npos) {
        if (tagIs(run, i, "rPr") && run[i + 1] != '/') {
            size_t open = run.find('>', i);
            if (open == std::string::npos) return "";
            if (run[open - 1] == '/') return "";  // <w:rPr/>
            size_t close = findClose(run, open, "rPr");
            if (close == std::string::npos) return "";
            return run.substr(open + 1, close - open - 1);
        }
        size_t e = run.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return "";
}

// A "simple" run carries plain text only (rPr + w:t). Runs holding drawings,
// tabs, breaks, fields, etc. are "complex" and left untouched. Sets hasText.
bool isComplexRun(const std::string& run, bool& hasText) {
    static const char* kComplex[] = {
        "tab", "br", "cr", "ptab", "drawing", "pict", "object", "fldChar",
        "fldSimple", "instrText", "sym", "noBreakHyphen", "softHyphen",
        "footnoteReference", "endnoteReference", "ruby", nullptr};
    hasText = false;
    bool complex = false;
    size_t i = 0;
    while ((i = run.find('<', i)) != std::string::npos) {
        if (tagIs(run, i, "t") && run[i + 1] != '/') {
            hasText = true;
        } else {
            for (int k = 0; kComplex[k]; ++k)
                if (tagIs(run, i, kComplex[k])) { complex = true; break; }
        }
        size_t e = run.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return complex;
}

// ---- Tag scanning ---------------------------------------------------------

// Find the next FIXED table tag (e.g. "\tablewage") at or after `from`. These
// are bare tags registered in docxform::fixedTables(), each always mapped to one
// table kind. On success sets begin/end (range of the tag) and `tag` (the
// matched literal, which is used as the lookup key). When several tags would
// match, the earliest position wins (the longest tag on a tie). A match must end
// on a word boundary, so "\tablewage" does not fire inside "\tablewages".
bool nextFixedTable(const std::string& s, size_t from, size_t& begin,
                    size_t& end, std::string& tag) {
    size_t best = std::string::npos, bestEnd = 0;
    std::string bestTag;
    for (const docxform::FixedTable& ft : docxform::fixedTables()) {
        if (ft.tag.empty()) continue;
        size_t p = s.find(ft.tag, from);
        while (p != std::string::npos) {
            size_t after = p + ft.tag.size();
            char c = (after < s.size()) ? s[after] : '\0';
            bool boundary =
                !(std::isalnum(static_cast<unsigned char>(c)) || c == '_');
            if (boundary) {
                if (p < best || (p == best && ft.tag.size() > bestTag.size())) {
                    best = p;
                    bestEnd = after;
                    bestTag = ft.tag;
                }
                break;  // earliest occurrence of this tag is enough
            }
            p = s.find(ft.tag, p + 1);
        }
    }
    if (best == std::string::npos) return false;
    begin = best;
    end = bestEnd;
    tag = bestTag;
    return true;
}

// Find the next FIXED text tag (e.g. "\company") at or after `from`. These are
// bare tags registered in docxform::fixedTexts(), each replaced inline by the
// text its function returns. Same matching rules as nextFixedTable (earliest
// position, longest tag on a tie, must end on a word boundary).
bool nextFixedText(const std::string& s, size_t from, size_t& begin,
                   size_t& end, std::string& tag) {
    size_t best = std::string::npos, bestEnd = 0;
    std::string bestTag;
    for (const docxform::FixedText& ft : docxform::fixedTexts()) {
        if (ft.tag.empty()) continue;
        size_t p = s.find(ft.tag, from);
        while (p != std::string::npos) {
            size_t after = p + ft.tag.size();
            char c = (after < s.size()) ? s[after] : '\0';
            bool boundary =
                !(std::isalnum(static_cast<unsigned char>(c)) || c == '_');
            if (boundary) {
                if (p < best || (p == best && ft.tag.size() > bestTag.size())) {
                    best = p;
                    bestEnd = after;
                    bestTag = ft.tag;
                }
                break;
            }
            p = s.find(ft.tag, p + 1);
        }
    }
    if (best == std::string::npos) return false;
    begin = best;
    end = bestEnd;
    tag = bestTag;
    return true;
}

// The replacement text for a fixed text tag (runs its function), or "" if the
// tag is not registered.
std::string fixedTextValue(const std::string& tag) {
    for (const docxform::FixedText& ft : docxform::fixedTexts())
        if (ft.tag == tag && ft.build) return ft.build(tag);
    return std::string();
}

// Concatenated text of a paragraph's simple runs (in document order).
std::string concatParagraphText(const std::string& inner) {
    std::string P;
    size_t i = 0;
    while (i < inner.size()) {
        if (inner[i] == '<' && i + 1 < inner.size() && inner[i + 1] != '/' &&
            tagIs(inner, i, "r")) {
            size_t open = inner.find('>', i);
            if (open == std::string::npos) break;
            if (inner[open - 1] == '/') { i = open + 1; continue; }
            size_t close = findClose(inner, open, "r");
            if (close == std::string::npos) break;
            size_t closeGt = inner.find('>', close);
            if (closeGt == std::string::npos) break;
            std::string body = inner.substr(open + 1, close - open - 1);
            bool hasText = false;
            if (!isComplexRun(body, hasText) && hasText) P += runText(body);
            i = closeGt + 1;
            continue;
        }
        ++i;
    }
    return P;
}

// ---- Document transformation ----------------------------------------------

// Rebuild one paragraph's inner XML, replacing each fixed text tag (\company, …,
// matched across the paragraph's merged simple-run text) with the text its
// builder returns. When `highlight` is true the inserted text is highlighted
// yellow; otherwise it is inserted without any highlight.
std::string transformParagraph(const std::string& inner, bool highlight) {
    struct Tok {
        bool simple;
        std::string raw;   // verbatim XML (non-simple content)
        std::string rpr;   // simple run: rPr inner
        std::string text;  // simple run: decoded text
        size_t g;          // simple run: start offset in P
    };
    std::vector<Tok> toks;
    std::string verb, P;
    size_t i = 0;
    while (i < inner.size()) {
        if (inner[i] == '<' && i + 1 < inner.size() && inner[i + 1] != '/' &&
            tagIs(inner, i, "r")) {
            size_t open = inner.find('>', i);
            if (open == std::string::npos) { verb += inner.substr(i); break; }
            if (inner[open - 1] == '/') {  // <w:r/>
                verb += inner.substr(i, open + 1 - i);
                i = open + 1;
                continue;
            }
            size_t close = findClose(inner, open, "r");
            if (close == std::string::npos) { verb += inner.substr(i); break; }
            size_t closeGt = inner.find('>', close);
            if (closeGt == std::string::npos) { verb += inner.substr(i); break; }
            size_t runEnd = closeGt + 1;
            std::string full = inner.substr(i, runEnd - i);
            std::string body = inner.substr(open + 1, close - open - 1);
            bool hasText = false;
            if (!isComplexRun(body, hasText) && hasText) {
                if (!verb.empty()) {
                    toks.push_back({false, std::move(verb), "", "", 0});
                    verb.clear();
                }
                Tok t;
                t.simple = true;
                t.rpr = rPrInner(body);
                t.text = runText(body);
                t.g = P.size();
                P += t.text;
                toks.push_back(std::move(t));
            } else {
                verb += full;
            }
            i = runEnd;
            continue;
        }
        verb += inner[i++];
    }
    if (!verb.empty()) toks.push_back({false, std::move(verb), "", "", 0});

    // Resolve every fixed text tag (\company, …, code-defined) to its builder's
    // text. (Fixed TABLE tags are handled in transformDocument.) The `hl` flag is
    // honoured only when `highlight` is on.
    struct M { size_t b, e; std::string repl; bool hl; };
    std::vector<M> ms;
    {  // fixed text tags, e.g. \company -> its function's text
        size_t pos = 0, b, e;
        std::string tag;
        while (nextFixedText(P, pos, b, e, tag)) {
            ms.push_back({b, e, fixedTextValue(tag), true});  // honour highlight
            pos = e;
        }
    }
    if (ms.empty()) return inner;  // nothing to do: keep paragraph byte-for-byte

    // Order all matches by position and drop any overlap (keep the earliest), so
    // the two scans above combine into one sorted, non-overlapping list.
    std::sort(ms.begin(), ms.end(),
              [](const M& a, const M& b) { return a.b < b.b; });
    {
        std::vector<M> kept;
        size_t lastEnd = 0;
        for (M& m : ms)
            if (kept.empty() || m.b >= lastEnd) {
                lastEnd = m.e;
                kept.push_back(std::move(m));
            }
        ms.swap(kept);
    }

    std::vector<int> beginAt(P.size() + 1, -1), inAt(P.size() + 1, -1);
    for (size_t k = 0; k < ms.size(); ++k) {
        beginAt[ms[k].b] = static_cast<int>(k);
        for (size_t x = ms[k].b; x < ms[k].e; ++x) inAt[x] = static_cast<int>(k);
    }
    std::vector<size_t> nb(P.size() + 1);
    nb[P.size()] = P.size();
    for (size_t x = P.size(); x-- > 0;)
        nb[x] = (beginAt[x] >= 0) ? x : nb[x + 1];

    auto emitRun = [&](const std::string& rpr, const std::string& text,
                       bool hl) -> std::string {
        if (text.empty()) return std::string();
        std::string r;
        if (hl)
            r = "<w:rPr>" + rpr + "<w:highlight w:val=\"yellow\"/></w:rPr>";
        else if (!rpr.empty())
            r = "<w:rPr>" + rpr + "</w:rPr>";
        // Split the text on '\n' and put a <w:br/> between the pieces, so a
        // newline in a substituted value becomes a real line break in Word (a
        // raw '\n' inside a single <w:t> is otherwise ignored). Everything stays
        // in ONE run, keeping the run's formatting/highlight.
        std::string body;
        size_t start = 0;
        for (size_t i = 0; i <= text.size(); ++i) {
            if (i == text.size() || text[i] == '\n') {
                body += "<w:t xml:space=\"preserve\">" +
                        xmlEscape(text.substr(start, i - start)) + "</w:t>";
                if (i < text.size()) body += "<w:br/>";  // break between lines
                start = i + 1;
            }
        }
        return "<w:r>" + r + body + "</w:r>";
    };

    std::string out;
    for (const Tok& t : toks) {
        if (!t.simple) { out += t.raw; continue; }
        size_t g = t.g, len = t.text.size(), x = g;
        while (x < g + len) {
            if (beginAt[x] >= 0) {  // a placeholder starts here -> value run
                const M& m = ms[beginAt[x]];
                out += emitRun(t.rpr, m.repl, m.hl && highlight);
                x = m.e;  // skip the whole match (may extend past this run)
                continue;
            }
            if (inAt[x] >= 0) {  // tail of a placeholder that started earlier
                size_t to = std::min(ms[inAt[x]].e, g + len);
                x = (to > x) ? to : x + 1;
                continue;
            }
            size_t stop = std::min(g + len, nb[x + 1]);  // plain text run
            if (stop <= x) stop = x + 1;
            out += emitRun(t.rpr, t.text.substr(x - g, stop - x), false);
            x = stop;
        }
    }
    return out;
}

// Build every registered fixed-tag table into `out`, keyed by its tag (so e.g.
// out["\\tablewage"] holds the content for \tablewage) by running each tag's own
// builder. Both the GUI and the headless path call this, so fixed tables are
// always inserted automatically, with no user input.
void addFixedTables(std::map<std::string, docxform::TableData>& out) {
    for (const docxform::FixedTable& ft : docxform::fixedTables())
        if (ft.build) out[ft.tag] = ft.build(ft.tag);
}

// Walk document.xml and, for every <w:p>:
//   - if (at body level) it holds a fixed table tag (\tablewage, …), replace the
//     WHOLE paragraph with the generated <w:tbl>;
//   - otherwise expand its fixed text tags (\company, …).
// `tables` maps each fixed tag to its table content (see addFixedTables).
// When `highlight` is true the inserted text AND the generated table contents
// are highlighted yellow; when false everything is inserted without highlight.
std::string transformDocument(
    const std::string& xml,
    const std::map<std::string, docxform::TableData>& tables,
    bool highlight) {
    std::string out;
    size_t i = 0, copyFrom = 0;
    int tableDepth = 0;
    while ((i = xml.find('<', i)) != std::string::npos) {
        bool closing = (i + 1 < xml.size() && xml[i + 1] == '/');

        if (tagIs(xml, i, "tbl")) {  // track nesting so we only insert at body level
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            bool selfClose = (xml[e - 1] == '/');
            if (closing) { if (tableDepth > 0) --tableDepth; }
            else if (!selfClose) ++tableDepth;
            i = e + 1;
            continue;
        }

        if (!closing && tagIs(xml, i, "p")) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            if (xml[e - 1] == '/') { i = e + 1; continue; }  // empty <w:p/>
            size_t close = findClose(xml, e + 1, "p");
            if (close == std::string::npos) break;
            std::string inner = xml.substr(e + 1, close - e - 1);

            // Does this paragraph contain a fixed table tag (\tablewage, …)?
            // (body only) Each tag's content is looked up in `tables` by the tag.
            std::string tablesXml;
            if (tableDepth == 0 && !tables.empty()) {
                std::string P = concatParagraphText(inner);
                size_t pos = 0, fb, fe;
                std::string ftag;
                while (nextFixedTable(P, pos, fb, fe, ftag)) {
                    auto it = tables.find(ftag);
                    if (it != tables.end()) {
                        // Two <w:tbl> in a row must be separated by a paragraph
                        // (OOXML would otherwise merge them), but no trailing
                        // paragraph is added after the (last) table.
                        if (!tablesXml.empty()) tablesXml += "<w:p/>";
                        tablesXml += docxform::buildTableXml(it->second, highlight);
                    }
                    pos = fe;
                }
            }

            if (!tablesXml.empty()) {  // replace the whole <w:p>...</w:p>
                size_t closeGt = xml.find('>', close);
                if (closeGt == std::string::npos) break;
                size_t pEnd = closeGt + 1;
                out += xml.substr(copyFrom, i - copyFrom);  // up to <w:p ...>
                out += tablesXml;
                copyFrom = pEnd;
                i = pEnd;
                continue;
            }

            out += xml.substr(copyFrom, e + 1 - copyFrom);  // up to <w:p ...>
            out += transformParagraph(inner, highlight);
            copyFrom = close;
            i = close;
            continue;
        }

        size_t e = xml.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    out += xml.substr(copyFrom);
    return out;
}

// Collect, in first-seen order, every fixed tag the given scanner finds across
// the document. Used by --tags (with nextFixedTable / nextFixedText) so you can
// confirm a tag is recognised. These tags need no user input.
using TagScanner = bool (*)(const std::string&, size_t, size_t&, size_t&,
                            std::string&);
std::vector<std::string> collectTags(const std::string& xml, TagScanner scan) {
    std::vector<std::string> out;
    std::set<std::string> seen;
    size_t i = 0;
    while ((i = xml.find('<', i)) != std::string::npos) {
        if (i + 1 < xml.size() && xml[i + 1] != '/' && tagIs(xml, i, "p")) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            if (xml[e - 1] == '/') { i = e + 1; continue; }
            size_t close = findClose(xml, e + 1, "p");
            if (close == std::string::npos) break;
            std::string P = concatParagraphText(xml.substr(e + 1, close - e - 1));
            size_t pos = 0, b, en;
            std::string tag;
            while (scan(P, pos, b, en, tag)) {
                if (seen.insert(tag).second) out.push_back(tag);
                pos = en;
            }
            i = close;
            continue;
        }
        size_t e = xml.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return out;
}

// ---- Rendering core -------------------------------------------------------

// Transform a whole .docx held in memory (`zip`, the raw archive bytes) into the
// generated .docx bytes (`out`): expand every fixed table/text tag, highlighting
// the inserted text and table contents yellow when `highlight` is on. On failure
// returns false and, if `error` is non-null, stores a message. This is the one
// place the actual generation happens; every public entry point funnels here.
bool renderZipBytes(const std::string& zip, bool highlight, std::string& out,
                    QString* error) {
    auto fail = [&](const QString& msg) -> bool {
        if (error) *error = msg;
        return false;
    };

    std::string xml;
    if (!extractZipMember(zip, "word/document.xml", xml))
        return fail(QString::fromUtf8(
            "Это не похоже на .docx (нет word/document.xml)."));

    std::vector<std::pair<std::string, std::string>> members;
    if (!listZipMembers(zip, members))
        return fail(QString::fromUtf8("Не удалось разобрать исходный .docx."));

    // Fixed table tags (\tablewage, ...) are always inserted, no user input.
    std::map<std::string, docxform::TableData> tableData;
    addFixedTables(tableData);

    std::string newXml = transformDocument(xml, tableData, highlight);
    for (auto& m : members)
        if (m.first == "word/document.xml") m.second = newXml;

    out = buildZipStored(members);
    return true;
}

// ---- GUI ------------------------------------------------------------------

// Shared by the fillTemplate() overloads: given the template .docx bytes and a
// suggested output name, ask WHERE to save, render and report success/failure via
// message boxes. Returns true only if a file was written.
bool saveAndRenderZip(const std::string& zip, const QString& suggested,
                      bool highlight, QWidget* parent) {
    QString outPath = QFileDialog::getSaveFileName(
        parent, QString::fromUtf8("Сохранить документ"), suggested,
        QString::fromUtf8("Документ Word (*.docx)"));
    if (outPath.isEmpty()) return false;  // user cancelled the save dialog
    if (!outPath.endsWith(".docx", Qt::CaseInsensitive)) outPath += ".docx";

    std::string bytes;
    QString err;
    if (!renderZipBytes(zip, highlight, bytes, &err)) {
        QMessageBox::critical(parent, QString::fromUtf8("Ошибка"), err);
        return false;
    }
    if (!writeWholeFile(outPath, bytes)) {
        QMessageBox::critical(
            parent, QString::fromUtf8("Ошибка"),
            QString::fromUtf8("Не удалось записать файл:\n%1").arg(outPath));
        return false;
    }
    QMessageBox::information(
        parent, QString::fromUtf8("Готово"),
        QString::fromUtf8("Документ сохранён:\n%1").arg(outPath));
    return true;
}

}  // namespace

// ---- Public embedding API (see docxform.h) --------------------------------

namespace docxform {

bool renderTemplate(const QString& templatePath, const QString& outPath,
                    bool highlight, QString* error) {
    auto fail = [&](const QString& msg) -> bool {
        if (error) *error = msg;
        return false;
    };

    std::string zip;
    if (!readWholeFile(templatePath, zip))
        return fail(QString::fromUtf8("Не удалось открыть файл:\n%1")
                        .arg(templatePath));
    std::string bytes;
    if (!renderZipBytes(zip, highlight, bytes, error)) return false;
    if (!writeWholeFile(outPath, bytes))
        return fail(
            QString::fromUtf8("Не удалось записать файл:\n%1").arg(outPath));
    return true;
}

// Default output name "<base>_filled.docx" next to `templateName` (a path or
// file name; may be empty, in which case a generic name is used).
static QString suggestedOutput(const QString& templateName) {
    QFileInfo info(templateName);
    QString base = info.completeBaseName();
    if (base.isEmpty()) base = QString::fromUtf8("document");
    QString dir = info.absolutePath();
    return dir.isEmpty() ? base + "_filled.docx"
                         : QDir(dir).filePath(base + "_filled.docx");
}

bool fillTemplate(const QString& templatePath, bool highlight,
                  QWidget* parent) {
    if (templatePath.isEmpty()) return false;  // nothing to fill
    std::string zip;
    if (!readWholeFile(templatePath, zip)) {
        QMessageBox::critical(
            parent, QString::fromUtf8("Ошибка"),
            QString::fromUtf8("Не удалось открыть файл:\n%1").arg(templatePath));
        return false;
    }
    // Ask only where to save (the template is already known), then render.
    return saveAndRenderZip(zip, suggestedOutput(templatePath), highlight,
                            parent);
}

bool fillTemplate(QFile& templateFile, bool highlight, QWidget* parent) {
    std::string zip;
    if (!readWholeQFile(templateFile, zip)) {
        QMessageBox::critical(
            parent, QString::fromUtf8("Ошибка"),
            QString::fromUtf8("Не удалось прочитать шаблон .docx."));
        return false;
    }
    // Suggest a name from the QFile's own file name (empty for handles with no
    // path). Ask only where to save, then render.
    return saveAndRenderZip(zip, suggestedOutput(templateFile.fileName()),
                            highlight, parent);
}

bool fillTemplate(bool highlight, QWidget* parent) {
    QString templatePath = QFileDialog::getOpenFileName(
        parent, QString::fromUtf8("Выберите шаблон .docx"), QString(),
        QString::fromUtf8("Документ Word (*.docx)"));
    if (templatePath.isEmpty()) return false;  // user cancelled the chooser
    // Immediately ask where to save the generated document (no intermediate
    // window) and render — same as passing the path in directly.
    return fillTemplate(templatePath, highlight, parent);
}

}  // namespace docxform

// Headless: list every fixed tag the document uses, one per line. Useful for
// inspection/integration (confirming a tag is recognised).
//   docxform --tags <in.docx>
// Output (tab-separated):
//   FIXEDTABLE\t<tag>               (always-inserted table tag, e.g. \tablewage)
//   FIXEDTEXT\t<tag>                (inline text tag, e.g. \company)
int listTags(int argc, char** argv) {
    if (argc < 3) {
        std::fprintf(stderr, "Usage: %s --tags <in.docx>\n", argv[0]);
        return 1;
    }
    std::string zip, xml;
    if (!readWholeFile(QFile::decodeName(argv[2]), zip) ||
        !extractZipMember(zip, "word/document.xml", xml)) {
        std::fprintf(stderr, "Error: cannot read .docx '%s'\n", argv[2]);
        return 1;
    }
    // Fixed-tag tables/texts present in the document (always auto-inserted).
    for (const auto& tag : collectTags(xml, nextFixedTable))
        std::printf("FIXEDTABLE\t%s\n", tag.c_str());
    for (const auto& tag : collectTags(xml, nextFixedText))
        std::printf("FIXEDTEXT\t%s\n", tag.c_str());
    return 0;
}

// Headless mode (no GUI), handy for scripting and testing:
//   docxform --render <in.docx> <out.docx> [--no-highlight]
// Fixed table/text tags (\tablewage, \company, …) are always expanded
// automatically. By default the inserted text and table contents are highlighted
// yellow; pass --no-highlight to insert them without any highlight.
int renderHeadless(int argc, char** argv) {
    if (argc < 4) {
        std::fprintf(stderr,
                     "Usage: %s --render <in.docx> <out.docx> [--no-highlight]\n",
                     argv[0]);
        return 1;
    }
    bool highlight = true;  // default: highlight inserted text/tables yellow
    for (int i = 4; i < argc; ++i)
        if (std::strcmp(argv[i], "--no-highlight") == 0) highlight = false;

    QString err;
    if (!docxform::renderTemplate(QFile::decodeName(argv[2]),
                                  QFile::decodeName(argv[3]), highlight, &err)) {
        std::fprintf(stderr, "Error: %s\n", err.toUtf8().constData());
        return 1;
    }
    return 0;
}

// The standalone executable's entry point. Define DOCXFORM_NO_MAIN when reusing
// docxform.cpp as a library inside another program (which has its own main()).
//
// GUI:  ./docxform [template.docx] [--no-highlight]
//   Without a path: asks for the template, then the output path. With a
//   template.docx path: skips the template chooser and asks only where to save.
//   Inserted text and table contents are highlighted yellow unless
//   --no-highlight is given (the `highlight` bool passed to fillTemplate()).
#ifndef DOCXFORM_NO_MAIN
int main(int argc, char** argv) {
    if (argc >= 2 && std::strcmp(argv[1], "--render") == 0)
        return renderHeadless(argc, argv);
    if (argc >= 2 && std::strcmp(argv[1], "--tags") == 0)
        return listTags(argc, argv);

    QApplication app(argc, argv);

    bool highlight = true;  // highlight inserted text/tables yellow by default
    QString templatePath;   // a non-flag argument = template to fill directly
    for (int i = 1; i < argc; ++i) {
        if (std::strcmp(argv[i], "--no-highlight") == 0) highlight = false;
        else if (templatePath.isEmpty()) templatePath = QFile::decodeName(argv[i]);
    }

    if (templatePath.isEmpty())
        docxform::fillTemplate(highlight);  // open template -> save -> render
    else
        docxform::fillTemplate(templatePath, highlight);  // save -> render
    return 0;
}
#endif  // DOCXFORM_NO_MAIN
