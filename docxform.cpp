// docxform - a Jinja-style GUI templater for .docx files (Qt5).
//
// Workflow:
//   1. In Word/Google Docs/LibreOffice you write a document and mark the parts
//      that should change with \var{...} placeholders named in plain ASCII.
//      Two forms are accepted:
//        \var{name}        - a variable called `name`; the field gets the
//                            default tooltip.
//        \var{name|info}   - same variable `name`, but `info` (free-form text
//                            after the first '|') becomes the field's tooltip
//                            instead of the default one.
//      e.g. "Dear \var{client_name}, your order \var{order_id|the order number}
//      is \var{status}."
//   2. docxform opens the .docx, finds every \var{...}, and shows a window with
//      one text field per variable (the field label is the variable name).
//   3. You type the values and press "Создать документ…".
//   4. docxform writes a new .docx where each \var{...} is replaced by the value
//      you entered, and every substituted value is highlighted yellow.
//
// Placeholders split by the editor across several runs (a very common thing in
// Word) are handled: a paragraph's text is matched as a whole and only the
// matched span is rebuilt. Images, tables, tabs and other formatting are
// preserved; placeholders inside tables are templated too.
//
// Tables: a fixed tag like \tablewage is replaced by a generated <w:tbl>. Each
// table has its own bare tag bound to a builder in tablekinds.cpp - that is
// where you add new ones. Fixed tags need no GUI control and no headless
// argument; they are always inserted automatically.
//
// Variants: a \variant{name|text 1|text 2|...} placeholder lets the user PICK
// one of several wordings. `name` may be non-ASCII and contain spaces; the
// options are the |-separated parts after it (taken verbatim). The GUI shows a
// drop-down per variant and substitutes the chosen text (highlighted yellow).
//
// The form lists variables and variants in the SAME order they appear in the
// document (interleaved, not grouped by type).
//
// Build:  make docxform
//   or:   g++ -O2 -std=c++17 -fPIC docxform.cpp tablekinds.cpp -o docxform
//             $(pkg-config --cflags --libs Qt5Widgets) -lz
// Usage:  ./docxform [template.docx]   (without an argument it asks for a file)

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
#include <QCheckBox>
#include <QComboBox>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QFormLayout>
#include <QLabel>
#include <QLineEdit>
#include <QMessageBox>
#include <QPushButton>
#include <QScrollArea>
#include <QVBoxLayout>
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

// ---- Placeholder scanning -------------------------------------------------

// Trim leading/trailing ASCII whitespace from a string (in place by value).
std::string trimAscii(const std::string& s) {
    size_t a = 0, b = s.size();
    while (a < b && std::isspace(static_cast<unsigned char>(s[a]))) ++a;
    while (b > a && std::isspace(static_cast<unsigned char>(s[b - 1]))) --b;
    return s.substr(a, b - a);
}

// Find the next \var{...} variable at or after `from`. Two forms are accepted:
//   \var{name}        -> name = name, info = ""    (default tooltip)
//   \var{name|info}   -> name = name, info = info  (custom tooltip text)
// `name` is the variable identifier [A-Za-z0-9_]+ (surrounding ASCII spaces
// trimmed); it is both the field label and the substitution key — a name filled
// once is substituted at every occurrence. `info`, when present, is everything
// after the FIRST '|' inside the braces, trimmed of ASCII whitespace — so it may
// itself contain '|', commas and other punctuation, as well as spaces and
// non-ASCII text (e.g. Russian). The '|' separator mirrors \variant{...}. On
// success sets begin/end (byte range of the whole \var{...}), name and info; a
// malformed \var{...} (empty/invalid name) is skipped. Note: "\var{" is not a
// prefix of "\variant{", so the two never collide.
bool nextVarPlaceholder(const std::string& s, size_t from, size_t& begin,
                        size_t& end, std::string& name, std::string& info) {
    static const std::string kMarker = "\\var{";
    size_t p = from;
    while ((p = s.find(kMarker, p)) != std::string::npos) {
        size_t close = s.find('}', p + kMarker.size());
        if (close == std::string::npos) return false;
        std::string inner =
            s.substr(p + kMarker.size(), close - (p + kMarker.size()));

        // Split the name (before the first '|') from the optional info
        // (everything after it). With no '|' the whole thing is the name.
        std::string namePart, infoPart;
        size_t bar = inner.find('|');
        if (bar == std::string::npos) {
            namePart = inner;
        } else {
            namePart = inner.substr(0, bar);
            infoPart = inner.substr(bar + 1);
        }

        std::string nm = trimAscii(namePart);
        bool validName = !nm.empty();
        for (char c : nm)
            if (!(std::isalnum(static_cast<unsigned char>(c)) || c == '_')) {
                validName = false;
                break;
            }
        if (validName) {
            begin = p;
            end = close + 1;
            name = nm;
            info = trimAscii(infoPart);
            return true;
        }
        p = close + 1;  // not a valid \var{...}; keep scanning
    }
    return false;
}

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

// Find the next \variant{name|opt1|opt2|...} at or after `from`. The name is the
// text before the first '|' (trimmed of ASCII whitespace; it may contain spaces
// and non-ASCII bytes, e.g. Russian). The options are the '|'-separated parts
// after the name, taken VERBATIM (not trimmed), so inline substitutions keep
// their exact spacing. A valid variant needs a non-empty name and >= 1 option.
bool nextVariantPlaceholder(const std::string& s, size_t from, size_t& begin,
                            size_t& end, std::string& name,
                            std::vector<std::string>& options) {
    static const std::string kMarker = "\\variant{";
    size_t p = from;
    while ((p = s.find(kMarker, p)) != std::string::npos) {
        size_t close = s.find('}', p + kMarker.size());
        if (close == std::string::npos) return false;
        std::string inner =
            s.substr(p + kMarker.size(), close - (p + kMarker.size()));
        std::vector<std::string> parts;
        size_t start = 0;
        for (;;) {
            size_t bar = inner.find('|', start);
            if (bar == std::string::npos) {
                parts.push_back(inner.substr(start));
                break;
            }
            parts.push_back(inner.substr(start, bar - start));
            start = bar + 1;
        }
        std::string nm = parts.empty() ? std::string() : parts.front();
        size_t a = 0, b = nm.size();  // trim ASCII whitespace around the name
        while (a < b && std::isspace(static_cast<unsigned char>(nm[a]))) ++a;
        while (b > a && std::isspace(static_cast<unsigned char>(nm[b - 1]))) --b;
        nm = nm.substr(a, b - a);
        if (!nm.empty() && parts.size() >= 2) {
            begin = p;
            end = close + 1;
            name = nm;
            options.assign(parts.begin() + 1, parts.end());
            return true;
        }
        p = close + 1;  // not a valid variant; keep scanning
    }
    return false;
}

// ---- Unified token model --------------------------------------------------
//
// A template has two kinds of fillable placeholders. To present them in the GUI
// (and list them headless) in DOCUMENT ORDER, we scan a paragraph for the
// earliest token of any kind and walk forward. (Fixed table tags need no input
// and are handled separately, at the paragraph level.)

enum class FieldKind { Var, Variant };

struct Token {
    FieldKind kind;
    size_t begin = 0;
    size_t end = 0;
    std::string name;
    std::string info;                  // Var only: custom tooltip ("" if none)
    std::vector<std::string> options;  // Variant only
};

// Find the earliest of \var{...} or \variant{...} at or after `from`. The two
// syntaxes never overlap, so picking the smallest begin is unambiguous. Returns
// false when no token remains.
bool nextAnyToken(const std::string& s, size_t from, Token& tok) {
    size_t best = std::string::npos;
    Token chosen;
    size_t b, e;
    std::string nm;

    {
        std::string inf;
        if (nextVarPlaceholder(s, from, b, e, nm, inf) && b < best) {
            best = b;
            chosen = {FieldKind::Var, b, e, nm, std::move(inf), {}};
        }
    }
    {
        std::vector<std::string> opts;
        std::string vn;
        if (nextVariantPlaceholder(s, from, b, e, vn, opts) && b < best) {
            best = b;
            chosen = {FieldKind::Variant, b, e, vn, "", std::move(opts)};
        }
    }

    if (best == std::string::npos) return false;
    tok = std::move(chosen);
    return true;
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

// Rebuild one paragraph's inner XML, replacing each \var{...} and each
// \variant{...} (matched across the paragraph's merged simple-run text) with its
// value, highlighted yellow. `values` maps a variable name to its text;
// `variantChoices` maps a variant name to the chosen option text (when a variant
// is not listed, its first option is used).
std::string transformParagraph(
    const std::string& inner,
    const std::map<std::string, std::string>& values,
    const std::map<std::string, std::string>& variantChoices,
    bool highlight) {
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

    // Resolve every \var{...} / \variant{...} token to its replacement text up
    // front. nextAnyToken yields tokens in increasing position, so `ms` stays
    // sorted by begin. (Fixed table tags are handled in transformDocument.)
    struct M { size_t b, e; std::string repl; bool hl; };
    std::vector<M> ms;
    {
        Token tok;
        size_t pos = 0;
        while (nextAnyToken(P, pos, tok)) {
            if (tok.kind == FieldKind::Var) {
                auto it = values.find(tok.name);
                if (it != values.end()) {
                    ms.push_back({tok.begin, tok.end, it->second, true});
                } else {
                    // Unfilled (e.g. headless without this NAME=): keep the
                    // original \var{...} literal in place, unhighlighted.
                    ms.push_back({tok.begin, tok.end,
                                  P.substr(tok.begin, tok.end - tok.begin),
                                  false});
                }
            } else if (tok.kind == FieldKind::Variant) {
                auto it = variantChoices.find(tok.name);
                std::string val =
                    (it != variantChoices.end())
                        ? it->second
                        : (tok.options.empty() ? std::string()
                                               : tok.options.front());
                ms.push_back({tok.begin, tok.end, std::move(val), true});
            }
            pos = tok.end;
        }
    }
    if (ms.empty()) return inner;  // nothing to do: keep paragraph byte-for-byte

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
        return "<w:r>" + r + "<w:t xml:space=\"preserve\">" + xmlEscape(text) +
               "</w:t></w:r>";
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
//   - otherwise substitute its \var{...} variables as usual.
// `tables` maps each fixed tag to its table content (see addFixedTables);
// `variantChoices` maps a \variant{...} name to its chosen option text.
// When `highlight` is true the substituted values are highlighted yellow (as
// before); when false they are inserted without any highlight.
std::string transformDocument(
    const std::string& xml,
    const std::map<std::string, std::string>& values,
    const std::map<std::string, std::string>& variantChoices,
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
                        tablesXml += docxform::buildTableXml(it->second);
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
            out += transformParagraph(inner, values, variantChoices, highlight);
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

// A fillable item discovered in a template: a \var{...} variable or a
// \variant{...} choice.
struct FormField {
    FieldKind kind;
    std::string name;
    std::string info;                  // Var: custom tooltip text ("" if none)
    std::vector<std::string> options;  // Variant: the available choices
};

// A paragraph that introduces one or more fillable items: the paragraph text
// (placeholder tokens kept, for context) and the NEW items it introduces, in
// document order. An item seen in an earlier paragraph is not repeated.
struct FormBlock {
    std::string text;
    std::vector<FormField> fields;
};

// For each \var name, the FIRST non-empty `info` seen anywhere in the document
// (in document order). Variables are deduplicated to one field, so a hint
// written on ANY occurrence (e.g. \var{DoxNum|Номер документа}) should apply to
// that field even when the variable's first occurrence was a bare \var{DoxNum}
// with no hint.
std::map<std::string, std::string> firstVarInfos(const std::string& xml) {
    std::map<std::string, std::string> infos;
    size_t i = 0;
    while ((i = xml.find('<', i)) != std::string::npos) {
        if (i + 1 < xml.size() && xml[i + 1] != '/' && tagIs(xml, i, "p")) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            if (xml[e - 1] == '/') { i = e + 1; continue; }
            size_t close = findClose(xml, e + 1, "p");
            if (close == std::string::npos) break;
            std::string P = concatParagraphText(xml.substr(e + 1, close - e - 1));
            Token tok;
            size_t pos = 0;
            while (nextAnyToken(P, pos, tok)) {
                if (tok.kind == FieldKind::Var && !tok.info.empty())
                    infos.emplace(tok.name, tok.info);  // keeps the first only
                pos = tok.end;
            }
            i = close;
            continue;
        }
        size_t e = xml.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return infos;
}

// Walk the document paragraph by paragraph and collect fillable items in the
// exact order they appear (variables and variants interleaved). Each item is
// listed once, under the paragraph of its first occurrence; paragraphs without
// any new item are skipped. This single ordered list drives both the GUI layout
// and the headless listing, so they always mirror the document.
std::vector<FormBlock> collectFormBlocks(const std::string& xml) {
    std::vector<FormBlock> blocks;
    // A \var hint may be written on any occurrence of the name; resolve each
    // name's hint up front so the (deduplicated) field gets it regardless of
    // which occurrence comes first.
    std::map<std::string, std::string> varInfos = firstVarInfos(xml);
    std::set<std::string> seenVar, seenVariant;
    size_t i = 0;
    while ((i = xml.find('<', i)) != std::string::npos) {
        if (i + 1 < xml.size() && xml[i + 1] != '/' && tagIs(xml, i, "p")) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            if (xml[e - 1] == '/') { i = e + 1; continue; }
            size_t close = findClose(xml, e + 1, "p");
            if (close == std::string::npos) break;
            std::string P = concatParagraphText(xml.substr(e + 1, close - e - 1));
            std::vector<FormField> fields;
            Token tok;
            size_t pos = 0;
            while (nextAnyToken(P, pos, tok)) {
                bool isNew = (tok.kind == FieldKind::Var)
                                 ? seenVar.insert(tok.name).second
                                 : seenVariant.insert(tok.name).second;
                if (isNew) {
                    std::string info;
                    if (tok.kind == FieldKind::Var) {
                        auto it = varInfos.find(tok.name);
                        if (it != varInfos.end()) info = it->second;
                    }
                    fields.push_back({tok.kind, tok.name, info, tok.options});
                }
                pos = tok.end;
            }
            if (!fields.empty())
                blocks.push_back({std::move(P), std::move(fields)});
            i = close;
            continue;
        }
        size_t e = xml.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return blocks;
}

// Collect the fixed-tag tables (\tablewage, ...) actually present in the
// document, in first-seen order. Used by --vars so you can confirm a tag is
// recognised. These need no user input — see addFixedTables.
std::vector<std::string> collectFixedTables(const std::string& xml) {
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
            while (nextFixedTable(P, pos, b, en, tag)) {
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

// ---- GUI ------------------------------------------------------------------

// A QWidget-only window (no Q_OBJECT, so no moc is needed): one control per
// fillable item — a text field for a \var{...} variable and a drop-down for a
// \variant{...} — laid out in document order, plus a button that writes the
// filled .docx. The widget is self-contained, so it can be embedded in another
// Qt application as-is.
class FormWindow : public QWidget {
public:
    FormWindow(const QString& path, std::string zip, std::string xml,
               std::vector<FormBlock> blocks)
        : path_(path),
          zip_(std::move(zip)),
          xml_(std::move(xml)),
          blocks_(std::move(blocks)) {
        setWindowTitle(QString::fromUtf8("docxform — шаблонизатор .docx"));
        resize(560, 540);

        int nVar = 0, nVariant = 0;
        for (const FormBlock& blk : blocks_)
            for (const FormField& f : blk.fields) {
                if (f.kind == FieldKind::Var) ++nVar;
                else ++nVariant;
            }

        auto* root = new QVBoxLayout(this);

        QString head =
            QString::fromUtf8("Шаблон: %1\nПеременных: %2,  вариантов: %3")
                .arg(QFileInfo(path_).fileName())
                .arg(nVar)
                .arg(nVariant);
        auto* header = new QLabel(head, this);
        header->setWordWrap(true);
        root->addWidget(header);

        auto* container = new QWidget;
        auto* form = new QFormLayout(container);
        form->setFieldGrowthPolicy(QFormLayout::AllNonFixedFieldsGrow);

        // One block per paragraph, in document order. The paragraph (its items
        // highlighted) is shown once, then a control per item in the order the
        // items appear in the paragraph — so the whole form mirrors the document.
        for (const FormBlock& blk : blocks_) {
            std::set<std::string> active;
            for (const FormField& f : blk.fields) active.insert(f.name);

            auto* ctx = new QLabel(contextHtml(blk.text, active), container);
            ctx->setTextFormat(Qt::RichText);
            ctx->setWordWrap(true);
            ctx->setStyleSheet("color:#555; margin-top:8px;");
            form->addRow(ctx);  // full-width context line

            for (const FormField& f : blk.fields) {
                QString qname = QString::fromStdString(f.name);
                if (f.kind == FieldKind::Var) {
                    auto* edit = new QLineEdit(container);
                    // Short prompt in the field (a long paragraph would be
                    // clipped); full guidance lives in the tooltip.
                    edit->setPlaceholderText(
                        QString::fromUtf8("значение для %1").arg(qname));
                    edit->setToolTip(varHintHtml(f.name, f.info));
                    form->addRow(qname, edit);
                    edits_.emplace_back(f.name, edit);
                } else {  // Variant
                    auto* combo = new QComboBox(container);
                    for (const std::string& opt : f.options)
                        combo->addItem(QString::fromStdString(opt));
                    combo->setToolTip(variantHintHtml(f.name));
                    form->addRow(qname, combo);
                    variantCtls_.push_back({f.name, combo, f.options});
                }
            }
        }

        if (blocks_.empty()) {
            form->addRow(new QLabel(
                QString::fromUtf8("В документе нет плейсхолдеров \\var{…} "
                                  "или \\variant{…}."),
                container));
        }

        auto* scroll = new QScrollArea(this);
        scroll->setWidgetResizable(true);
        scroll->setWidget(container);
        root->addWidget(scroll, 1);

        // Toggle for highlighting inserted values yellow in the output. On by
        // default (the previous, unconditional behaviour); unchecking it inserts
        // the values without any highlight.
        highlightCheck_ = new QCheckBox(
            QString::fromUtf8("Выделять вставленный текст жёлтым"), this);
        highlightCheck_->setChecked(true);
        root->addWidget(highlightCheck_);

        auto* generate =
            new QPushButton(QString::fromUtf8("Создать документ…"), this);
        generate->setDefault(true);
        root->addWidget(generate);
        connect(generate, &QPushButton::clicked, this, [this] { onGenerate(); });
    }

private:
    // Render a paragraph as rich text for a block's context line. Items in
    // `active` (those with a control right below this line) are highlighted; the
    // same placeholder introduced by an earlier block is greyed. Variants get a
    // ▾ marker.
    static QString contextHtml(const std::string& text,
                               const std::set<std::string>& active) {
        QString html;
        size_t cur = 0, pos = 0;
        Token tok;
        while (nextAnyToken(text, pos, tok)) {
            html += QString::fromStdString(text.substr(cur, tok.begin - cur))
                        .toHtmlEscaped();
            QString name = QString::fromStdString(tok.name).toHtmlEscaped();
            bool on = active.count(tok.name) != 0;
            QString shown = name;
            if (tok.kind == FieldKind::Variant)
                shown += QString::fromUtf8(" \xe2\x96\xbe");  // ▾
            if (on)
                html += "<span style=\"background-color:#fff3a0;"
                        "font-weight:bold;\">" + shown + "</span>";
            else if (tok.kind == FieldKind::Variant)
                html += "<span style=\"color:#999;\">" + shown + "</span>";
            else
                html += "<span style=\"color:#999;\">\\var{" + name + "}</span>";
            cur = tok.end;
            pos = tok.end;
        }
        html += QString::fromStdString(text.substr(cur)).toHtmlEscaped();
        return html;
    }

    // The general filling rules, shared by every variable field. Kept as one
    // paragraph so the help text and the per-field tooltip stay in sync.
    static QString fillingRules() {
        return QString::fromUtf8(
            "Введённое значение подставляется как есть во всех местах документа, "
            "где встречается этот плейсхолдер, поэтому обратите внимание на "
            "регистр, пробелы и знаки препинания. Если оставить поле пустым, "
            "плейсхолдер будет удалён из итогового файла.");
    }

    // Full guidance for one variable — used as the field's tooltip. When the
    // \var{...} carried an `info` part (\var{name|info}), that text REPLACES
    // the default tooltip, exactly as authored. Otherwise the default guidance
    // is shown. Wrapped in <span> so Qt always renders it as rich text (and
    // escaped entities display correctly).
    static QString varHintHtml(const std::string& name, const std::string& info) {
        if (!info.empty())
            return QString::fromUtf8("<span>%1</span>")
                .arg(QString::fromStdString(info).toHtmlEscaped());
        QString qname = QString::fromStdString(name).toHtmlEscaped();
        return QString::fromUtf8(
                   "<b>%1</b> — введите значение для переменной "
                   "<code>\\var{%1}</code>. %2")
            .arg(qname, fillingRules());
    }

    // Tooltip for a \variant{...} drop-down.
    static QString variantHintHtml(const std::string& name) {
        QString qname = QString::fromStdString(name).toHtmlEscaped();
        return QString::fromUtf8(
                   "<b>%1</b> — выберите один из вариантов текста. Выбранный "
                   "вариант будет подставлен в документ и выделен жёлтым.")
            .arg(qname);
    }

    void onGenerate() {
        std::map<std::string, std::string> values;
        for (const auto& kv : edits_)
            values[kv.first] = kv.second->text().toUtf8().toStdString();

        // The chosen text for each \variant{name} (the selected drop-down item).
        std::map<std::string, std::string> variantChoices;
        for (const auto& vc : variantCtls_) {
            int sel = vc.combo->currentIndex();
            if (sel >= 0 && sel < static_cast<int>(vc.options.size()))
                variantChoices[vc.name] = vc.options[sel];
        }

        // Fixed table tags (\tablewage, ...) are always inserted, no control.
        std::map<std::string, docxform::TableData> tableData;
        addFixedTables(tableData);

        QString suggested = QFileInfo(path_).absoluteDir().filePath(
            QFileInfo(path_).completeBaseName() + "_filled.docx");
        QString outPath = QFileDialog::getSaveFileName(
            this, QString::fromUtf8("Сохранить документ"), suggested,
            QString::fromUtf8("Документ Word (*.docx)"));
        if (outPath.isEmpty()) return;
        if (!outPath.endsWith(".docx", Qt::CaseInsensitive)) outPath += ".docx";

        std::vector<std::pair<std::string, std::string>> members;
        if (!listZipMembers(zip_, members)) {
            QMessageBox::critical(
                this, QString::fromUtf8("Ошибка"),
                QString::fromUtf8("Не удалось разобрать исходный .docx."));
            return;
        }
        std::string newXml = transformDocument(
            xml_, values, variantChoices, tableData,
            highlightCheck_->isChecked());
        for (auto& m : members)
            if (m.first == "word/document.xml") m.second = newXml;

        std::string bytes = buildZipStored(members);
        if (!writeWholeFile(outPath, bytes)) {
            QMessageBox::critical(
                this, QString::fromUtf8("Ошибка"),
                QString::fromUtf8("Не удалось записать файл:\n%1").arg(outPath));
            return;
        }

        QMessageBox::information(
            this, QString::fromUtf8("Готово"),
            QString::fromUtf8("Документ сохранён:\n%1").arg(outPath));
    }

    // A \variant{...} control plus the option texts backing its drop-down.
    struct VariantCtl {
        std::string name;
        QComboBox* combo;
        std::vector<std::string> options;
    };

    QString path_;
    std::string zip_;
    std::string xml_;
    std::vector<FormBlock> blocks_;  // fillable items in document order
    QCheckBox* highlightCheck_ = nullptr;  // highlight inserted text yellow?
    std::vector<std::pair<std::string, QLineEdit*>> edits_;  // \var{...}
    std::vector<VariantCtl> variantCtls_;                    // \variant{}
};

}  // namespace

// ---- Public embedding API (see docxform.h) --------------------------------

namespace docxform {

QWidget* openTemplateForm(const QString& templatePath, QString* error,
                          QWidget* parent) {
    auto fail = [&](const QString& msg) -> QWidget* {
        if (error) *error = msg;
        return nullptr;
    };

    std::string zip;
    if (!readWholeFile(templatePath, zip))
        return fail(QString::fromUtf8("Не удалось открыть файл:\n%1")
                        .arg(templatePath));
    std::string xml;
    if (!extractZipMember(zip, "word/document.xml", xml))
        return fail(QString::fromUtf8(
            "Это не похоже на .docx (нет word/document.xml)."));

    std::vector<FormBlock> blocks = collectFormBlocks(xml);
    auto* w = new FormWindow(templatePath, std::move(zip), std::move(xml),
                             std::move(blocks));
    if (parent) w->setParent(parent, w->windowFlags());
    return w;
}

QWidget* showTemplateForm(QWidget* parent) {
    QString path = QFileDialog::getOpenFileName(
        parent, QString::fromUtf8("Выберите шаблон .docx"), QString(),
        QString::fromUtf8("Документ Word (*.docx)"));
    if (path.isEmpty()) return nullptr;  // user cancelled the chooser

    QString err;
    QWidget* w = openTemplateForm(path, &err, parent);
    if (!w) {
        QMessageBox::critical(parent, QString::fromUtf8("Ошибка"), err);
        return nullptr;
    }
    w->setAttribute(Qt::WA_DeleteOnClose);  // tie lifetime to the window
    w->show();
    w->raise();
    w->activateWindow();
    return w;
}

}  // namespace docxform

// Headless: list every fillable item in DOCUMENT ORDER, one per line. Useful
// for inspection/integration.
//   docxform --vars <in.docx>
// Output (tab-separated):
//   VAR\t<name>\t<paragraph context>
//   VARIANT\t<name>\t<opt1>|<opt2>|...
//   FIXEDTABLE\t<tag>               (always-inserted fixed tag, e.g. \tablewage)
int listVars(int argc, char** argv) {
    if (argc < 3) {
        std::fprintf(stderr, "Usage: %s --vars <in.docx>\n", argv[0]);
        return 1;
    }
    std::string zip, xml;
    if (!readWholeFile(QFile::decodeName(argv[2]), zip) ||
        !extractZipMember(zip, "word/document.xml", xml)) {
        std::fprintf(stderr, "Error: cannot read .docx '%s'\n", argv[2]);
        return 1;
    }
    for (const auto& blk : collectFormBlocks(xml)) {
        for (const auto& f : blk.fields) {
            if (f.kind == FieldKind::Var) {
                std::printf("VAR\t%s\t%s\n", f.name.c_str(), blk.text.c_str());
            } else {  // Variant
                std::string joined;
                for (size_t k = 0; k < f.options.size(); ++k) {
                    if (k) joined += "|";
                    joined += f.options[k];
                }
                std::printf("VARIANT\t%s\t%s\n", f.name.c_str(), joined.c_str());
            }
        }
    }
    // Fixed-tag tables present in the document (always auto-inserted).
    for (const auto& tag : collectFixedTables(xml))
        std::printf("FIXEDTABLE\t%s\n", tag.c_str());
    return 0;
}

// Headless mode (no GUI), handy for scripting and testing:
//   docxform --render <in.docx> <out.docx>
//       NAME=VALUE ... [~VARIANT=choice ...] [--no-highlight]
// A "~name=choice" argument picks a \variant{name|...}: `choice` may be a 1-based
// option index or the option text verbatim. Variant names with spaces need
// quoting, e.g. "~обращение к клиенту=1". Fixed table tags (\tablewage, …) are
// always inserted automatically, no argument needed. By default inserted values
// are highlighted yellow; pass --no-highlight to insert them without highlight.
int renderHeadless(int argc, char** argv) {
    if (argc < 4) {
        std::fprintf(stderr,
                     "Usage: %s --render <in.docx> <out.docx> "
                     "[NAME=VALUE ...] [~VARIANT=choice ...] [--no-highlight]\n",
                     argv[0]);
        return 1;
    }
    std::string zip;
    if (!readWholeFile(QFile::decodeName(argv[2]), zip)) {
        std::fprintf(stderr, "Error: cannot read '%s'\n", argv[2]);
        return 1;
    }
    std::string xml;
    if (!extractZipMember(zip, "word/document.xml", xml)) {
        std::fprintf(stderr, "Error: '%s' is not a valid .docx\n", argv[2]);
        return 1;
    }
    std::map<std::string, std::string> values;
    std::map<std::string, std::string> variantChoices;
    std::map<std::string, docxform::TableData> tableData;

    // Variant options (name -> options), to resolve a numeric "~name=index".
    std::map<std::string, std::vector<std::string>> variantOptions;
    for (const auto& blk : collectFormBlocks(xml))
        for (const auto& f : blk.fields)
            if (f.kind == FieldKind::Variant)
                variantOptions.emplace(f.name, f.options);

    bool highlight = true;  // default: highlight inserted values yellow
    for (int i = 4; i < argc; ++i) {
        std::string a = argv[i];
        if (a == "--no-highlight") {  // insert values without yellow highlight
            highlight = false;
            continue;
        }
        size_t eq = a.find('=');
        if (eq == std::string::npos) continue;
        if (!a.empty() && a[0] == '~') {  // variant: ~name=index|text
            std::string name = a.substr(1, eq - 1);
            std::string val = a.substr(eq + 1);
            std::string chosen = val;  // default: treat as literal option text
            auto it = variantOptions.find(name);
            if (it != variantOptions.end() && !val.empty() &&
                val.find_first_not_of("0123456789") == std::string::npos) {
                size_t idx = 0;
                for (char c : val) idx = idx * 10 + static_cast<size_t>(c - '0');
                if (idx >= 1 && idx <= it->second.size())
                    chosen = it->second[idx - 1];
            }
            variantChoices[name] = chosen;
        } else {  // variable: name=value
            values[a.substr(0, eq)] = a.substr(eq + 1);
        }
    }
    // Fixed tags (\tablewage, ...) are always inserted, no argument needed.
    addFixedTables(tableData);

    std::vector<std::pair<std::string, std::string>> members;
    if (!listZipMembers(zip, members)) {
        std::fprintf(stderr, "Error: cannot parse '%s'\n", argv[2]);
        return 1;
    }
    std::string newXml =
        transformDocument(xml, values, variantChoices, tableData, highlight);
    for (auto& m : members)
        if (m.first == "word/document.xml") m.second = newXml;
    std::string bytes = buildZipStored(members);
    if (!writeWholeFile(QFile::decodeName(argv[3]), bytes)) {
        std::fprintf(stderr, "Error: cannot write '%s'\n", argv[3]);
        return 1;
    }
    return 0;
}

// The standalone executable's entry point. Define DOCXFORM_NO_MAIN when reusing
// docxform.cpp as a library inside another program (which has its own main()).
#ifndef DOCXFORM_NO_MAIN
int main(int argc, char** argv) {
    if (argc >= 2 && std::strcmp(argv[1], "--render") == 0)
        return renderHeadless(argc, argv);
    if (argc >= 2 && std::strcmp(argv[1], "--vars") == 0)
        return listVars(argc, argv);

    QApplication app(argc, argv);

    if (argc >= 2) {
        // A template path was given on the command line: open it directly.
        QString err;
        QWidget* w =
            docxform::openTemplateForm(QFile::decodeName(argv[1]), &err);
        if (!w) {
            QMessageBox::critical(nullptr, QString::fromUtf8("Ошибка"), err);
            return 1;
        }
        w->setAttribute(Qt::WA_DeleteOnClose);
        w->show();
    } else {
        // No file given: the same interactive flow exposed to embedders.
        if (!docxform::showTemplateForm()) return 0;  // cancelled or failed
    }
    return app.exec();
}
#endif  // DOCXFORM_NO_MAIN
