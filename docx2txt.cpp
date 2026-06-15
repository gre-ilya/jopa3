// docx2txt - extract paragraph text from a .docx file into a .txt file.
//
// Reads a Word/Google-Docs/LibreOffice .docx (a ZIP archive) and writes a .txt
// with two sections:
//
//   ========== VARS ==========
//   VAR1=<text that was highlighted yellow>
//   VAR2=...
//
//   ========== TEXT ==========
//   ... the document's paragraph text, where every yellow-highlighted span is
//       replaced by a \var{VARn} placeholder ...
//
// Rules for the TEXT section: contents of tables are ignored, paragraphs are
// separated by a single blank line (LaTeX-style), and every output line is
// word-wrapped to at most 75 characters.
//
// Alongside the .txt it also writes "<input-stem>_template.docx": a copy of the
// input document in which every highlighted span is replaced, in place, by a
// {{VARn}} placeholder that keeps the original yellow highlight.
//
// Build:  g++ -O2 -std=c++17 -o docx2txt docx2txt.cpp -lz -lstdc++fs
//         (-lstdc++fs only needed on GCC 8 and older)
// Usage:  ./docx2txt <input.docx> <output.txt>

#include <cctype>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include <zlib.h>

namespace {

constexpr size_t kMaxLineLength = 75;

// ---- File I/O -------------------------------------------------------------

bool readWholeFile(const std::string& path, std::string& out) {
    std::ifstream in(path, std::ios::binary);
    if (!in) return false;
    out.assign(std::istreambuf_iterator<char>(in),
               std::istreambuf_iterator<char>());
    return true;
}

// ---- Minimal ZIP reader ---------------------------------------------------
//
// We only need to locate a single member ("word/document.xml") and inflate it,
// so we parse just enough of the ZIP format: the End Of Central Directory
// record, the central directory entries, and the per-member local header.
// Sizes/method come from the central directory, so archives that use data
// descriptors (e.g. those written by LibreOffice) are handled correctly.

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
    for (size_t i = zip.size() - 22 + 1; i-- > lowest;) {
        if (rd32(zip, i) == kEocdSig) {
            eocd = i;
            break;
        }
    }
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
            if (method == 0) {  // stored
                out.assign(data, compSize);
                return true;
            }
            if (method == 8) {  // deflate
                return inflateRaw(data, compSize, out);
            }
            return false;  // unsupported compression method
        }
        p += 46 + nameLen + extraLen + commentLen;
    }
    return false;
}

// Enumerate every member of a ZIP archive as (name, uncompressed bytes).
// Used to copy an input .docx through while swapping a single part.
bool listZipMembers(const std::string& zip,
                    std::vector<std::pair<std::string, std::string>>& out) {
    const uint32_t kEocdSig = 0x06054b50;
    if (zip.size() < 22) return false;
    size_t eocd = std::string::npos;
    size_t lowest = zip.size() > 65557 ? zip.size() - 65557 : 0;
    for (size_t i = zip.size() - 22 + 1; i-- > lowest;) {
        if (rd32(zip, i) == kEocdSig) { eocd = i; break; }
    }
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
        if (method == 0) {
            content.assign(data, compSize);
        } else if (method == 8) {
            if (!inflateRaw(data, compSize, content)) return false;
        } else {
            return false;  // unsupported compression method
        }
        out.emplace_back(std::move(name), std::move(content));
        p += 46 + nameLen + extraLen + commentLen;
    }
    return true;
}

// ---- Minimal ZIP writer ---------------------------------------------------
//
// Members are written STORED (uncompressed) — valid and openable by Word, and
// simple enough to need no deflate on the write side. (Same approach as the
// txt2docx writer.)

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
        put32(out, 0x04034b50);  // local header signature
        put16(out, 20);          // version needed
        put16(out, 0);           // flags
        put16(out, 0);           // method: stored
        put16(out, 0);           // mod time
        put16(out, 0);           // mod date
        put32(out, crcs[k]);
        put32(out, static_cast<uint32_t>(data.size()));  // compressed size
        put32(out, static_cast<uint32_t>(data.size()));  // uncompressed size
        put16(out, static_cast<uint16_t>(name.size()));
        put16(out, 0);  // extra length
        out += name;
        out += data;
    }
    uint32_t cdStart = static_cast<uint32_t>(out.size());
    for (size_t k = 0; k < members.size(); ++k) {
        const std::string& name = members[k].first;
        const std::string& data = members[k].second;
        put32(out, 0x02014b50);  // central directory signature
        put16(out, 20);          // version made by
        put16(out, 20);          // version needed
        put16(out, 0);           // flags
        put16(out, 0);           // method
        put16(out, 0);           // mod time
        put16(out, 0);           // mod date
        put32(out, crcs[k]);
        put32(out, static_cast<uint32_t>(data.size()));
        put32(out, static_cast<uint32_t>(data.size()));
        put16(out, static_cast<uint16_t>(name.size()));
        put16(out, 0);  // extra length
        put16(out, 0);  // comment length
        put16(out, 0);  // disk number start
        put16(out, 0);  // internal attributes
        put32(out, 0);  // external attributes
        put32(out, offsets[k]);
        out += name;
    }
    uint32_t cdEnd = static_cast<uint32_t>(out.size());

    put32(out, 0x06054b50);  // end of central directory
    put16(out, 0);           // this disk
    put16(out, 0);           // disk with central directory
    put16(out, static_cast<uint16_t>(members.size()));
    put16(out, static_cast<uint16_t>(members.size()));
    put32(out, cdEnd - cdStart);  // central directory size
    put32(out, cdStart);          // central directory offset
    put16(out, 0);                // comment length
    return out;
}

// ---- XML helpers ----------------------------------------------------------

// Decode the handful of XML entities Word emits in text runs.
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
            if (s.compare(i, 2, "&#") == 0) {  // numeric character reference
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
                    if (ok) {  // encode the code point as UTF-8
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

// Does the tag starting at `pos` (a '<') have local name `name`?
// Ignores any namespace prefix (the "w:" in "w:p").
bool tagIs(const std::string& xml, size_t pos, const char* name) {
    size_t i = pos + 1;
    if (i < xml.size() && xml[i] == '/') ++i;  // closing tag
    size_t colon = xml.find(':', i);
    size_t gt = xml.find_first_of(" \t\r\n/>", i);
    if (colon != std::string::npos && (gt == std::string::npos || colon < gt))
        i = colon + 1;
    size_t len = std::strlen(name);
    if (xml.compare(i, len, name) != 0) return false;
    char after = xml[i + len];
    return after == ' ' || after == '\t' || after == '\r' || after == '\n' ||
           after == '>' || after == '/';
}

// Value of attribute `name` within a single tag string (local name match,
// e.g. "val" matches w:val="..."). Returns "" if absent.
std::string attrValue(const std::string& tag, const char* name) {
    std::string key = name;
    key += "=\"";
    size_t p = tag.find(key);
    if (p == std::string::npos) return "";
    // Make sure we matched a whole attribute name, not a suffix (e.g. avoid
    // "val" matching inside "interval"). The char before must be a delimiter.
    if (p > 0) {
        char before = tag[p - 1];
        if (!(before == ' ' || before == ':' || before == '\t' ||
              before == '<'))
            return "";  // be conservative; these tags use only clean attrs
    }
    p += key.size();
    size_t q = tag.find('"', p);
    if (q == std::string::npos) return "";
    return tag.substr(p, q - p);
}

bool iequals(const std::string& a, const char* b) {
    size_t n = std::strlen(b);
    if (a.size() != n) return false;
    for (size_t i = 0; i < n; ++i)
        if (std::tolower(static_cast<unsigned char>(a[i])) !=
            std::tolower(static_cast<unsigned char>(b[i])))
            return false;
    return true;
}

// ---- Document model -------------------------------------------------------

struct Segment {
    std::string text;
    bool highlighted = false;
    std::string imageRid;  // non-empty => this segment is an image (rId)
};
struct Paragraph {
    std::vector<Segment> segs;
    bool isChapter = false;  // whole paragraph is bold -> a chapter heading
};

// Find the next closing tag "</...local>" at or after `from`.
size_t findClose(const std::string& xml, size_t from, const char* local) {
    size_t p = from;
    while ((p = xml.find("</", p)) != std::string::npos) {
        if (tagIs(xml, p, local)) return p;
        p += 2;
    }
    return std::string::npos;
}

// Is this run marked yellow? Different editors encode "standard yellow"
// differently, so we accept any of:
//   <w:highlight w:val="yellow"/>  - MS Word text highlight
//   <w:shd w:fill="FFFF00"/>       - background shading (LibreOffice/Word)
//   <w:color w:val="FFFF00"/>      - font color (Google Docs)
bool runIsYellow(const std::string& run) {
    size_t i = 0;
    while ((i = run.find('<', i)) != std::string::npos) {
        size_t e = run.find('>', i);
        if (e == std::string::npos) break;
        std::string tag = run.substr(i, e - i + 1);
        if (tagIs(run, i, "highlight")) {
            if (iequals(attrValue(tag, "val"), "yellow")) return true;
        } else if (tagIs(run, i, "shd")) {
            if (iequals(attrValue(tag, "fill"), "FFFF00")) return true;
        } else if (tagIs(run, i, "color")) {
            if (iequals(attrValue(tag, "val"), "FFFF00")) return true;
        }
        i = e + 1;
    }
    return false;
}

// Is this run bold via direct formatting (Ctrl-B), i.e. <w:b/> or
// <w:b w:val="true|1"/>? An explicit w:val of false/0/off/none turns it off.
bool runIsBold(const std::string& run) {
    size_t i = 0;
    while ((i = run.find('<', i)) != std::string::npos) {
        size_t e = run.find('>', i);
        if (e == std::string::npos) break;
        if (tagIs(run, i, "b")) {  // local name "b" only (not "bCs")
            std::string v = attrValue(run.substr(i, e - i + 1), "val");
            if (v.empty()) return true;
            return !(iequals(v, "false") || v == "0" || iequals(v, "off") ||
                     iequals(v, "none"));
        }
        i = e + 1;
    }
    return false;
}

// If this run holds an image, return its relationship id, else "".
// DrawingML: <a:blip r:embed="rIdN"/>; legacy VML: <v:imagedata r:id="rIdN"/>.
std::string runImageRid(const std::string& run) {
    size_t p = run.find("<a:blip");
    if (p != std::string::npos) {
        size_t e = run.find('>', p);
        if (e != std::string::npos) {
            std::string v = attrValue(run.substr(p, e - p + 1), "embed");
            if (!v.empty()) return v;
        }
    }
    p = run.find("<v:imagedata");
    if (p != std::string::npos) {
        size_t e = run.find('>', p);
        if (e != std::string::npos) {
            std::string v = attrValue(run.substr(p, e - p + 1), "id");
            if (!v.empty()) return v;
        }
    }
    return "";
}

// Parse word/_rels/document.xml.rels into a map of rId -> Target.
std::map<std::string, std::string> parseRels(const std::string& xml) {
    std::map<std::string, std::string> m;
    size_t i = 0;
    while ((i = xml.find("<Relationship", i)) != std::string::npos) {
        size_t e = xml.find('>', i);
        if (e == std::string::npos) break;
        std::string tag = xml.substr(i, e - i + 1);
        std::string id = attrValue(tag, "Id");
        if (!id.empty()) m[id] = attrValue(tag, "Target");
        i = e + 1;
    }
    return m;
}

// Concatenate the visible text (<w:t> contents) of a single run.
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

// Split one paragraph's inner XML into highlighted / plain segments, merging
// adjacent runs of the same kind (Word often splits a span into several runs).
Paragraph parseParagraph(const std::string& inner) {
    Paragraph para;
    bool anyText = false, allBold = true;
    size_t i = 0;
    while ((i = inner.find('<', i)) != std::string::npos) {
        if (tagIs(inner, i, "r") && inner[i + 1] != '/') {
            size_t open = inner.find('>', i);
            if (open == std::string::npos) break;
            if (inner[open - 1] == '/') { i = open + 1; continue; }  // <w:r/>
            size_t close = findClose(inner, open, "r");
            if (close == std::string::npos) break;
            std::string run = inner.substr(open + 1, close - open - 1);
            std::string imgRid = runImageRid(run);
            if (!imgRid.empty()) para.segs.push_back(Segment{"", false, imgRid});
            std::string text = runText(run);
            if (!text.empty()) {
                bool hl = runIsYellow(run);
                if (!para.segs.empty() && para.segs.back().imageRid.empty() &&
                    para.segs.back().highlighted == hl)
                    para.segs.back().text += text;
                else
                    para.segs.push_back({text, hl, ""});
                anyText = true;
                if (!runIsBold(run)) allBold = false;
            }
            i = close;
            continue;
        }
        size_t e = inner.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    // A chapter heading is a paragraph whose every text run is bold.
    para.isChapter = anyText && allBold;
    return para;
}

// Extract every body paragraph (skipping anything inside a table) as a list of
// segments.
std::vector<Paragraph> parseBody(const std::string& xml) {
    std::vector<Paragraph> paras;
    int tableDepth = 0;
    size_t i = 0;
    while (i < xml.size()) {
        if (xml[i] != '<') { ++i; continue; }
        bool closing = (i + 1 < xml.size() && xml[i + 1] == '/');

        if (tagIs(xml, i, "tbl")) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            bool selfClose = (xml[e - 1] == '/');
            if (closing) { if (tableDepth > 0) --tableDepth; }
            else if (!selfClose) ++tableDepth;
            i = e + 1;
            continue;
        }

        if (tableDepth == 0 && tagIs(xml, i, "p") && !closing) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            if (xml[e - 1] == '/') {  // empty <w:p/>
                paras.emplace_back();
                i = e + 1;
                continue;
            }
            // Paragraphs don't nest, so the next </w:p> closes this one.
            size_t close = findClose(xml, e + 1, "p");
            if (close == std::string::npos) break;
            paras.push_back(parseParagraph(xml.substr(e + 1, close - e - 1)));
            i = close;
            continue;
        }

        size_t e = xml.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    return paras;
}

// ---- Template .docx generation --------------------------------------------
//
// Build a "template" copy of the document where every yellow-highlighted span
// becomes a {{VARn}} placeholder, keeping the run's formatting so the
// placeholder itself stays yellow. The numbering and grouping below MUST mirror
// parseParagraph / parseBody and the VARS-assignment loop in main(), so that
// {{VARn}} in the template lines up with VARn in the .txt.

// Replace a run's visible text with `replacement`: the first <w:t> receives it,
// any further <w:t> in the same run are emptied. The run properties (and thus
// the highlight) are preserved untouched. `replacement` must be XML-safe (our
// placeholders contain no special characters).
std::string rewriteRunText(const std::string& run,
                           const std::string& replacement) {
    std::string out;
    size_t i = 0, copyFrom = 0;
    bool first = true;
    while ((i = run.find('<', i)) != std::string::npos) {
        if (tagIs(run, i, "t") && run[i + 1] != '/') {
            size_t open = run.find('>', i);
            if (open == std::string::npos) break;
            if (run[open - 1] == '/') { i = open + 1; continue; }  // <w:t/>
            size_t close = findClose(run, open, "t");
            if (close == std::string::npos) break;
            out += run.substr(copyFrom, open + 1 - copyFrom);  // ...<w:t ...>
            if (first) { out += replacement; first = false; }   // else: emptied
            copyFrom = close;  // resume at </w:t>
            i = close;
            continue;
        }
        size_t e = run.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    out += run.substr(copyFrom);
    return out;
}

// Kind of the most recently seen paragraph segment, used to reproduce
// parseParagraph's run-merging (adjacent highlighted runs form one variable).
enum class SegKind { None, Image, PlainText, HlText };

// Transform one paragraph's inner XML, advancing the shared variable counter.
std::string transformParagraphXml(const std::string& inner, int& counter) {
    std::string out;
    SegKind last = SegKind::None;
    size_t i = 0, copyFrom = 0;
    while ((i = inner.find('<', i)) != std::string::npos) {
        if (tagIs(inner, i, "r") && inner[i + 1] != '/') {
            size_t open = inner.find('>', i);
            if (open == std::string::npos) break;
            if (inner[open - 1] == '/') { i = open + 1; continue; }  // <w:r/>
            size_t close = findClose(inner, open, "r");
            if (close == std::string::npos) break;
            size_t closeGt = inner.find('>', close);
            if (closeGt == std::string::npos) break;
            size_t runEnd = closeGt + 1;

            out += inner.substr(copyFrom, i - copyFrom);  // gap before the run
            std::string fullRun = inner.substr(i, runEnd - i);
            std::string run = inner.substr(open + 1, close - open - 1);

            std::string imgRid = runImageRid(run);
            std::string text = runText(run);
            if (!imgRid.empty()) last = SegKind::Image;
            if (!text.empty()) {
                if (runIsYellow(run)) {
                    std::string placeholder =
                        (last == SegKind::HlText)
                            ? std::string()  // continuation of one variable
                            : "{{VAR" + std::to_string(++counter) + "}}";
                    out += rewriteRunText(fullRun, placeholder);
                    last = SegKind::HlText;
                } else {
                    out += fullRun;  // plain text, unchanged
                    last = SegKind::PlainText;
                }
            } else {
                out += fullRun;  // image-only or empty run, unchanged
            }
            copyFrom = runEnd;
            i = runEnd;
            continue;
        }
        size_t e = inner.find('>', i);
        if (e == std::string::npos) break;
        i = e + 1;
    }
    out += inner.substr(copyFrom);
    return out;
}

// Rewrite a whole word/document.xml into its template form. Tables are left
// untouched (their text is not extracted as variables, mirroring parseBody).
std::string buildTemplateXml(const std::string& xml) {
    std::string out;
    int tableDepth = 0, counter = 0;
    size_t i = 0, copyFrom = 0;
    while (i < xml.size()) {
        if (xml[i] != '<') { ++i; continue; }
        bool closing = (i + 1 < xml.size() && xml[i + 1] == '/');

        if (tagIs(xml, i, "tbl")) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            bool selfClose = (xml[e - 1] == '/');
            if (closing) { if (tableDepth > 0) --tableDepth; }
            else if (!selfClose) ++tableDepth;
            i = e + 1;
            continue;
        }

        if (tableDepth == 0 && tagIs(xml, i, "p") && !closing) {
            size_t e = xml.find('>', i);
            if (e == std::string::npos) break;
            if (xml[e - 1] == '/') { i = e + 1; continue; }  // empty <w:p/>
            size_t close = findClose(xml, e + 1, "p");
            if (close == std::string::npos) break;
            out += xml.substr(copyFrom, e + 1 - copyFrom);  // up to <w:p ...>
            out += transformParagraphXml(xml.substr(e + 1, close - e - 1),
                                         counter);
            copyFrom = close;  // resume at </w:p>
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

// ---- Word wrapping --------------------------------------------------------

size_t utf8Length(const std::string& s) {
    size_t n = 0;
    for (unsigned char c : s)
        if ((c & 0xC0) != 0x80) ++n;
    return n;
}

// A paragraph rendered as a flat character stream, where tag[i] is the
// variable id of byte text[i] (0 = plain text). This preserves the exact
// original spacing so highlighted spans abutting punctuation stay glued.
struct TaggedText {
    std::string text;
    std::vector<int> tag;
    bool chapter = false;
};

bool isSpace(char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; }

// Greedily wrap a tagged paragraph so no line exceeds kMaxLineLength code
// points (measured on the ORIGINAL text), then replace each maximal run of
// same-variable characters with a single {{VARn}} placeholder. Wrapping by the
// original width keeps line breaks identical to the source document; only the
// rendered text shows placeholders.
std::vector<std::string> wrapTagged(const TaggedText& p) {
    const std::string& t = p.text;
    // Split into words (byte ranges of non-whitespace runs).
    std::vector<std::pair<size_t, size_t>> words;
    for (size_t i = 0; i < t.size();) {
        while (i < t.size() && isSpace(t[i])) ++i;
        if (i >= t.size()) break;
        size_t s = i;
        while (i < t.size() && !isSpace(t[i])) ++i;
        words.emplace_back(s, i);
    }

    // Greedy line fill by original word width.
    std::vector<std::vector<std::pair<size_t, size_t>>> lines;
    std::vector<std::pair<size_t, size_t>> cur;
    size_t curLen = 0;
    for (auto& w : words) {
        size_t wl = utf8Length(t.substr(w.first, w.second - w.first));
        if (cur.empty()) {
            cur.push_back(w);
            curLen = wl;
        } else if (curLen + 1 + wl <= kMaxLineLength) {
            cur.push_back(w);
            curLen += 1 + wl;
        } else {
            lines.push_back(cur);
            cur.clear();
            cur.push_back(w);
            curLen = wl;
        }
    }
    if (!cur.empty()) lines.push_back(cur);

    std::vector<std::string> out;
    for (const auto& ln : lines) {
        // Reassemble the line's char+tag stream, inserting one separator space
        // between words. The separator belongs to a variable only when it sits
        // inside a single span (both neighbours share the same nonzero varId).
        std::string chars;
        std::vector<int> tags;
        for (size_t k = 0; k < ln.size(); ++k) {
            if (k > 0) {
                int prev = p.tag[ln[k - 1].second - 1];
                int next = p.tag[ln[k].first];
                chars += ' ';
                tags.push_back((prev != 0 && prev == next) ? prev : 0);
            }
            for (size_t i = ln[k].first; i < ln[k].second; ++i) {
                chars += t[i];
                tags.push_back(p.tag[i]);
            }
        }
        // Collapse each run of equal nonzero tag into a single placeholder.
        std::string line;
        for (size_t i = 0; i < chars.size();) {
            int tg = tags[i];
            if (tg != 0) {
                size_t j = i;
                while (j < chars.size() && tags[j] == tg) ++j;
                line += "\\var{VAR" + std::to_string(tg) + "}";
                i = j;
            } else {
                line += chars[i++];
            }
        }
        out.push_back(std::move(line));
    }
    return out;
}

}  // namespace

int main(int argc, char** argv) {
    if (argc != 3) {
        std::cerr << "Usage: " << argv[0] << " <input.docx> <output.txt>\n";
        return 1;
    }

    std::string zip;
    if (!readWholeFile(argv[1], zip)) {
        std::cerr << "Error: cannot read input file '" << argv[1] << "'\n";
        return 1;
    }

    std::string xml;
    if (!extractZipMember(zip, "word/document.xml", xml)) {
        std::cerr << "Error: '" << argv[1]
                  << "' is not a valid .docx (no word/document.xml)\n";
        return 1;
    }

    std::vector<Paragraph> paras = parseBody(xml);

    // Relationship map for resolving image rIds to media targets.
    std::map<std::string, std::string> rels;
    std::string relsXml;
    if (extractZipMember(zip, "word/_rels/document.xml.rels", relsXml))
        rels = parseRels(relsXml);

    // Images get extracted next to the output: <stem>_images/<file>.
    namespace fs = std::filesystem;
    fs::path outPath(argv[2]);
    std::string stem = outPath.stem().string();
    fs::path imgDir = outPath.parent_path() / (stem + "_images");
    std::string relPrefix = stem + "_images/";

    // Assign a variable to every highlighted span and an IMAGEn to every image
    // (both numbered in document order) and turn each paragraph into a tagged
    // character stream for wrapping. Images become literal \img{IMAGEn} tokens.
    std::vector<std::pair<std::string, std::string>> vars;
    std::vector<std::pair<std::string, std::string>> images;  // name -> path
    std::vector<TaggedText> paraData;
    int counter = 0, imgCounter = 0;
    for (const Paragraph& para : paras) {
        TaggedText p;
        p.chapter = para.isChapter;
        for (const Segment& seg : para.segs) {
            if (!seg.imageRid.empty()) {
                std::string name = "IMAGE" + std::to_string(++imgCounter);
                std::string target =
                    rels.count(seg.imageRid) ? rels[seg.imageRid] : "";
                std::string path = target;  // fallback (e.g. external)
                if (!target.empty()) {
                    std::string bytes;
                    std::string base = fs::path(target).filename().string();
                    if (extractZipMember(zip, "word/" + target, bytes)) {
                        std::error_code ec;
                        fs::create_directories(imgDir, ec);
                        std::ofstream f((imgDir / base).string(),
                                        std::ios::binary);
                        f.write(bytes.data(),
                                static_cast<std::streamsize>(bytes.size()));
                        path = relPrefix + base;
                    }
                }
                images.emplace_back(name, path);
                std::string ph = "\\img{" + name + "}";
                p.text += ph;
                p.tag.insert(p.tag.end(), ph.size(), 0);
                continue;
            }
            int varId = 0;
            if (seg.highlighted) {
                varId = ++counter;
                vars.emplace_back("VAR" + std::to_string(varId), seg.text);
            }
            p.text += seg.text;
            p.tag.insert(p.tag.end(), seg.text.size(), varId);
        }
        paraData.push_back(std::move(p));
    }

    std::ofstream out(argv[2], std::ios::binary);
    if (!out) {
        std::cerr << "Error: cannot write output file '" << argv[2] << "'\n";
        return 1;
    }

    // IMAGES section: IMAGEn=path (extracted media file).
    out << "========== IMAGES ==========\n";
    for (const auto& im : images) out << im.first << "=" << im.second << "\n";
    out << "\n";

    // VARS section: VARn=highlighted text (raw, without any wrapper).
    out << "========== VARS ==========\n";
    for (const auto& v : vars)
        out << v.first << "=" << v.second << "\n";
    out << "\n";

    // SETTINGS section: font sizes (points) for the rebuilt .docx. These are
    // editable defaults consumed by txt2docx.
    out << "========== SETTINGS ==========\n";
    out << "DEFAULT_FONT_SIZE=13\n";
    out << "HEADER_FONT_SIZE=16\n";
    out << "\n";

    // TEXT section: skip empty paragraphs, blank line after each, wrap at 75
    // chars by original width with {{VARn}} placeholders substituted in.
    // Chapter paragraphs (whole paragraph bold) are wrapped as \ch{...}.
    out << "========== TEXT ==========\n";
    for (const TaggedText& p : paraData) {
        std::vector<std::string> lines = wrapTagged(p);
        if (lines.empty()) continue;
        if (p.chapter) {
            lines.front().insert(0, "\\ch{");
            lines.back().push_back('}');
        }
        for (const std::string& line : lines) out << line << '\n';
        out << '\n';
    }

    // Also emit "<input-stem>_template.docx" next to the .txt: a copy of the
    // input document with every highlighted span replaced by a {{VARn}}
    // placeholder that keeps the original yellow highlight.
    std::vector<std::pair<std::string, std::string>> members;
    if (listZipMembers(zip, members)) {
        for (auto& m : members)
            if (m.first == "word/document.xml")
                m.second = buildTemplateXml(xml);
        std::string inStem = fs::path(argv[1]).stem().string();
        fs::path templatePath =
            outPath.parent_path() / (inStem + "_template.docx");
        std::string tzip = buildZipStored(members);
        std::ofstream tf(templatePath.string(), std::ios::binary);
        if (tf)
            tf.write(tzip.data(), static_cast<std::streamsize>(tzip.size()));
        else
            std::cerr << "Warning: cannot write template '"
                      << templatePath.string() << "'\n";
    } else {
        std::cerr << "Warning: could not build template .docx from input\n";
    }

    return 0;
}
