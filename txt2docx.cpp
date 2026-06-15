// txt2docx - fill a text template with variable values and write a .docx.
//
// The companion to docx2txt. It reads a template produced by docx2txt:
//
//   ========== VARS ==========
//   VAR1={{default value}}
//   ...
//   ========== TEXT ==========
//   ... paragraphs with {{VAR1}}, {{VAR2}} placeholders ...
//
// substitutes a value for every {{VARn}} placeholder, and emits a finished
// .docx. Values come from an optional values file (VARn=value, one per line);
// any variable not given a value falls back to the default in the template's
// VARS section, so an empty values file reproduces the original text.
//
// The 75-char wrapping in the template is undone (lines of a paragraph are
// rejoined with spaces) so Word re-wraps the text itself.
//
// Build:  g++ -O2 -std=c++17 -o txt2docx txt2docx.cpp -lz
// Usage:  ./txt2docx <template.txt> <values.txt> <output.docx>
//         ./txt2docx <template.txt> <output.docx>   (defaults only)

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <map>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include <zlib.h>  // crc32

namespace {

// ---- File I/O -------------------------------------------------------------

bool readWholeFile(const std::string& path, std::string& out) {
    std::ifstream in(path, std::ios::binary);
    if (!in) return false;
    out.assign(std::istreambuf_iterator<char>(in),
               std::istreambuf_iterator<char>());
    return true;
}

// ---- Small string helpers -------------------------------------------------

std::string rstripCR(const std::string& s) {
    size_t e = s.size();
    while (e > 0 && (s[e - 1] == '\r' || s[e - 1] == '\n')) --e;
    return s.substr(0, e);
}

std::string trim(const std::string& s) {
    size_t b = 0, e = s.size();
    while (b < e && std::isspace(static_cast<unsigned char>(s[b]))) ++b;
    while (e > b && std::isspace(static_cast<unsigned char>(s[e - 1]))) --e;
    return s.substr(b, e - b);
}

// Remove a surrounding {{ ... }} wrapper if present (the form docx2txt uses for
// default values), so values can be copied straight from the VARS section.
std::string stripBraces(const std::string& s) {
    if (s.size() >= 4 && s.compare(0, 2, "{{") == 0 &&
        s.compare(s.size() - 2, 2, "}}") == 0)
        return s.substr(2, s.size() - 4);
    return s;
}

std::vector<std::string> splitLines(const std::string& s) {
    std::vector<std::string> lines;
    size_t start = 0;
    for (size_t i = 0; i <= s.size(); ++i) {
        if (i == s.size() || s[i] == '\n') {
            lines.push_back(rstripCR(s.substr(start, i - start)));
            start = i + 1;
        }
    }
    if (!s.empty() && s.back() == '\n') lines.pop_back();
    return lines;
}

// ---- Template parsing -----------------------------------------------------

const char* kImagesHeader = "========== IMAGES ==========";
const char* kVarsHeader = "========== VARS ==========";
const char* kSettingsHeader = "========== SETTINGS ==========";
const char* kTextHeader = "========== TEXT ==========";

// Parse a positive integer, falling back to `def` on anything unexpected.
int parseIntOr(const std::string& s, int def) {
    try {
        size_t pos = 0;
        int v = std::stoi(trim(s), &pos);
        return v > 0 ? v : def;
    } catch (...) {
        return def;
    }
}

// Parse "VARn=value" lines into a map (used for both the template defaults and
// the values file). Blank lines and lines starting with '#' are ignored.
void parseAssignments(const std::vector<std::string>& lines, size_t from,
                      size_t to, std::map<std::string, std::string>& out) {
    for (size_t i = from; i < to && i < lines.size(); ++i) {
        std::string line = trim(lines[i]);
        if (line.empty() || line[0] == '#') continue;
        size_t eq = line.find('=');
        if (eq == std::string::npos) continue;
        std::string key = trim(line.substr(0, eq));
        std::string val = stripBraces(trim(line.substr(eq + 1)));
        if (!key.empty()) out[key] = val;
    }
}

// Split the TEXT section into paragraphs (blank-line separated), rejoining each
// paragraph's wrapped lines with single spaces.
std::vector<std::string> parseParagraphs(const std::vector<std::string>& lines,
                                         size_t from) {
    std::vector<std::string> paras;
    std::string cur;
    for (size_t i = from; i < lines.size(); ++i) {
        if (trim(lines[i]).empty()) {
            if (!cur.empty()) { paras.push_back(cur); cur.clear(); }
        } else {
            if (!cur.empty()) cur += ' ';
            cur += lines[i];
        }
    }
    if (!cur.empty()) paras.push_back(cur);
    return paras;
}

// ---- Placeholder substitution ---------------------------------------------

std::string substitute(const std::string& text,
                       const std::map<std::string, std::string>& values,
                       std::vector<std::string>& missing) {
    std::string out;
    size_t i = 0;
    while (i < text.size()) {
        if (text.compare(i, 5, "\\var{") == 0) {
            size_t close = text.find('}', i + 5);
            if (close != std::string::npos) {
                std::string name = trim(text.substr(i + 5, close - i - 5));
                auto it = values.find(name);
                if (it != values.end()) {
                    out += it->second;
                } else {
                    missing.push_back(name);
                    out += "\\var{" + name + "}";  // leave visible
                }
                i = close + 1;
                continue;
            }
        }
        out += text[i++];
    }
    return out;
}

// ---- Minimal .docx (ZIP) writer -------------------------------------------

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

struct ZipEntry {
    std::string name;
    std::string data;
    uint32_t crc;
    uint32_t offset;
};

// Build a ZIP archive with STORED (uncompressed) members — valid and openable
// by Word, and simple enough to need no deflate on the write side.
std::string buildZip(std::vector<ZipEntry> entries) {
    std::string out;
    for (auto& e : entries) {
        e.crc = crc32(0, reinterpret_cast<const Bytef*>(e.data.data()),
                      static_cast<uInt>(e.data.size()));
        e.offset = static_cast<uint32_t>(out.size());
        put32(out, 0x04034b50);             // local header signature
        put16(out, 20);                     // version needed
        put16(out, 0);                      // flags
        put16(out, 0);                      // method: stored
        put16(out, 0);                      // mod time
        put16(out, 0);                      // mod date
        put32(out, e.crc);
        put32(out, static_cast<uint32_t>(e.data.size()));  // compressed size
        put32(out, static_cast<uint32_t>(e.data.size()));  // uncompressed size
        put16(out, static_cast<uint16_t>(e.name.size()));
        put16(out, 0);                      // extra length
        out += e.name;
        out += e.data;
    }

    uint32_t cdStart = static_cast<uint32_t>(out.size());
    for (const auto& e : entries) {
        put32(out, 0x02014b50);             // central directory signature
        put16(out, 20);                     // version made by
        put16(out, 20);                     // version needed
        put16(out, 0);                      // flags
        put16(out, 0);                      // method
        put16(out, 0);                      // mod time
        put16(out, 0);                      // mod date
        put32(out, e.crc);
        put32(out, static_cast<uint32_t>(e.data.size()));
        put32(out, static_cast<uint32_t>(e.data.size()));
        put16(out, static_cast<uint16_t>(e.name.size()));
        put16(out, 0);                      // extra length
        put16(out, 0);                      // comment length
        put16(out, 0);                      // disk number start
        put16(out, 0);                      // internal attributes
        put32(out, 0);                      // external attributes
        put32(out, e.offset);
        out += e.name;
    }
    uint32_t cdEnd = static_cast<uint32_t>(out.size());

    put32(out, 0x06054b50);                 // end of central directory
    put16(out, 0);                          // this disk
    put16(out, 0);                          // disk with central directory
    put16(out, static_cast<uint16_t>(entries.size()));
    put16(out, static_cast<uint16_t>(entries.size()));
    put32(out, cdEnd - cdStart);            // central directory size
    put32(out, cdStart);                    // central directory offset
    put16(out, 0);                          // comment length
    return out;
}

// Escape text for inclusion in an XML element body.
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

// One output paragraph: its text and whether it is a chapter heading.
struct Block {
    std::string text;
    bool heading = false;
};

// An embedded image: its relationship id and display size in EMU.
struct ImgInfo {
    std::string rId;
    long cx = 0, cy = 0;
};

// Pixel size of a PNG or JPEG from its header (0,0 if unknown).
std::pair<long, long> imagePxSize(const std::string& b) {
    auto u = [&](size_t o) { return (unsigned char)b[o]; };
    auto be16 = [&](size_t o) { return (long)(u(o) << 8 | u(o + 1)); };
    auto be32 = [&](size_t o) {
        return (long)((u(o) << 24) | (u(o + 1) << 16) | (u(o + 2) << 8) |
                      u(o + 3));
    };
    if (b.size() >= 24 && u(0) == 0x89 && b[1] == 'P')  // PNG: IHDR w,h
        return {be32(16), be32(20)};
    if (b.size() >= 4 && u(0) == 0xFF && u(1) == 0xD8) {  // JPEG: scan SOF
        size_t i = 2;
        while (i + 9 < b.size()) {
            if (u(i) != 0xFF) { ++i; continue; }
            unsigned char m = u(i + 1);
            if (m == 0xD8 || m == 0xD9) { i += 2; continue; }
            long len = be16(i + 2);
            if ((m >= 0xC0 && m <= 0xC3) || (m >= 0xC5 && m <= 0xC7) ||
                (m >= 0xC9 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF))
                return {be16(i + 7), be16(i + 5)};  // width, height
            i += 2 + len;
        }
    }
    return {0, 0};
}

std::string mimeForExt(const std::string& ext) {
    if (ext == "png") return "image/png";
    if (ext == "jpg" || ext == "jpeg") return "image/jpeg";
    if (ext == "gif") return "image/gif";
    if (ext == "bmp") return "image/bmp";
    if (ext == "tif" || ext == "tiff") return "image/tiff";
    if (ext == "emf") return "image/x-emf";
    if (ext == "wmf") return "image/x-wmf";
    return "application/octet-stream";
}

// An inline <w:drawing> referencing image rId, sized cx x cy EMU.
std::string drawingXml(const ImgInfo& im, int id) {
    std::string cx = std::to_string(im.cx), cy = std::to_string(im.cy),
                sid = std::to_string(id);
    return "<w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" "
           "distR=\"0\"><wp:extent cx=\"" + cx + "\" cy=\"" + cy + "\"/>"
           "<wp:docPr id=\"" + sid + "\" name=\"img" + sid + "\"/>"
           "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/"
           "drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"" +
           sid + "\" name=\"img" + sid + "\"/><pic:cNvPicPr/></pic:nvPicPr>"
           "<pic:blipFill><a:blip r:embed=\"" + im.rId + "\"/><a:stretch>"
           "<a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>"
           "<a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + cx + "\" cy=\"" + cy +
           "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
           "</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>"
           "</w:drawing>";
}

std::string buildDocumentXml(const std::vector<Block>& paragraphs,
                             int defaultPt, int headerPt,
                             const std::map<std::string, ImgInfo>& imgs) {
    std::string body;
    body +=
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/"
        "wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats."
        "org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas."
        "openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a="
        "\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic="
        "\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
        "<w:body>";
    // One blank line, in twips (1/20 pt), measured in the body font. Used as
    // the space above and below chapter headings.
    std::string gap = std::to_string(defaultPt * 20);
    int docId = 0;  // unique drawing ids

    for (const Block& p : paragraphs) {
        // Font size is stored in half-points in OOXML, hence *2.
        int half = (p.heading ? headerPt : defaultPt) * 2;
        std::string sz = std::to_string(half);
        std::string rpr =
            "<w:rPr><w:rFonts w:ascii=\"Times New Roman\" "
            "w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>";
        if (p.heading) rpr += "<w:b/><w:bCs/>";  // chapters are bold
        rpr += "<w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/>";
        rpr += "</w:rPr>";

        // No spacing between ordinary paragraphs; one line above and below
        // chapter headings. (Word/LibreOffice take after(prev)+before(next),
        // so a heading is framed by exactly one line on each side.)
        std::string spacing =
            p.heading ? "<w:spacing w:before=\"" + gap + "\" w:after=\"" + gap +
                            "\" w:line=\"240\" w:lineRule=\"auto\"/>"
                      : "<w:spacing w:before=\"0\" w:after=\"0\" "
                        "w:line=\"240\" w:lineRule=\"auto\"/>";

        body += "<w:p><w:pPr>" + spacing + rpr + "</w:pPr>";
        // Ordinary paragraphs start with one tab; chapter headings do not.
        if (!p.heading) body += "<w:r>" + rpr + "<w:tab/></w:r>";

        // Split the paragraph on \img{NAME} tokens: text -> text run,
        // image -> drawing run.
        auto emitText = [&](const std::string& s) {
            if (!s.empty())
                body += "<w:r>" + rpr + "<w:t xml:space=\"preserve\">" +
                        xmlEscape(s) + "</w:t></w:r>";
        };
        const std::string& c = p.text;
        size_t i = 0;
        while (i < c.size()) {
            size_t pos = c.find("\\img{", i);
            if (pos == std::string::npos) { emitText(c.substr(i)); break; }
            emitText(c.substr(i, pos - i));
            size_t close = c.find('}', pos + 5);
            if (close == std::string::npos) { emitText(c.substr(pos)); break; }
            std::string name = c.substr(pos + 5, close - pos - 5);
            auto it = imgs.find(name);
            if (it != imgs.end())
                body += "<w:r>" + drawingXml(it->second, ++docId) + "</w:r>";
            else
                emitText("\\img{" + name + "}");  // unknown -> leave visible
            i = close + 1;
        }
        body += "</w:p>";
    }
    body += "</w:body></w:document>";
    return body;
}

}  // namespace

int main(int argc, char** argv) {
    std::string templatePath, valuesPath, outputPath;
    if (argc == 4) {
        templatePath = argv[1];
        valuesPath = argv[2];
        outputPath = argv[3];
    } else if (argc == 3) {
        templatePath = argv[1];
        outputPath = argv[2];
    } else {
        std::cerr << "Usage: " << argv[0]
                  << " <template.txt> [values.txt] <output.docx>\n";
        return 1;
    }

    std::string tmpl;
    if (!readWholeFile(templatePath, tmpl)) {
        std::cerr << "Error: cannot read template '" << templatePath << "'\n";
        return 1;
    }
    std::vector<std::string> lines = splitLines(tmpl);

    // Locate the section headers.
    size_t imagesAt = lines.size(), varsAt = lines.size(),
           settingsAt = lines.size(), textAt = lines.size();
    for (size_t i = 0; i < lines.size(); ++i) {
        if (trim(lines[i]) == kImagesHeader) imagesAt = i;
        else if (trim(lines[i]) == kVarsHeader) varsAt = i;
        else if (trim(lines[i]) == kSettingsHeader) settingsAt = i;
        else if (trim(lines[i]) == kTextHeader) textAt = i;
    }
    if (textAt == lines.size()) {
        std::cerr << "Error: template has no '" << kTextHeader << "' section\n";
        return 1;
    }

    // Variable defaults from the template's VARS section (bounded before the
    // SETTINGS section), overridden by the values file.
    std::map<std::string, std::string> values;
    size_t varsEnd = std::min(settingsAt, textAt);
    if (varsAt < varsEnd) parseAssignments(lines, varsAt + 1, varsEnd, values);

    // Image paths from the template IMAGES section (bounded before VARS).
    std::map<std::string, std::string> imgPaths;
    if (imagesAt < lines.size()) {
        size_t imgEnd = std::min({varsAt, settingsAt, textAt});
        parseAssignments(lines, imagesAt + 1, imgEnd, imgPaths);
    }

    // Font sizes from the template SETTINGS section.
    std::map<std::string, std::string> settings;
    if (settingsAt < textAt)
        parseAssignments(lines, settingsAt + 1, textAt, settings);

    if (!valuesPath.empty()) {
        std::string vfile;
        if (!readWholeFile(valuesPath, vfile)) {
            std::cerr << "Error: cannot read values '" << valuesPath << "'\n";
            return 1;
        }
        std::vector<std::string> vlines = splitLines(vfile);
        // Values file may override variables, font settings and image paths.
        parseAssignments(vlines, 0, vlines.size(), values);
        parseAssignments(vlines, 0, vlines.size(), settings);
        parseAssignments(vlines, 0, vlines.size(), imgPaths);
    }

    int defaultPt = 13, headerPt = 16;
    if (settings.count("DEFAULT_FONT_SIZE"))
        defaultPt = parseIntOr(settings["DEFAULT_FONT_SIZE"], 13);
    if (settings.count("HEADER_FONT_SIZE"))
        headerPt = parseIntOr(settings["HEADER_FONT_SIZE"], 16);

    // Split TEXT into paragraphs; detect \ch{...} chapter headings, then
    // substitute placeholders in the (un-wrapped) content.
    std::vector<std::string> rawParas = parseParagraphs(lines, textAt + 1);
    std::vector<Block> paragraphs;
    std::vector<std::string> missing;
    for (const std::string& raw : rawParas) {
        bool heading = raw.size() >= 5 && raw.compare(0, 4, "\\ch{") == 0 &&
                       raw.back() == '}';
        std::string content = heading ? raw.substr(4, raw.size() - 5) : raw;
        paragraphs.push_back({substitute(content, values, missing), heading});
    }
    if (!missing.empty()) {
        std::cerr << "Warning: no value for:";
        for (const std::string& m : missing) std::cerr << " \\var{" << m << "}";
        std::cerr << " (left as-is)\n";
    }

    namespace fs = std::filesystem;
    fs::path tmplDir = fs::path(templatePath).parent_path();

    // Embed each referenced image: read the file, size it, add a media part
    // and a document relationship. EMU = px * 9525 (96 dpi); cap width at 6in.
    std::map<std::string, ImgInfo> imgInfo;
    std::vector<ZipEntry> media;
    std::string docRels;
    std::set<std::string> exts;
    int rid = 0;
    for (const auto& kv : imgPaths) {
        const std::string& name = kv.first;
        bool used = false;  // only embed images actually referenced in TEXT
        for (const Block& b : paragraphs)
            if (b.text.find("\\img{" + name + "}") != std::string::npos) {
                used = true;
                break;
            }
        if (!used) continue;

        fs::path ip(kv.second);
        if (ip.is_relative()) ip = tmplDir / ip;
        std::string bytes;
        if (!readWholeFile(ip.string(), bytes)) {
            std::cerr << "Warning: cannot read image '" << kv.second
                      << "' for " << name << " (skipped)\n";
            continue;
        }
        auto [w, h] = imagePxSize(bytes);
        if (w <= 0) w = 400;
        if (h <= 0) h = 300;
        long cx = w * 9525, cy = h * 9525;
        const long maxCx = 5486400;  // 6 inches
        if (cx > maxCx) { cy = cy * maxCx / cx; cx = maxCx; }

        std::string ext = ip.extension().string();
        if (!ext.empty() && ext[0] == '.') ext.erase(0, 1);
        for (char& c : ext) c = (char)std::tolower((unsigned char)c);
        exts.insert(ext);

        ++rid;
        std::string rId = "rId" + std::to_string(rid);
        std::string target = "media/image" + std::to_string(rid) + "." + ext;
        media.push_back({"word/" + target, bytes, 0, 0});
        docRels += "<Relationship Id=\"" + rId +
                   "\" Type=\"http://schemas.openxmlformats.org/officeDocument/"
                   "2006/relationships/image\" Target=\"" + target + "\"/>";
        imgInfo[name] = {rId, cx, cy};
    }

    // Content types: base + a Default for each image extension used.
    std::string ctypes =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/"
        "content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/"
        "vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
    for (const std::string& e : exts)
        ctypes += "<Default Extension=\"" + e + "\" ContentType=\"" +
                  mimeForExt(e) + "\"/>";
    ctypes +=
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/"
        "vnd.openxmlformats-officedocument.wordprocessingml.document.main"
        "+xml\"/></Types>";

    std::vector<ZipEntry> entries;
    entries.push_back({"[Content_Types].xml", ctypes, 0, 0});
    entries.push_back(
        {"_rels/.rels",
         "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/"
         "2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://"
         "schemas.openxmlformats.org/officeDocument/2006/relationships/"
         "officeDocument\" Target=\"word/document.xml\"/></Relationships>",
         0, 0});
    entries.push_back(
        {"word/document.xml",
         buildDocumentXml(paragraphs, defaultPt, headerPt, imgInfo), 0, 0});
    // word/_rels/document.xml.rels (image relationships).
    entries.push_back(
        {"word/_rels/document.xml.rels",
         "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/"
         "2006/relationships\">" + docRels + "</Relationships>",
         0, 0});
    for (auto& m : media) entries.push_back(std::move(m));

    std::string zip = buildZip(std::move(entries));

    std::ofstream out(outputPath, std::ios::binary);
    if (!out) {
        std::cerr << "Error: cannot write output '" << outputPath << "'\n";
        return 1;
    }
    out.write(zip.data(), static_cast<std::streamsize>(zip.size()));
    return 0;
}
