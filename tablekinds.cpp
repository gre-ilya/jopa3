// tablekinds.cpp - table builders and OOXML table rendering for docxform.
//
// =====================================================================
//  HOW TO ADD A NEW KIND OF TABLE
// =====================================================================
//  1. Write a function that returns a TableData (see the examples below).
//     - `headers` are the column titles (the top row). Leave it empty for a
//       table without a header row.
//     - `rows` is a vector of rows; each row is a vector of cell strings.
//     - You can hard-code the content, or COMPUTE it at runtime (loops,
//       dates, data from your own app, etc.) - this function runs when the
//       document is generated.
//     - The `tag` argument is the fixed tag being filled (e.g. "\\tablewage"),
//       in case one function serves several tags.
//  2. Bind it to a tag by adding one entry to the vector returned by
//     fixedTables():
//         fixed.push_back({"\\mytable", myBuilderFunction});
//     Now writing \mytable in a document always inserts that table.
//  3. Rebuild: `make docxform`.
//
//  Cells are plain text (UTF-8). Newlines inside a cell are kept as separate
//  lines. Borders and centered cell text are added automatically by
//  buildTableXml(), so builders only deal with content.
// =====================================================================

#include "tablekinds.h"

#include <algorithm>
#include <ctime>
#include <string>
#include <vector>

namespace docxform {
namespace {

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

// Two-digit zero-padded number (helper for the runtime demo below).
std::string pad2(int n) {
    std::string s = std::to_string(n);
    return s.size() < 2 ? "0" + s : s;
}

// ---- Demo table builders --------------------------------------------------
// Each returns the content of one table. Add your own next to these.

// 1) An empty 3x3 grid the user fills in by hand in Word afterwards.
TableData blankGrid(const std::string&) {
    TableData t;
    t.headers = {"Колонка 1", "Колонка 2", "Колонка 3"};
    t.rows = {
        {"", "", ""},
        {"", "", ""},
        {"", "", ""},
    };
    return t;
}

// 2) A small price list with fixed sample content.
TableData priceList(const std::string&) {
    TableData t;
    t.headers = {"Товар", "Кол-во", "Цена"};
    t.rows = {
        {"Бумага A4, 500 л.", "10", "350"},
        {"Ручка шариковая",   "50", "25"},
        {"Папка-скоросшиватель", "20", "40"},
    };
    return t;
}

// 3) A list of employees with fixed sample content.
TableData employees(const std::string&) {
    TableData t;
    t.headers = {"ФИО", "Должность", "Телефон"};
    t.rows = {
        {"Иванов И. И.",  "Директор",    "+7 900 000-00-01"},
        {"Петрова А. С.", "Бухгалтер",   "+7 900 000-00-02"},
        {"Сидоров П. П.", "Менеджер",    "+7 900 000-00-03"},
    };
    return t;
}

// 4) A payment schedule whose rows are GENERATED IN CODE at run time. This
//    shows how to compute content dynamically instead of hard-coding it.
TableData paymentSchedule(const std::string&) {
    TableData t;
    t.headers = {"№", "Дата платежа", "Сумма, руб."};
    const int months = 6;
    const long monthly = 15000;
    for (int m = 1; m <= months; ++m) {
        std::string date = "2026-" + pad2(m) + "-15";  // computed value
        t.rows.push_back({std::to_string(m), date, std::to_string(monthly)});
    }
    return t;
}

// 5) FIXED COLUMNS, RUNTIME ROWS. The columns never change, but the cell values
//    come from data your program supplies at generation time via tableContext(),
//    keyed by the placeholder/tag name (`name`). This is the pattern to use for a
//    fixed tag whose contents differ per run — fill tableContext()[tag] from your
//    host app before generating, and the rows appear here. See README and the
//    "\tableruntime" fixed tag below.
TableData runtimeTable(const std::string& name) {
    TableData t;
    t.headers = {"Колонка A", "Колонка B"};  // fixed columns, change as needed
    auto it = tableContext().find(name);
    if (it != tableContext().end())
        t.rows = it->second;  // cells supplied by the host at runtime
    return t;
}

// 6) Logic dedicated to ONE specific tag (\tablewage). A fixed tag is bound
//    straight to a function like this, so you run exactly the code you want for
//    that tag. Compute the rows however you like (here: a tiny hard-coded
//    example; in your app pull them from your data, a file, a DB, etc.).
TableData wageReport(const std::string& /*tag*/) {
    TableData t;
    t.headers = {"Сотрудник", "Оклад, руб."};  // fixed columns
    t.rows = {
        {"Иванов И. И.", "100000"},
        {"Петров П. П.", "90000"},
    };
    return t;
}

// 7) Rows with DIFFERENT numbers of columns (merged cells via column spans). The
//    first row has 3 cells, the second has 5; cell 2 of row 1 spans columns 2-3
//    and cell 3 spans columns 4-5 (cell 1 sits over column 1). The grid has 5
//    columns; spans say how many each cell covers.
TableData spanDemo(const std::string& /*tag*/) {
    TableData t;
    t.headers = {"Группа A", "Группа B", "Группа C"};
    t.headerSpans = {1, 2, 2};  // 1 + 2 + 2 = 5 grid columns
    t.rows = {
        {"a", "b1", "b2", "c1", "c2"},  // 5 cells, each spans 1
    };
    return t;
}

// ---- Demo TEXT builders ---------------------------------------------------
// Each returns the text that its tag expands to (see fixedTexts() below).

// A constant text tag: \company always expands to this string.
std::string companyText(const std::string& /*tag*/) {
    return "ООО «Ромашка»";
}

// A COMPUTED text tag: \today expands to the current date (YYYY-MM-DD) at the
// moment the document is generated.
std::string todayText(const std::string& /*tag*/) {
    std::time_t now = std::time(nullptr);
    std::tm* lt = std::localtime(&now);
    return std::to_string(lt->tm_year + 1900) + "-" + pad2(lt->tm_mon + 1) +
           "-" + pad2(lt->tm_mday);
}

// A text tag whose value the HOST program supplies at runtime via textContext(),
// keyed by the tag. Fill textContext()[tag] before generating; this returns it.
std::string contextText(const std::string& tag) {
    auto it = textContext().find(tag);
    return it != textContext().end() ? it->second : std::string();
}

}  // namespace

// Runtime data store read by builders such as runtimeTable(). One process-wide
// instance; your host application fills it before generating (see the header).
std::map<std::string, std::vector<std::vector<std::string>>>& tableContext() {
    static std::map<std::string, std::vector<std::vector<std::string>>> ctx;
    return ctx;
}

// Runtime text store read by contextText(); the host fills it before generating.
std::map<std::string, std::string>& textContext() {
    static std::map<std::string, std::string> ctx;
    return ctx;
}

// ---- The table registry ---------------------------------------------------
// Each table in your documents has its own bare tag (e.g. "\tablewage") bound to
// the builder that produces it. Writing the tag in a document always inserts
// that table — no GUI, no headless argument. Add or change entries here.
//
// To add a new table: write a builder above, then add a {tag, builder} line.
// The tag is the exact text you type in Word, with the leading backslash.
std::vector<FixedTable> fixedTables() {
    std::vector<FixedTable> fixed;
    fixed.push_back({"\\tableblank",     blankGrid});        // empty 3x3 grid
    fixed.push_back({"\\tableprice",     priceList});        // sample price list
    fixed.push_back({"\\tableemployees", employees});        // sample employees
    fixed.push_back({"\\tableschedule",  paymentSchedule});  // generated in code
    fixed.push_back({"\\tablewage",      wageReport});        // logic for this tag
    fixed.push_back({"\\tableruntime",   runtimeTable});      // rows from tableContext()
    fixed.push_back({"\\tablespan",      spanDemo});          // merged cells / spans
    return fixed;
}

// ---- The text-tag registry ------------------------------------------------
// Each entry binds a bare tag (e.g. "\company") to a function returning the text
// it expands to, inline. Writing the tag in a document always inserts that text.
//
// To add a new text tag: write a function returning a std::string (a constant or
// computed), then add a {tag, function} line. The tag is the exact text you type
// in Word, with the leading backslash.
std::vector<FixedText> fixedTexts() {
    std::vector<FixedText> texts;
    texts.push_back({"\\company", companyText});  // constant text
    texts.push_back({"\\today",   todayText});    // computed at generation time
    texts.push_back({"\\note",    contextText});  // text from textContext() (host)
    return texts;
}

// ---- OOXML rendering ------------------------------------------------------

std::string buildTableXml(const TableData& table) {
    // The span of one cell (how many grid columns it occupies): the matching
    // entry in `spans` if present and positive, otherwise 1.
    auto spanOf = [](const std::vector<int>& spans, size_t c) {
        return (c < spans.size() && spans[c] > 0) ? spans[c] : 1;
    };
    // Total grid columns a row occupies = sum of its cells' spans.
    auto rowTotal = [&](const std::vector<std::string>& cells,
                        const std::vector<int>& spans) {
        int total = 0;
        for (size_t c = 0; c < cells.size(); ++c) total += spanOf(spans, c);
        return total;
    };

    // Grid column count = widest row's total span (header included), at least 1.
    // Without any spans this is just the widest row's cell count (as before).
    int gridCols = rowTotal(table.headers, table.headerSpans);
    for (size_t r = 0; r < table.rows.size(); ++r) {
        const std::vector<int>& sp =
            r < table.rowSpans.size() ? table.rowSpans[r] : std::vector<int>();
        gridCols = std::max(gridCols, rowTotal(table.rows[r], sp));
    }
    if (gridCols < 1) gridCols = 1;

    // Equal grid-column widths, independent of content: the table spans a fixed
    // total width split evenly. With a fixed table layout below, Word keeps every
    // grid column the same width; a cell spanning N columns is N times as wide.
    const int kTotalDxa = 9360;  // ~6.5" in twips; close to a default page's text width
    int colDxa = kTotalDxa / gridCols;
    if (colDxa < 1) colDxa = 1;
    const int tableDxa = colDxa * gridCols;
    const std::string colW = std::to_string(colDxa);

    // One cell occupying `span` grid columns: width = span*colDxa, plus a
    // <w:gridSpan> when it spans more than one. Centered horizontally/vertically.
    auto cell = [&](const std::string& text, int span) {
        std::string tcpr = "<w:tcPr><w:tcW w:w=\"" +
                           std::to_string(colDxa * span) + "\" w:type=\"dxa\"/>";
        if (span > 1) tcpr += "<w:gridSpan w:val=\"" + std::to_string(span) + "\"/>";
        tcpr += "<w:vAlign w:val=\"center\"/></w:tcPr>";
        // Split the cell text on newlines into separate (centered) paragraphs.
        std::string body;
        size_t start = 0;
        for (size_t i = 0; i <= text.size(); ++i) {
            if (i == text.size() || text[i] == '\n') {
                std::string line = text.substr(start, i - start);
                body += "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r>"
                        "<w:t xml:space=\"preserve\">" + xmlEscape(line) +
                        "</w:t></w:r></w:p>";
                start = i + 1;
            }
        }
        if (body.empty())  // a cell must contain a paragraph
            body = "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:p>";
        return "<w:tc>" + tcpr + body + "</w:tc>";
    };

    auto row = [&](const std::vector<std::string>& cells,
                   const std::vector<int>& spans) {
        std::string r = "<w:tr>";
        int used = 0;
        for (size_t c = 0; c < cells.size(); ++c) {
            int span = spanOf(spans, c);
            r += cell(cells[c], span);
            used += span;
        }
        // Pad short rows to the full grid width with empty single-column cells.
        for (; used < gridCols; ++used) r += cell(std::string(), 1);
        r += "</w:tr>";
        return r;
    };

    std::string x = "<w:tbl><w:tblPr>";
    x += "<w:tblW w:w=\"" + std::to_string(tableDxa) + "\" w:type=\"dxa\"/>";
    x += "<w:tblBorders>";
    for (const char* side :
         {"top", "left", "bottom", "right", "insideH", "insideV"})
        x += std::string("<w:") + side +
             " w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
    x += "</w:tblBorders>";
    x += "<w:tblLayout w:type=\"fixed\"/>";  // honour the column widths exactly
    x += "</w:tblPr>";

    x += "<w:tblGrid>";
    for (int c = 0; c < gridCols; ++c)
        x += "<w:gridCol w:w=\"" + colW + "\"/>";
    x += "</w:tblGrid>";

    if (!table.headers.empty()) x += row(table.headers, table.headerSpans);
    for (size_t r = 0; r < table.rows.size(); ++r)
        x += row(table.rows[r],
                 r < table.rowSpans.size() ? table.rowSpans[r]
                                           : std::vector<int>());
    if (table.headers.empty() && table.rows.empty())
        x += row({}, {});  // guarantee at least one row (valid OOXML)

    x += "</w:tbl>";
    return x;
}

}  // namespace docxform
