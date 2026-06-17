// tablekinds.cpp - table kinds and OOXML table rendering for docxform.
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
//       user presses "Создать документ".
//     - The `name` argument is the {name} from \table{name} in the document,
//       in case you want the content to depend on it.
//  2. Register it by adding one entry to the vector returned by tableKinds():
//         kinds.push_back({"my_id", "Название в списке", myBuilderFunction});
//     `id` is a stable identifier (used by --render and code); `title` is what
//     the user sees in the drop-down.
//  3. Rebuild: `make docxform`. The new kind appears automatically in the GUI
//     for every \table{...} placeholder.
//
//  Cells are plain text (UTF-8). Newlines inside a cell are kept as separate
//  lines. Borders and a bold header row are added automatically by
//  buildTableXml(), so builders only deal with content.
// =====================================================================

#include "tablekinds.h"

#include <algorithm>
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

}  // namespace

// ---- The registry ---------------------------------------------------------
// Add or remove entries here to change what the GUI offers.
std::vector<TableKind> tableKinds() {
    std::vector<TableKind> kinds;
    kinds.push_back({"blank3x3",  "Пустая таблица 3×3",            blankGrid});
    kinds.push_back({"price",     "Прайс-лист (пример)",            priceList});
    kinds.push_back({"employees", "Сотрудники (пример)",            employees});
    kinds.push_back({"schedule",  "График платежей (генерируется)", paymentSchedule});
    return kinds;
}

// ---- The fixed-tag registry -----------------------------------------------
// Map bare tags like "\tablewage" to a specific kind id from tableKinds() above.
// Wherever such a tag appears in a document, that table is inserted
// automatically — no GUI drop-down, no headless argument. Add your own lines.
//
// Steps to add a new fixed tag:
//   1. Make sure the table kind exists in tableKinds() (note its id).
//   2. Add a {tag, kindId} line below. The tag is the exact text you type in
//      Word, including the leading backslash, e.g. "\\tablewage".
std::vector<FixedTable> fixedTables() {
    std::vector<FixedTable> fixed;
    fixed.push_back({"\\tableprice", "price"});      // \tableprice -> price kind
    fixed.push_back({"\\tablewage",  "employees"});  // \tablewage  -> employees kind
    return fixed;
}

// ---- OOXML rendering ------------------------------------------------------

std::string buildTableXml(const TableData& table) {
    // Column count = widest of the header and all rows (at least 1).
    size_t cols = table.headers.size();
    for (const auto& r : table.rows) cols = std::max(cols, r.size());
    if (cols == 0) cols = 1;

    // Equal column widths, independent of content: the table spans a fixed total
    // width (~full text area of a default page) split evenly between the columns.
    // Combined with a fixed table layout below, Word keeps every column the same
    // width regardless of what each cell contains.
    const int kTotalDxa = 9360;  // ~6.5" in twips; close to a default page's text width
    int colDxa = static_cast<int>(kTotalDxa / cols);
    if (colDxa < 1) colDxa = 1;
    const int tableDxa = colDxa * static_cast<int>(cols);
    const std::string colW = std::to_string(colDxa);

    auto cell = [&](const std::string& text, bool header) {
        std::string rpr = header ? "<w:rPr><w:b/></w:rPr>" : "";
        // Header cells are bold but keep the same white background as the rest
        // of the table (no <w:shd> fill). Every cell gets the same preferred
        // width so all columns come out equal.
        std::string tcpr =
            "<w:tcPr><w:tcW w:w=\"" + colW + "\" w:type=\"dxa\"/></w:tcPr>";
        // Split the cell text on newlines into separate paragraphs.
        std::string body;
        size_t start = 0;
        for (size_t i = 0; i <= text.size(); ++i) {
            if (i == text.size() || text[i] == '\n') {
                std::string line = text.substr(start, i - start);
                body += "<w:p><w:r>" + rpr + "<w:t xml:space=\"preserve\">" +
                        xmlEscape(line) + "</w:t></w:r></w:p>";
                start = i + 1;
            }
        }
        if (body.empty()) body = "<w:p/>";  // a cell must contain a paragraph
        return "<w:tc>" + tcpr + body + "</w:tc>";
    };

    auto row = [&](const std::vector<std::string>& cells, bool header) {
        std::string r = "<w:tr>";
        for (size_t c = 0; c < cols; ++c)
            r += cell(c < cells.size() ? cells[c] : std::string(), header);
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
    for (size_t c = 0; c < cols; ++c)
        x += "<w:gridCol w:w=\"" + colW + "\"/>";
    x += "</w:tblGrid>";

    if (!table.headers.empty()) x += row(table.headers, true);
    for (const auto& r : table.rows) x += row(r, false);
    if (table.headers.empty() && table.rows.empty())
        x += row({}, false);  // guarantee at least one row (valid OOXML)

    x += "</w:tbl>";
    return x;
}

}  // namespace docxform
