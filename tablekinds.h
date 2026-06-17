// tablekinds.h - the table module for docxform.
//
// This is the CUSTOMIZATION POINT of the program: everything about which tables
// exist, how many rows/columns they have, their column titles and their content
// lives in tablekinds.cpp. The rest of docxform never needs to change when you
// add a new kind of table.
//
// It is intentionally self-contained (it does not depend on the rest of
// docxform), so the whole thing can be reused as a module inside another Qt
// application: link tablekinds.cpp, call tableKinds() to list the kinds and
// buildTableXml() to turn a chosen kind's data into OOXML.

#ifndef DOCXFORM_TABLEKINDS_H
#define DOCXFORM_TABLEKINDS_H

#include <functional>
#include <string>
#include <vector>

namespace docxform {

// The concrete content of one table: a header row (column titles) plus any
// number of body rows. Ragged rows are allowed (short rows are padded with
// empty cells to the widest row / header).
struct TableData {
    std::vector<std::string> headers;             // column titles (top row)
    std::vector<std::vector<std::string>> rows;   // body rows (cells, UTF-8)
};

// A selectable kind of table, shown to the user in a drop-down in the GUI.
struct TableKind {
    std::string id;     // stable identifier (not shown in the GUI)
    std::string title;  // label shown in the GUI drop-down

    // Produces the table content. `name` is the {name} taken from the
    // \table{name} placeholder in the document, so a kind can adapt its content
    // to the placeholder it is filling (or ignore it). This is called at
    // generation time, so you may compute rows on the fly here.
    std::function<TableData(const std::string& name)> build;
};

// THE REGISTRY. Returns every table kind the GUI offers. Edit the body of this
// function (in tablekinds.cpp) to add, remove or change kinds.
std::vector<TableKind> tableKinds();

// A "fixed" table tag: a bare placeholder (no braces, e.g. "\\tablewage") that
// is ALWAYS replaced by one specific table kind, with no drop-down in the GUI
// and no argument in headless mode. Use it for tables that always look the same
// in your documents.
struct FixedTable {
    std::string tag;     // the literal placeholder in the document, e.g. "\\tablewage"
    std::string kindId;  // id of the TableKind (see tableKinds()) to always use
};

// THE FIXED-TAG REGISTRY. Returns every fixed tag and the table kind it always
// maps to. Edit the body of this function (in tablekinds.cpp) to add your own,
// e.g. map "\\tableprice" to the "price" kind. Tags must start with '\\', be
// distinct, and not be a prefix of one another or of "\\table{".
std::vector<FixedTable> fixedTables();

// Render a TableData as an OOXML <w:tbl>...</w:tbl> fragment (with visible
// borders and a bold header row). Reusable on its own.
std::string buildTableXml(const TableData& table);

}  // namespace docxform

#endif  // DOCXFORM_TABLEKINDS_H
