// tablekinds.h - the custom-tag module for docxform.
//
// This is the CUSTOMIZATION POINT of the program: the tags your documents use
// and the content they expand to live in tablekinds.cpp. Two kinds of tag:
//   - fixedTables() — tags that insert a generated table (e.g. \tablewage);
//   - fixedTexts()  — tags replaced inline by text your code returns (e.g.
//                     \company).
// The rest of docxform never needs to change when you add a new tag.
//
// It is intentionally self-contained (it does not depend on the rest of
// docxform), so the whole thing can be reused as a module inside another Qt
// application: link tablekinds.cpp, define your tags in fixedTables()/
// fixedTexts() and use buildTableXml() to turn a table's data into OOXML.

#ifndef DOCXFORM_TABLEKINDS_H
#define DOCXFORM_TABLEKINDS_H

#include <functional>
#include <map>
#include <string>
#include <vector>

namespace docxform {

// The concrete content of one table: a header row (column titles) plus any
// number of body rows. Ragged rows are allowed (short rows are padded with
// empty cells to the widest row / header).
struct TableData {
    std::vector<std::string> headers;             // column titles (top row)
    std::vector<std::vector<std::string>> rows;   // body rows (cells, UTF-8)

    // Optional horizontal column spans (merged cells), for tables whose rows have
    // DIFFERENT numbers of columns. `headerSpans` is parallel to `headers`,
    // `rowSpans[i]` is parallel to `rows[i]`; a cell's span is how many columns of
    // the underlying grid it occupies (missing / <= 0 means 1). The grid has as
    // many columns as the widest row's total span. Leave both empty for an
    // ordinary equal-column table.
    //
    // Example — a 3-cell row over a 5-cell row (2nd cell spans cols 2-3, 3rd spans
    // cols 4-5):
    //     t.headers = {"A", "B", "C"};   t.headerSpans = {1, 2, 2};
    //     t.rows    = {{"1","2","3","4","5"}};   // rowSpans omitted -> all 1
    std::vector<int> headerSpans;            // parallel to `headers` (optional)
    std::vector<std::vector<int>> rowSpans;  // parallel to `rows`    (optional)

    // Optional widths of the underlying GRID columns, as RELATIVE weights (any
    // positive numbers, e.g. {2,1,1} = first column twice as wide as the others).
    // One entry per grid column; missing / <= 0 entries default to weight 1.
    // Empty = all columns equal. A cell that spans several columns is as wide as
    // their widths combined. The table always fills the same total page width.
    std::vector<int> colWidths;
};

// A "fixed" table tag: a bare placeholder (no braces, e.g. "\\tablewage") bound
// DIRECTLY to its own logic. Wherever the tag appears it is always replaced by
// whatever its `build` returns — no GUI, no headless argument. Each kind of
// table in your documents gets its own tag here.
struct FixedTable {
    std::string tag;  // the literal placeholder in the document, e.g. "\\tablewage"

    // The logic that produces THIS tag's table. Called at generation time with
    // the tag itself (so one function can serve several tags if you want).
    // Return the table content: keep `headers` fixed and compute `rows` however
    // you like.
    std::function<TableData(const std::string& tag)> build;
};

// THE TABLE REGISTRY. Returns every fixed tag and the builder it always runs.
// This is the CUSTOMIZATION POINT: edit the body of this function (in
// tablekinds.cpp) to add your own, e.g. bind "\\tablewage" to your buildWage().
// Tags must start with '\\' and be distinct (and not a prefix of one another).
std::vector<FixedTable> fixedTables();

// A custom TEXT tag: a bare placeholder (e.g. "\\company") replaced INLINE by
// text your code returns. Unlike a fixed table (which replaces a whole
// paragraph), a text tag substitutes just the tag and keeps the surrounding text
// and run formatting — like a code-defined \var. No GUI, no headless argument.
struct FixedText {
    std::string tag;  // the literal placeholder in the document, e.g. "\\company"

    // Returns the replacement text for this tag. Called at generation time with
    // the tag itself, so the text may be constant or computed (date, data from
    // your app, etc.). It is plain text and is XML-escaped on insertion.
    std::function<std::string(const std::string& tag)> build;
};

// THE TEXT-TAG REGISTRY. Returns every custom text tag and the function that
// produces its text. Edit the body of this function (in tablekinds.cpp) to add
// your own, e.g. bind "\\company" to a function returning your company name.
// Tags follow the same rules as table tags (start with '\\', be distinct, not a
// prefix of one another or of a table tag).
std::vector<FixedText> fixedTexts();

// Runtime data your application feeds to table builders. A build() function runs
// at GENERATION time, so it can return DIFFERENT rows every time depending on
// what your program put here — while the columns (headers) stay fixed. Key it by
// the tag the table fills (the `tag` build() receives, e.g. "\\tablewage") so
// each table reads its own rows.
//
// From your host application, before generating, just fill it:
//     docxform::tableContext()["\\tablewage"] = {
//         {"Иванов И. И.", "100000"},
//         {"Петров П. П.", "90000"},
//     };
// and have that tag's build() copy `tableContext()[tag]` into TableData::rows
// (see the runtimeTable() demo in tablekinds.cpp). Returns a reference to a
// process-wide store; change its value type here if you need richer data.
std::map<std::string, std::vector<std::vector<std::string>>>& tableContext();

// Runtime data your application feeds to TEXT builders — the text counterpart of
// tableContext(). Key it by the tag (e.g. "\\company") so each text tag reads its
// own value:
//     docxform::textContext()["\\company"] = "ООО «Ромашка»";
// and bind that tag to a build() that returns textContext()[tag] (see the
// contextText() demo in tablekinds.cpp). Returns a process-wide store.
std::map<std::string, std::string>& textContext();

// Render a TableData as an OOXML <w:tbl>...</w:tbl> fragment (with visible
// borders and centered cell text). Reusable on its own.
std::string buildTableXml(const TableData& table);

}  // namespace docxform

#endif  // DOCXFORM_TABLEKINDS_H
