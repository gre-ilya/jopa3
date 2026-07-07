// docxform.h - public API of the docxform .docx templater.
//
// The module fills a .docx template from YOUR OWN tags (defined in
// tablekinds.cpp): fixed table tags like \tablewage and inline text tags like
// \company. There is no interactive per-variable form — calling the module pops
// a file chooser for the template, then immediately a save dialog for the output,
// and writes the generated document.
//
// To reuse it, link docxform.cpp (built WITH the macro DOCXFORM_NO_MAIN, so its
// own main() is dropped) together with tablekinds.cpp, and include this header:
//
//     #define DOCXFORM_NO_MAIN          // when compiling docxform.cpp
//
//     #include "docxform.h"
//     ...
//     // Highlight the inserted text and table contents yellow? -> bool argument.
//     docxform::fillTemplate(/*highlight=*/true, this);   // a QApplication must exist
//
// Everything table/text-related stays in tablekinds.h / tablekinds.cpp, so you
// customise the tags exactly as in the standalone tool.

#ifndef DOCXFORM_H
#define DOCXFORM_H

class QString;
class QWidget;

namespace docxform {

// THE MODULE'S MAIN ENTRY POINT. Runs the whole interactive flow:
//   1. pops a file chooser for the .docx template;
//   2. immediately pops a save dialog for the output path;
//   3. expands every fixed table/text tag and writes the generated .docx.
//
// `highlight` decides whether the inserted text AND the generated table contents
// are highlighted yellow (true) or inserted without any highlight (false) — it is
// the single knob the host application controls.
//
// Returns true if a document was written, false if the user cancelled either
// dialog or an error occurred (an error is shown to the user via a message box).
// A QApplication must already exist. `parent`, if given, owns the dialogs for
// stacking and lifetime.
bool fillTemplate(bool highlight, QWidget* parent = nullptr);

// Same as above, but the template is passed in directly (`templatePath`), so the
// template chooser is SKIPPED — only the save dialog is shown before generating.
// Use this when your application already knows which template to fill. Returns
// false if `templatePath` is empty, the user cancelled the save dialog, or an
// error occurred (shown via a message box). A QApplication must already exist.
bool fillTemplate(const QString& templatePath, bool highlight,
                  QWidget* parent = nullptr);

// Headless core used by fillTemplate() (and the --render command line): read the
// template at `templatePath`, expand every fixed table/text tag, and write the
// result to `outPath`. `highlight` highlights the inserted text and table
// contents yellow when true. Shows no dialogs. On failure returns false and, if
// `error` is non-null, stores a human-readable message in it. Handy for scripting
// or when the host already knows both paths.
bool renderTemplate(const QString& templatePath, const QString& outPath,
                    bool highlight, QString* error = nullptr);

}  // namespace docxform

#endif  // DOCXFORM_H
