// docxform.h - public API of the docxform GUI templater.
//
// This header lets you embed the docxform window inside another Qt application:
// you give it the path of a .docx template (with {{placeholders}} and optional
// \table{name} markers) and get back a ready-to-show widget that asks the user
// for the values and writes the filled .docx itself.
//
// To reuse it, link docxform.cpp (built WITH the macro DOCXFORM_NO_MAIN, so its
// own main() is dropped) together with tablekinds.cpp, and include this header:
//
//     #define DOCXFORM_NO_MAIN          // when compiling docxform.cpp
//
//     #include "docxform.h"
//     ...
//     QString err;
//     if (QWidget* form = docxform::openTemplateForm("contract.docx", &err))
//         form->show();                 // a QApplication must already exist
//     else
//         /* show err to the user */;
//
// Everything table-related stays in tablekinds.h / tablekinds.cpp, so you can
// customise the offered tables exactly as in the standalone tool.

#ifndef DOCXFORM_H
#define DOCXFORM_H

class QString;
class QWidget;

namespace docxform {

// Build the templating form for the .docx at `templatePath` and return it as a
// top-level QWidget. The caller OWNS the returned widget: call show() on it (or
// give it a parent, see below) and delete it when done — setting
// Qt::WA_DeleteOnClose is a convenient way to tie it to the window's lifetime.
//
// A QApplication (or QGuiApplication-based app) must already exist before this
// is called, exactly as for any other Qt widget.
//
// On failure (file missing, or not a valid .docx) the function returns nullptr
// and, if `error` is non-null, stores a human-readable message in it. It never
// pops up a dialog itself, so the embedding program stays in control of how
// errors are shown.
//
// If `parent` is non-null the window is parented to it (kept as a separate
// top-level window, but owned by the parent for stacking and lifetime).
QWidget* openTemplateForm(const QString& templatePath, QString* error = nullptr,
                          QWidget* parent = nullptr);

}  // namespace docxform

#endif  // DOCXFORM_H
