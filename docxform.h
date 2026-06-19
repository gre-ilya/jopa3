// docxform.h - public API of the docxform GUI templater.
//
// This header lets you embed the docxform window inside another Qt application:
// you give it the path of a .docx template (with \var{...} variables and
// optional \variant{...} choices and fixed table tags like \tablewage) and get
// back a ready-to-show widget that asks for the values and writes the .docx.
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

// Convenience entry point for embedding: reproduce the WHOLE standalone GUI flow
// from a host application with one call. It pops a file chooser for a .docx
// template, builds the form and shows it as a top-level window (deleted on
// close); on error it shows a message box itself. Returns the shown window, or
// nullptr if the user cancelled the chooser or opening failed.
//
// This is exactly what the standalone program does when launched without a file
// argument, so you can wire it straight to a button's clicked() slot to start
// the module "as if launched on its own":
//
//     connect(myButton, &QPushButton::clicked, this,
//             [this]{ docxform::showTemplateForm(this); });
//
// A QApplication must already exist (your app has one). `parent`, if given, owns
// the dialog and the window for stacking and lifetime.
QWidget* showTemplateForm(QWidget* parent = nullptr);

}  // namespace docxform

#endif  // DOCXFORM_H
