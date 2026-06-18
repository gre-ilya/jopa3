// CI build/link check for the embedding API. Compiled (not run) with
// docxform.pri to prove the module embeds and links on each platform/compiler.
#include <QApplication>

#include "docxform.h"

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    QString err;
    // Exercise the public API so both entry points must compile and link.
    QWidget* w = docxform::openTemplateForm(QString::fromUtf8("none.docx"), &err);
    (void)w;
    (void)&docxform::showTemplateForm;  // force showTemplateForm to be linked
    return 0;
}
