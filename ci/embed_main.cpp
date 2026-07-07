// CI build/link check for the embedding API. Compiled (not run) with
// docxform.pri to prove the module embeds and links on each platform/compiler.
#include <QApplication>
#include <QFile>

#include "docxform.h"

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    QString err;
    // Exercise the public API so both entry points must compile and link.
    bool ok = docxform::renderTemplate(QString::fromUtf8("none.docx"),
                                       QString::fromUtf8("out.docx"),
                                       /*highlight=*/true, &err);
    (void)ok;
    // Force every fillTemplate overload to be linked (disambiguated by cast).
    (void)static_cast<bool (*)(bool, QWidget*)>(&docxform::fillTemplate);
    (void)static_cast<bool (*)(const QString&, bool, QWidget*)>(
        &docxform::fillTemplate);
    (void)static_cast<bool (*)(QFile&, bool, QWidget*)>(&docxform::fillTemplate);
    return 0;
}
