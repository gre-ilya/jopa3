# docxform.pri — include this from your qmake .pro to embed docxform as a
# third-party component:
#
#     include(path/to/docxform/docxform.pri)
#
# It adds the docxform + tablekinds sources/headers, requires the Qt Widgets
# module, links zlib and defines DOCXFORM_NO_MAIN so docxform's own main() is
# dropped (your application keeps its own). After including it, call
# docxform::openTemplateForm() from your code (see docxform.h).

QT += widgets

INCLUDEPATH += $$PWD

HEADERS += \
    $$PWD/docxform.h \
    $$PWD/tablekinds.h

SOURCES += \
    $$PWD/docxform.cpp \
    $$PWD/tablekinds.cpp

DEFINES += DOCXFORM_NO_MAIN

LIBS += -lz
