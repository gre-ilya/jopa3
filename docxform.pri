# docxform.pri — include this from your qmake .pro to embed docxform as a
# third-party component:
#
#     include(path/to/docxform/docxform.pri)
#
# It adds the docxform + tablekinds sources/headers, requires the Qt Widgets
# module, links zlib and defines DOCXFORM_NO_MAIN so docxform's own main() is
# dropped (your application keeps its own). After including it, call
# docxform::openTemplateForm() from your code (see docxform.h).
#
# Cross-platform: builds with qmake on Linux, macOS and Windows. The code itself
# uses only standard C++17, Qt and zlib (no OS-specific headers); the ZIP/.docx
# byte handling is endianness-independent and file access goes through QFile, so
# non-ASCII (e.g. Cyrillic) paths work on Windows too.

QT += widgets
CONFIG += c++17

INCLUDEPATH += $$PWD

HEADERS += \
    $$PWD/docxform.h \
    $$PWD/tablekinds.h

SOURCES += \
    $$PWD/docxform.cpp \
    $$PWD/tablekinds.cpp

DEFINES += DOCXFORM_NO_MAIN

# zlib. On Linux/macOS and MinGW the system zlib links with -lz. MSVC has no
# default zlib, so provide one yourself and point the build at it — for example:
#   win32-msvc* {
#       INCLUDEPATH += C:/path/to/zlib/include
#       LIBS += -LC:/path/to/zlib/lib -lzlib   # or zlibstatic.lib / zdll.lib
#   }
win32-msvc* {
    # Supply zlib for MSVC here (see the note above); -lz does not apply.
} else {
    LIBS += -lz
}
