# CI: a minimal host app that embeds docxform via docxform.pri. Building it
# validates the .pri (sources, QT += widgets, zlib linking, DOCXFORM_NO_MAIN) on
# each platform/compiler.
TEMPLATE = app
TARGET   = embedhost
CONFIG  += c++17 console
CONFIG  -= app_bundle

SOURCES += $$PWD/embed_main.cpp

include($$PWD/../docxform.pri)
