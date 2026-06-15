# Makefile for docxform — the Qt5 GUI .docx templater.
#
#   make            build docxform (needs qtbase5-dev + pkg-config)
#   make docxform   same as above
#   make clean      remove the built binary
#
# Override the compiler or flags from the command line if needed, e.g.:
#   make CXX=g++-9
#   make CXXFLAGS="-O3 -std=c++17"

CXX      ?= g++
CXXFLAGS ?= -O2 -std=c++17
LDLIBS   := -lz

# Qt5 flags for docxform, resolved via pkg-config.
QT_CFLAGS := $(shell pkg-config --cflags Qt5Widgets 2>/dev/null)
QT_LIBS   := $(shell pkg-config --libs Qt5Widgets 2>/dev/null)

.PHONY: all clean
all: docxform

docxform: docxform.cpp docxform.h tablekinds.cpp tablekinds.h
	@if [ -z "$(QT_LIBS)" ]; then \
	  echo "Qt5 not found. Install it, e.g.: sudo apt install qtbase5-dev"; \
	  exit 1; \
	fi
	$(CXX) $(CXXFLAGS) -fPIC $(QT_CFLAGS) -o $@ docxform.cpp tablekinds.cpp \
	    $(QT_LIBS) $(LDLIBS)

clean:
	rm -f docxform
