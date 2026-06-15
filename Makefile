# Makefile for docx2txt / txt2docx / docxform.
#
#   make            build the CLI tools (docx2txt, txt2docx)
#   make docx2txt   build only docx2txt
#   make txt2docx   build only txt2docx
#   make docxform   build the Qt5 GUI templater (needs qtbase5-dev)
#   make clean      remove the built binaries
#
# Override the compiler or flags from the command line if needed, e.g.:
#   make CXX=g++-9
#   make CXXFLAGS="-O3 -std=c++17"

CXX      ?= g++
CXXFLAGS ?= -O2 -std=c++17
LDLIBS   := -lz

# GCC 8 and older keep std::filesystem in a separate library and need
# -lstdc++fs at link time; GCC 9+ and Clang do not. Detect the GCC major
# version (kept paren-free so every make/shell parses it cleanly). Clang is
# skipped, since it does not ship libstdc++fs. Override LDLIBS to force it.
IS_CLANG   := $(shell $(CXX) -dM -E -x c++ /dev/null 2>/dev/null | grep __clang__)
GCC_MAJOR  := $(shell $(CXX) -dumpversion 2>/dev/null | cut -d. -f1)
ifeq ($(IS_CLANG),)
ifneq ($(GCC_MAJOR),)
ifeq ($(shell test $(GCC_MAJOR) -lt 9 2>/dev/null && echo yes),yes)
LDLIBS += -lstdc++fs
endif
endif
endif

# Qt5 flags for the GUI tool (docxform), resolved via pkg-config.
QT_CFLAGS := $(shell pkg-config --cflags Qt5Widgets 2>/dev/null)
QT_LIBS   := $(shell pkg-config --libs Qt5Widgets 2>/dev/null)

# CLI tools are built by default; docxform is opt-in (needs Qt5).
CLI_BINARIES := docx2txt txt2docx
BINARIES     := $(CLI_BINARIES) docxform

.PHONY: all clean
all: $(CLI_BINARIES)

docx2txt: docx2txt.cpp
	$(CXX) $(CXXFLAGS) -o $@ $< $(LDLIBS)

txt2docx: txt2docx.cpp
	$(CXX) $(CXXFLAGS) -o $@ $< $(LDLIBS)

docxform: docxform.cpp tablekinds.cpp tablekinds.h
	@if [ -z "$(QT_LIBS)" ]; then \
	  echo "Qt5 not found. Install it, e.g.: sudo apt install qtbase5-dev"; \
	  exit 1; \
	fi
	$(CXX) $(CXXFLAGS) -fPIC $(QT_CFLAGS) -o $@ docxform.cpp tablekinds.cpp \
	    $(QT_LIBS) $(LDLIBS)

clean:
	rm -f $(BINARIES)
