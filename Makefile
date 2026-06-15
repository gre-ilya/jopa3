# Makefile for docx2txt / txt2docx.
#
#   make            build both tools
#   make docx2txt   build only docx2txt
#   make txt2docx   build only txt2docx
#   make clean      remove the built binaries
#
# Override the compiler or flags from the command line if needed, e.g.:
#   make CXX=g++-9
#   make CXXFLAGS="-O3 -std=c++17"

CXX      ?= g++
CXXFLAGS ?= -O2 -std=c++17
LDLIBS   := -lz

# GCC 8 and older (and old Clang) keep std::filesystem in a separate library,
# so they need -lstdc++fs at link time. Detect this by trying to link a tiny
# std::filesystem program without the flag: if that fails, add the flag.
# Newer toolchains link fine without it, so it is omitted there.
FS_CHECK := $(shell t=$$(mktemp --suffix=.cpp 2>/dev/null || echo /tmp/.fscheck.cpp); \
	printf '#include <filesystem>\nint main(){return std::filesystem::exists("/")?0:1;}\n' > $$t; \
	if $(CXX) $(CXXFLAGS) $$t -o /dev/null >/dev/null 2>&1; then echo no; else echo yes; fi; \
	rm -f $$t)
ifeq ($(FS_CHECK),yes)
LDLIBS += -lstdc++fs
endif

BINARIES := docx2txt txt2docx

.PHONY: all clean
all: $(BINARIES)

docx2txt: docx2txt.cpp
	$(CXX) $(CXXFLAGS) -o $@ $< $(LDLIBS)

txt2docx: txt2docx.cpp
	$(CXX) $(CXXFLAGS) -o $@ $< $(LDLIBS)

clean:
	rm -f $(BINARIES)
