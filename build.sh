#!/bin/sh
# Your shell must be able to run g++ <https://gcc.gnu.org/> for Windows.
# I recommend you install MSYS2 <http://www.msys2.org>, run the mingw32.exe shell,
# in that shell run `pacman -S mingw-w64-i686-toolchain` and then this script.
# I did this on Windows 11; the resulting .exe worked on 32-bit Windows NT.

cd `dirname "$0"` || exit $?
exec g++ -static comProxy.cpp -o comProxy.exe -lWs2_32
