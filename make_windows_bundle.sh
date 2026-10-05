#!/bin/bash
# Assemble the Windows self-build bundle: a single zip the end user unpacks on
# a locked-down company PC (no Python, no admin rights, no PyPI access) to
# build the exe LOCALLY -- a locally built exe carries no Mark-of-the-Web, so
# download-blocking (SmartScreen / mail gateway) never sees it.
#
# Run on macOS:  bash make_windows_bundle.sh
# Output:        dist-windows/PDPC函釋工具-自建包.zip
#
# Contents assembled here:
#   WinPython/   portable CPython 3.13 (WinPython "dot" build; includes
#                tkinter, which python.org's embeddable build lacks), pruned
#   wheels/      offline win_amd64/cp313 wheels: requests, openpyxl, truststore,
#                pyinstaller + full dependency closure
#   app.py, core.py, requirements.txt, _build_exe.py
#   1-打包成exe.bat / 2-直接執行.bat / 說明.txt   (from bundle/)
set -euo pipefail
cd "$(dirname "$0")"

WINPYTHON_URL="https://github.com/winpython/winpython/releases/download/17.4.20260511final/WinPython64-3.13.13.0dot.zip"
WINPYTHON_DIR="WPy64-313130"          # top-level dir inside that zip
PY_TAG="3.13"                          # wheel ABI to download for
# Versions installed by the .bat files -- keep the three pins in sync with
# bundle/*.bat when bumping.
PKGS=(requests==2.34.2 openpyxl==3.1.5 truststore==0.10.4 pyinstaller==6.21.0)

WORK="dist-windows/bundle-build"
OUT_DIR="PDPC函釋工具-自建包"
OUT_ZIP="dist-windows/${OUT_DIR}.zip"
CACHE="dist-windows/winpython-cache.zip"

mkdir -p dist-windows
rm -rf "$WORK" "$OUT_ZIP"
mkdir -p "$WORK/$OUT_DIR"

# --- 1. WinPython: download (cached), extract, keep only python/, prune ----
if [ ! -f "$CACHE" ]; then
    echo "== Downloading WinPython..."
    # --fail: without it an HTTP error page is written to the cache path and
    # the `-f "$CACHE"` check above then reuses that poisoned cache forever.
    # Download to .part and rename, so an interrupted run leaves no cache.
    curl --fail -sL -o "$CACHE.part" "$WINPYTHON_URL"
    mv "$CACHE.part" "$CACHE"
fi
echo "== Extracting WinPython..."
unzip -q "$CACHE" -d "$WORK/wp"
mv "$WORK/wp/$WINPYTHON_DIR/python" "$WORK/$OUT_DIR/WinPython"
cp "$WORK/wp/$WINPYTHON_DIR/license.txt" "$WORK/$OUT_DIR/WinPython/"

echo "== Pruning WinPython (IDE launchers dropped with wp/; trimming python/)..."
(
    cd "$WORK/$OUT_DIR/WinPython"
    # Not needed to run tkinter apps, pip-install wheels, or run PyInstaller:
    #   idlelib/ensurepip (IDLE, bundled-pip installer), __pycache__ (regenerates),
    #   include+libs (C build files), WinPython's own tool packages, test .pyds,
    #   tix (legacy tk extension).
    rm -rf Lib/idlelib Lib/ensurepip Lib/__pycache__ include libs \
           Lib/site-packages/wppm* Lib/site-packages/sqlite_bro* \
           Lib/site-packages/sv_ttk* tcl/tix*
    rm -f DLLs/_test*.pyd DLLs/_ctypes_test.pyd \
          Scripts/sqlite_bro.exe Scripts/wppm.exe Scripts/pyproject-build.exe
)
rm -rf "$WORK/wp"

# tkinter must survive the prune -- the GUI depends on it.
for f in Lib/tkinter/__init__.py DLLs/_tkinter.pyd DLLs/tcl86t.dll \
         DLLs/tk86t.dll python.exe pythonw.exe; do
    [ -e "$WORK/$OUT_DIR/WinPython/$f" ] || { echo "!! missing $f"; exit 1; }
done

# --- 2. Offline wheels for win_amd64 ---------------------------------------
echo "== Downloading win_amd64 wheels..."
# pip evaluates environment markers against the RUNNING (mac) interpreter, so
# win32-only deps of pyinstaller are skipped and mac-only macholib is pulled
# in. Add the former explicitly; drop the latter.
./venv/bin/pip download --platform win_amd64 --python-version "$PY_TAG" \
    --implementation cp --only-binary=:all: -q \
    -d "$WORK/$OUT_DIR/wheels" "${PKGS[@]}" pefile pywin32-ctypes
rm -f "$WORK/$OUT_DIR/wheels"/macholib-*.whl
for dep in pefile pywin32_ctypes pyinstaller setuptools truststore; do
    ls "$WORK/$OUT_DIR/wheels"/${dep}-*.whl >/dev/null 2>&1 \
        || { echo "!! wheel missing: $dep"; exit 1; }
done

# --- 3. Sources + user-facing files ----------------------------------------
echo "== Copying sources and bundle files..."
cp app.py core.py requirements.txt bundle/_build_exe.py "$WORK/$OUT_DIR/"
# .bat files need CRLF (cmd.exe misparses GOTO/labels with bare LF); the
# repo copies stay LF for git. 說明.txt gets CRLF + BOM so Notepad renders it.
python3 - "$WORK/$OUT_DIR" <<'EOF'
import pathlib, sys
out = pathlib.Path(sys.argv[1])
for src in pathlib.Path("bundle").glob("*.bat"):
    text = src.read_text(encoding="utf-8").replace("\r\n", "\n")
    (out / src.name).write_bytes(text.replace("\n", "\r\n").encode("utf-8"))
text = pathlib.Path("bundle/說明.txt").read_text(encoding="utf-8").replace("\r\n", "\n")
(out / "說明.txt").write_bytes(b"\xef\xbb\xbf" + text.replace("\n", "\r\n").encode("utf-8"))
EOF

# --- 4. Zip + report --------------------------------------------------------
echo "== Zipping..."
# Python zipfile, NOT the zip CLI: macOS zip stores UTF-8 names without the
# UTF-8 flag (bit 0x800), so a Traditional-Chinese Windows (CP950) mangles
# every Chinese filename on extraction. zipfile sets the flag correctly.
python3 - "$WORK" "$OUT_DIR" "$OUT_ZIP" <<'EOF'
import os, sys, zipfile
work, out_dir, out_zip = sys.argv[1:]
with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for root, dirs, files in os.walk(os.path.join(work, out_dir)):
        dirs.sort()
        for f in sorted(files):
            full = os.path.join(root, f)
            z.write(full, os.path.relpath(full, work))
names = zipfile.ZipFile(out_zip).infolist()
bad = [i.filename for i in names
       if any(ord(c) > 127 for c in i.filename) and not (i.flag_bits & 0x800)]
assert bad == [], f"utf-8 flag missing on: {bad[:3]}"
print(f"   {len(names)} entries, "
      f"{sum(1 for i in names if i.filename.endswith('.whl'))} wheels")
EOF
unzip -tq "$OUT_ZIP" >/dev/null || { echo "!! zip integrity check failed"; exit 1; }
rm -rf "$WORK"

echo
echo "Done: $OUT_ZIP ($(du -h "$OUT_ZIP" | cut -f1 | tr -d ' '))"
