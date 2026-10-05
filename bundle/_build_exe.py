# -*- coding: utf-8 -*-
"""Build the Windows exe with PyInstaller. Called by 1-打包成exe.bat.

The Chinese exe name lives HERE rather than in the .bat: a Python source file
is reliably UTF-8, while cmd.exe's codepage handling of non-ASCII command-line
arguments after `chcp 65001` varies across Windows builds. Keeping the .bat
pure-ASCII sidesteps that entirely.

Flags mirror build_windows.bat, so the locally built exe is identical in
behaviour to one built elsewhere -- minus the Mark-of-the-Web, which is the
whole point of building on the user's machine.
"""
import importlib.util

import PyInstaller.__main__

args = [
    "--name", "PDPC函釋更新工具",
    "--windowed",
    "--onefile",
    "--noconfirm",
]

# truststore picks its backend behind a platform test, so the Windows submodule
# is invisible to PyInstaller's static analysis -- collect the package whole.
# It is an OPTIONAL dependency (core.py runs without it), and an older extracted
# bundle carries no truststore wheel, so only ask for it when it is installed --
# otherwise PyInstaller would fail on a module it cannot find.
if importlib.util.find_spec("truststore") is not None:
    args += ["--collect-submodules", "truststore"]
else:
    print("!! truststore not installed -- building without OS certificate "
          "store support (the tool still works; see 說明.txt Q4).")

PyInstaller.__main__.run(args + ["app.py"])
