#!/usr/bin/env python3
"""Build the ROI Model desktop app for whichever platform you run this on.

    python build.py

PyInstaller does not cross-compile, so run this on macOS to get the .app and
on Windows to get the .exe.

  macOS    ->  dist/ROI Model.app                  (a bundle, as Apple expects)
  Windows  ->  dist/ROI Model/ROI Model.exe        (exe plus its _internal folder)
  Linux    ->  dist/ROI Model/ROI Model

Every target is a folder build. A --onefile exe unpacks itself into %TEMP% and
runs from there on each launch, and Windows Smart App Control blocks that
outright - "An Application Control policy has blocked this file" - because it
is indistinguishable from self-extracting malware. Loading from a fixed folder
beside the exe is permitted, so the app runs unsigned with Smart App Control
left on. Keep the folder together; zip it to share.

Only Pillow is needed at runtime; everything else below is excluded to keep
the build small - the Dash front end's dependencies in particular.
"""

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
APP_NAME = "ROI Model"
BUNDLE_ID = "com.roimodel.desktop"

# Pulled in by other tooling in this project, never by the desktop app.
# openpyxl is a runtime dependency now - the export writes a live .xlsx -
# so it must stay in the bundle.
EXCLUDES = ["numpy", "scipy", "matplotlib", "pandas",
            "dash", "plotly", "flask", "werkzeug", "pydantic", "IPython",
            "PIL.ImageQt", "tkinter.test", "test"]


def main():
    if sys.platform == "darwin":
        target = f"dist/{APP_NAME}.app"
    elif sys.platform == "win32":
        target = f"dist/{APP_NAME}/{APP_NAME}.exe"
    else:
        target = f"dist/{APP_NAME}/{APP_NAME}"

    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--windowed",
        "--name", APP_NAME,
        "--distpath", str(HERE / "dist"),
        "--workpath", str(HERE / "build"),
        "--specpath", str(HERE / "build"),
    ]

    icon = HERE / ("roi_model.icns" if sys.platform == "darwin"
                   else "roi_model.ico")
    if icon.exists():
        args += ["--icon", str(icon)]

    if sys.platform == "darwin":
        args += ["--osx-bundle-identifier", BUNDLE_ID]

    for module in EXCLUDES:
        args += ["--exclude-module", module]

    args.append(str(HERE / "roi_model.py"))

    print("Building", target, "\n")
    result = subprocess.run(args, cwd=HERE)
    if result.returncode != 0:
        sys.exit(f"\nBuild failed (exit {result.returncode}).")

    print(f"\nBuilt {target}")
    if sys.platform == "darwin":
        print(
            "\nThe bundle is unsigned, so Gatekeeper will block a double-click\n"
            "the first time. Either right-click the app and choose Open, or:\n"
            f'    xattr -dr com.apple.quarantine "{target}"\n'
            "To share it more widely, sign and notarise it with your Apple\n"
            "Developer ID:\n"
            f'    codesign --deep --force --options runtime -s "Developer ID '
            f'Application: YOUR NAME" "{target}"'
        )


if __name__ == "__main__":
    main()
