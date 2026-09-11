#!/usr/bin/env python3
"""Create a double-clickable launcher for the ROI Model app.

    python make_launcher.py

Windows Smart App Control refuses the packaged exe: it is unsigned and has no
reputation, so Code Integrity blocks it outright ("did not meet the Enterprise
signing level requirements"). Signing it properly needs a paid code-signing
certificate, and turning Smart App Control off is irreversible without
reinstalling Windows.

pythonw.exe, though, is signed by the Python Software Foundation and is
trusted, so a shortcut that runs the script through it launches normally -
no console window, no security downgrade, nothing to approve. This writes
that shortcut, with an icon so it reads as an app rather than a script.

Re-run it if you move the folder; the shortcut stores absolute paths.
"""

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
APP_NAME = "ROI Model"
ICON = HERE / "roi_model.ico"
SCRIPT = HERE / "roi_model.py"

ACCENT = (5, 150, 105, 255)      # the app's emerald
WHITE = (255, 255, 255, 255)


def make_icon(path):
    """An emerald tile with the app's ascending bars, at every icon size."""
    from PIL import Image, ImageDraw

    def tile(size):
        ss = 8 if size <= 64 else 2          # supersample small sizes hardest
        im = Image.new("RGBA", (size * ss, size * ss), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        s = size * ss
        d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.22),
                            fill=ACCENT)
        # three rising bars, as in the net-cash chart
        pad, gap = s * 0.22, s * 0.06
        bw = (s - 2 * pad - 2 * gap) / 3
        for i, h in enumerate((0.26, 0.42, 0.56)):
            x = pad + i * (bw + gap)
            top = s - pad - s * h
            d.rounded_rectangle([x, top, x + bw, s - pad],
                                radius=int(bw * 0.28), fill=WHITE)
        return im.resize((size, size), Image.LANCZOS)

    sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
    base = tile(256)
    base.save(path, format="ICO",
              sizes=[(n, n) for n in sizes],
              append_images=[tile(n) for n in sizes if n != 256])
    return path


def make_shortcut(target_dir):
    """A .lnk that runs the script through the signed pythonw.exe."""
    import win32com.client

    pythonw = Path(sys.executable).with_name("pythonw.exe")
    if not pythonw.exists():                 # a console build will do, noisily
        pythonw = Path(sys.executable)

    link_path = Path(target_dir) / f"{APP_NAME}.lnk"
    shell = win32com.client.Dispatch("WScript.Shell")
    link = shell.CreateShortCut(str(link_path))
    link.TargetPath = str(pythonw)
    link.Arguments = f'"{SCRIPT}"'
    link.WorkingDirectory = str(HERE)
    link.Description = f"{APP_NAME} - six-year contract cash-flow model"
    if ICON.exists():
        link.IconLocation = f"{ICON},0"
    link.save()
    return link_path


def main():
    if sys.platform != "win32":
        sys.exit("This launcher is for Windows. On macOS run "
                 "'python3 build.py' for a .app, or 'python3 roi_model.py'.")
    if not SCRIPT.exists():
        sys.exit(f"Cannot find {SCRIPT}")

    make_icon(ICON)
    print(f"icon      {ICON}")
    here = make_shortcut(HERE)
    print(f"shortcut  {here}")

    desktop = Path.home() / "Desktop"
    if desktop.is_dir():
        there = make_shortcut(desktop)
        print(f"shortcut  {there}")

    print(f"\nDouble-click \"{APP_NAME}\" to run it. It goes through "
          f"pythonw.exe,\nwhich is signed, so Smart App Control allows it.")


if __name__ == "__main__":
    main()
