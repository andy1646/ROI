"""Hotel ROI Modeling Tool - the desktop front end for the model in roi_core.py.

Tk draws no antialiased primitives, so every rounded surface here - cards,
slider tracks and knobs, buttons, bar caps - is rendered by Pillow at 4x and
downsampled. Sprites are cached by geometry, and the pieces that move during a
drag are static images: the slider repositions them rather than redrawing.
"""

import json
import sys
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageDraw, ImageFilter, ImageTk

from roi_core import (CATEGORIES, DEFAULTS, DEFAULT_TERM, MAX_YEARS, ROWS,
                      ONLY_YEAR0, STAFF_TITLES, STAFF_TYPES, TERM_STEP,
                      YEAR0, term_weights, year_count,
                      as_years, compute, is_uniform, lead_value, money,
                      pct, staff_base)
from roi_export import write_workbook

APP_TITLE = "Hotel ROI Modeling Tool"
SCENARIOS = 3  # saved parameter sets you can flip between

# Marker written into saved setup files so a stray .json is not mistaken
# for one. Bump the version if the snapshot shape ever changes.
SETUP_FORMAT = "roi-model-setup"
SETUP_VERSION = 1

# --------------------------------------------------------------- geometry ---
# Pixel scale for the active display, filled in once the root window exists.
# Every hard-coded pixel size goes through px() so the layout holds up on
# 125% / 150% / 4K displays instead of being bitmap-stretched by Windows.
K = 1.0
SS = 4  # supersampling factor for Pillow-rendered art


def px(n):
    return int(round(n * K))


def _declare_dpi_aware():
    """Tell Windows we scale ourselves; must run before the root window."""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()   # pre-8.1 fallback
        except (AttributeError, OSError):
            pass


def _wheel_steps(e):
    """Wheel events differ per platform: Windows sends multiples of 120,
    macOS small signed counts, and X11 button 4/5 presses."""
    if getattr(e, "num", None) in (4, 5):
        return -1 if e.num == 4 else 1
    if sys.platform == "darwin":
        return -e.delta
    return int(-e.delta / 120)


# ---------------------------------------------------------------- palette ---
PAGE = "#eef1f5"        # app background
CARD = "#ffffff"
HERO = "#101828"        # deep slate hero panel
HERO_2 = "#1d2939"      # hairline inside the hero

INK = "#0f172a"
INK_2 = "#334155"
MUTED = "#8794a7"
MUTED_2 = "#64748b"
HAIR = "#e8ecf1"

ACCENT = "#059669"      # emerald 600
ACCENT_2 = "#10b981"    # emerald 500, chart bars
ACCENT_HI = "#34d399"   # emerald 400, on dark
ACCENT_BG = "#ecfdf5"
NEG = "#e11d48"
NEG_2 = "#f43f5e"

TRACK = "#e6eaf0"
KNOB_EDGE = "#cbd5e1"
SUBCARD = "#f8fafc"


# Weight variants are separate families on Windows but not elsewhere, so the
# family for each weight is resolved against what the system actually has,
# falling back to the base family (bold-flagged for semibold) when needed.
_FONT_PREFS = {
    "":   ["Segoe UI Variable Text", "Segoe UI", "SF Pro Text", "Helvetica Neue",
           "Lucida Grande", "DejaVu Sans"],
    "sb": ["Segoe UI Semibold", "SF Pro Text Semibold", "Helvetica Neue Medium"],
    "sl": ["Segoe UI Semilight", "SF Pro Display Light", "Helvetica Neue Light"],
    "lt": ["Segoe UI Light", "SF Pro Display Light", "Helvetica Neue Light"],
}
_FONTS = {k: ("TkDefaultFont",) for k in _FONT_PREFS}


def resolve_fonts(root):
    """Pick real font families once the root window exists."""
    import tkinter.font as tkfont
    have = set(tkfont.families(root))

    def pick(names):
        return next((n for n in names if n in have), None)

    base = pick(_FONT_PREFS[""]) or "TkDefaultFont"
    _FONTS[""] = (base,)
    for weight in ("sb", "sl", "lt"):
        fam = pick(_FONT_PREFS[weight])
        if fam:
            _FONTS[weight] = (fam,)
        elif weight == "sb":
            _FONTS[weight] = (base, "bold")   # no semibold family: use bold
        else:
            _FONTS[weight] = (base,)          # no light family: use regular


def F(size, w=""):
    """The UI font at a given weight: '' regular, 'sb' semibold, 'sl' semilight."""
    spec = _FONTS[w]
    return (spec[0], size) if len(spec) == 1 else (spec[0], size, spec[1])


def caps(s):
    """Fake letter-spacing for the small uppercase labels."""
    return " ".join(s.upper())


def _rgba(h, a=255):
    h = h.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), a)


# ============================================================ sprite cache ===
_SPRITES = {}


def sprite(key, build):
    """Cache a PhotoImage by key; PhotoImages must outlive their canvas item."""
    img = _SPRITES.get(key)
    if img is None:
        img = _SPRITES[key] = ImageTk.PhotoImage(build())
    return img


def _round_rect(w, h, r, fill):
    """An antialiased rounded rectangle, drawn big then shrunk."""
    im = Image.new("RGBA", (w * SS, h * SS), (0, 0, 0, 0))
    ImageDraw.Draw(im).rounded_rectangle(
        [0, 0, w * SS - 1, h * SS - 1], radius=r * SS, fill=fill)
    return im.resize((w, h), Image.LANCZOS)


_PROTOS = {}


def _card_proto(radius, pad, fill, shadow, alpha, blur, drop):
    """A small supersampled card, big enough to hold all four corners."""
    key = (radius, pad, fill, shadow, alpha, blur, drop)
    got = _PROTOS.get(key)
    if got:
        return got
    c = radius + pad + 2
    s = c * 2
    im = Image.new("RGBA", (s * SS, s * SS), (0, 0, 0, 0))
    box = [pad * SS, pad * SS, (s - pad) * SS, (s - pad) * SS]
    if shadow:
        sh = Image.new("RGBA", im.size, (0, 0, 0, 0))
        ImageDraw.Draw(sh).rounded_rectangle(
            [box[0], box[1] + drop * SS, box[2], box[3] + drop * SS],
            radius=radius * SS, fill=(16, 24, 40, alpha))
        im = Image.alpha_composite(
            im, sh.filter(ImageFilter.GaussianBlur(blur * SS / 2)))
    ImageDraw.Draw(im).rounded_rectangle(box, radius=radius * SS, fill=fill)
    got = _PROTOS[key] = (im.resize((s, s), Image.LANCZOS), c)
    return got


def card_image(w, h, radius, fill, pad, shadow=True, alpha=34, blur=5, drop=3):
    """Rounded card with a soft drop shadow, inset by `pad` to leave blur room.

    Assembled as a 9-slice from a small cached prototype. Supersampling a
    full-size card and blurring it costs ~260 ms for the tall panels, which is
    unusable while a window is being laid out or dragged.
    """
    proto, c = _card_proto(radius, pad, fill, shadow, alpha, blur, drop)
    if w < 2 * c or h < 2 * c:            # too small to slice; draw it directly
        im = Image.new("RGBA", (w * SS, h * SS), (0, 0, 0, 0))
        ImageDraw.Draw(im).rounded_rectangle(
            [pad * SS, pad * SS, (w - pad) * SS, (h - pad) * SS],
            radius=radius * SS, fill=fill)
        return im.resize((w, h), Image.LANCZOS)

    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    mw, mh = w - 2 * c, h - 2 * c         # span of the stretched middles

    im.paste(proto.crop((0, 0, c, c)), (0, 0))                    # corners
    im.paste(proto.crop((c, 0, 2 * c, c)), (w - c, 0))
    im.paste(proto.crop((0, c, c, 2 * c)), (0, h - c))
    im.paste(proto.crop((c, c, 2 * c, 2 * c)), (w - c, h - c))

    if mw:                                                        # edges
        im.paste(proto.crop((c - 1, 0, c, c)).resize((mw, c)), (c, 0))
        im.paste(proto.crop((c - 1, c, c, 2 * c)).resize((mw, c)), (c, h - c))
    if mh:
        im.paste(proto.crop((0, c - 1, c, c)).resize((c, mh)), (0, c))
        im.paste(proto.crop((c, c - 1, 2 * c, c)).resize((c, mh)), (w - c, c))
    if mw and mh:                                                 # solid centre
        im.paste(Image.new("RGBA", (mw, mh), fill), (c, c))
    return im


def knob_image(d, ring=None):
    """White slider knob with a soft shadow, plus an optional accent halo."""
    pad = max(4, d // 3)
    s = d + pad * 2
    im = Image.new("RGBA", (s * SS, s * SS), (0, 0, 0, 0))
    ImageDraw.Draw(im).ellipse(
        [pad * SS, (pad + 1) * SS, (pad + d) * SS, (pad + d + 1) * SS],
        fill=(16, 24, 40, 70))
    im = im.filter(ImageFilter.GaussianBlur(1.6 * SS))
    dr = ImageDraw.Draw(im)
    if ring:
        dr.ellipse([(pad - 3) * SS, (pad - 3) * SS,
                    (pad + d + 3) * SS, (pad + d + 3) * SS], fill=ring)
    dr.ellipse([pad * SS, pad * SS, (pad + d) * SS, (pad + d) * SS],
               fill=(255, 255, 255, 255), outline=_rgba(KNOB_EDGE),
               width=max(1, SS // 2))
    return im.resize((s, s), Image.LANCZOS)


def cap_image(r, h, colour, side):
    """One rounded end of a track: a pill, cropped to the half we need."""
    pill = _round_rect(r * 2, h, min(r, h // 2), _rgba(colour))
    return pill.crop((0, 0, r, h)) if side == "l" else pill.crop((r, 0, r * 2, h))


def bar_cap_image(w, r, colour, up=True):
    """Rounded top (or bottom) of a chart bar, composited over the white card."""
    im = Image.new("RGBA", (w * SS, r * SS), (0, 0, 0, 0))
    box = ([0, 0, w * SS - 1, r * SS * 2] if up
           else [0, -r * SS, w * SS - 1, r * SS - 1])
    ImageDraw.Draw(im).rounded_rectangle(box, radius=r * SS, fill=_rgba(colour))
    return im.resize((w, r), Image.LANCZOS)


# =================================================================== card ====
class Card(tk.Frame):
    """A rounded, softly shadowed panel hosting an ordinary Frame.

    The background is a `place`d Label behind the content, so it tracks the
    card's size without ever contributing to it. Sizing the background from
    inside the geometry it depends on is a feedback loop; this avoids it.
    """

    def __init__(self, parent, bg=CARD, radius=None, pad=18, page=PAGE,
                 shadow=True):
        super().__init__(parent, bg=page)
        self.radius = radius if radius is not None else px(14)
        self.fill = _rgba(bg)
        self.shadow = shadow
        self.m = px(7)                      # room for the shadow
        self._photo = None
        self._last = None

        self.back = tk.Label(self, bd=0, bg=page, highlightthickness=0)
        self.back.place(x=0, y=0, relwidth=1, relheight=1)

        self.body = tk.Frame(self, bg=bg)
        inset = self.m + px(pad)
        self.body.pack(fill="both", expand=True, padx=inset, pady=inset)

        self.bind("<Configure>", self._on_resize)

    def _on_resize(self, e):
        size = (e.width, e.height)
        if size == self._last or e.width < 4 or e.height < 4:
            return
        self._last = size
        self._photo = ImageTk.PhotoImage(
            card_image(e.width, e.height, self.radius, self.fill, self.m,
                       self.shadow))
        self.back.configure(image=self._photo)


# ================================================================ button ====
class PillButton(tk.Canvas):
    def __init__(self, parent, text, command, primary=False, page=PAGE):
        self.h = px(34)
        self.primary = primary
        self.command = command
        probe = tk.Label(parent, text=text, font=F(10, "sb"))
        self.w = probe.winfo_reqwidth() + px(26)
        probe.destroy()
        super().__init__(parent, width=self.w, height=self.h, bg=page,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._img = self.create_image(0, 0, anchor="nw")
        self._txt = self.create_text(self.w // 2, self.h // 2, text=text,
                                     font=F(10, "sb"))
        self._paint(False)
        self.bind("<Enter>", lambda e: self._paint(True))
        self.bind("<Leave>", lambda e: self._paint(False))
        self.bind("<Button-1>", lambda e: self.command())

    def _paint(self, hot):
        if self.primary:
            fill, fg = (ACCENT if not hot else "#047857"), "#ffffff"
        else:
            fill, fg = (CARD if not hot else "#f5f7fa"), INK_2
        img = sprite(("btn", self.w, self.h, fill),
                     lambda: card_image(self.w, self.h, px(9), _rgba(fill),
                                        px(4), True, 26, 4, 2))
        self.itemconfig(self._img, image=img)
        self.itemconfig(self._txt, fill=fg)


class ScenarioButton(tk.Canvas):
    """A scenario tab: its name over the ROI that scenario currently yields,
    so all three can be compared without switching between them."""

    def __init__(self, parent, label, command, page=PAGE):
        self.w, self.h = px(100), px(44)
        super().__init__(parent, width=self.w, height=self.h, bg=page,
                         highlightthickness=0, bd=0, cursor="hand2")
        self.command = command
        self.active = False
        self._img = self.create_image(0, 0, anchor="nw")
        self._name = self.create_text(self.w // 2, self.h // 2 - px(8),
                                      text=caps(label), font=F(7, "sb"))
        self._roi = self.create_text(self.w // 2, self.h // 2 + px(7),
                                     text="-", font=F(13, "sb"))
        self._paint(False)
        self.bind("<Enter>", lambda e: self._paint(True))
        self.bind("<Leave>", lambda e: self._paint(False))
        self.bind("<Button-1>", lambda e: self.command())

    def set_roi(self, text):
        self.itemconfig(self._roi, text=text)

    def set_active(self, on):
        self.active = on
        self._paint(False)

    def _paint(self, hot):
        if self.active:
            fill, name_fg, roi_fg = ACCENT, "#a7f3d0", "#ffffff"
        else:
            fill = CARD if not hot else "#f5f7fa"
            name_fg, roi_fg = MUTED, INK_2
        img = sprite(("scen", self.w, self.h, fill),
                     lambda: card_image(self.w, self.h, px(9), _rgba(fill),
                                        px(4), True, 26, 4, 2))
        self.itemconfig(self._img, image=img)
        self.itemconfig(self._name, fill=name_fg)
        self.itemconfig(self._roi, fill=roi_fg)


# ============================================================= pill entry ====
class PillEntry(tk.Canvas):
    """A borderless entry sitting on a soft rounded field."""

    def __init__(self, parent, width, on_commit, bg=CARD, field="#f1f4f8",
                 justify="right", weight="sb"):
        self.h = px(27)
        self.w = px(width)
        super().__init__(parent, width=self.w, height=self.h, bg=bg,
                         highlightthickness=0, bd=0)
        img = sprite(("field", self.w, self.h, field),
                     lambda: _round_rect(self.w, self.h, px(7), _rgba(field)))
        self.create_image(0, 0, image=img, anchor="nw")
        self.entry = tk.Entry(self, bd=0, highlightthickness=0, bg=field,
                              fg=INK, font=F(10, weight), justify=justify)
        self.create_window(self.w - px(9), self.h // 2, window=self.entry,
                           anchor="e", width=self.w - px(18), height=px(17))
        self.entry.bind("<Return>", on_commit)
        self.entry.bind("<FocusOut>", on_commit)

    def text(self):
        return self.entry.get()

    def set_text(self, s):
        self.entry.delete(0, "end")
        self.entry.insert(0, s)


# ============================================================== note field ===
class NoteEntry(tk.Frame):
    """A free-text label in the header: click it and type anything, up to
    LIMIT characters. Shows a hint while empty, and underlines in the accent
    colour while focused so it reads as editable rather than as a caption."""

    LIMIT = 100

    def __init__(self, parent, placeholder="Click to add a note", chars=40):
        super().__init__(parent, bg=PAGE)
        self.placeholder = placeholder
        self._empty = True
        self.var = tk.StringVar()

        allow = self.register(lambda proposed: len(proposed) <= self.LIMIT)
        self.entry = tk.Entry(self, textvariable=self.var, width=chars,
                              bd=0, highlightthickness=0, bg=PAGE, fg=MUTED,
                              font=F(13), insertbackground=INK_2,
                              validate="key", validatecommand=(allow, "%P"))
        self.entry.pack(fill="x", ipady=px(3))
        self.rule = tk.Frame(self, bg=HAIR, height=1)
        self.rule.pack(fill="x")

        self.var.set(placeholder)
        self.entry.bind("<FocusIn>", self._focus)
        self.entry.bind("<FocusOut>", self._blur)
        self.entry.bind("<Escape>", lambda e: self.entry.selection_clear()
                        or self.focus_set())
        self.entry.bind("<Return>", lambda e: self.focus_set())

    def _focus(self, _=None):
        if self._empty:
            self.var.set("")
            self._empty = False
        self.entry.configure(fg=INK_2)
        self.rule.configure(bg=ACCENT)

    def _blur(self, _=None):
        self.rule.configure(bg=HAIR)
        if not self.var.get().strip():
            self._empty = True
            self.entry.configure(fg=MUTED)
            self.var.set(self.placeholder)

    def get(self):
        return "" if self._empty else self.var.get().strip()

    def set(self, text):
        text = (text or "")[:self.LIMIT]
        if text.strip():
            self._empty = False
            self.entry.configure(fg=INK_2)
            self.var.set(text)
        else:
            self._empty = True
            self.entry.configure(fg=MUTED)
            self.var.set(self.placeholder)


# ============================================================ roster row ====
class RosterRow(ttk.Frame):
    """One line of the workbook's staffing table, C45:G59.

    Title, start month and salary are the yellow cells and are editable here.
    Type and the benefit load are fixed by the sheet, and the loaded cost is
    derived, so all three are shown but not editable.
    """

    def __init__(self, parent, kind, title, start, salary, load, on_change):
        super().__init__(parent, style="Sub.TFrame")
        self.kind, self.load = kind, load
        self.on_change = on_change

        top = ttk.Frame(self, style="Sub.TFrame")
        top.pack(fill="x")
        self.f_title = PillEntry(top, 168, self._edited, bg=SUBCARD,
                                 field="#e9eef4", justify="left", weight="")
        self.f_title.set_text(title)
        self.f_title.pack(side="left")
        ttk.Label(top, text="start", style="TickSub.TLabel").pack(
            side="left", padx=(px(9), px(4)))
        self.f_start = PillEntry(top, 48, self._edited, bg=SUBCARD,
                                 field="#e9eef4")
        self.f_start.set_text(f"{start:g}")
        self.f_start.pack(side="left")

        # The salary line is labelled with the position's type, straight from
        # column C of the workbook's staffing table.
        self.s_salary = Slider(self, kind, 0, 300_000, 500, salary,
                               "money", self._edited, style="Sub.TFrame",
                               bg=SUBCARD, field="#e9eef4")
        self.s_salary.pack(fill="x", pady=(px(5), 0))
        self.note = ttk.Label(self, text="", style="TickSub.TLabel")
        self.note.pack(anchor="w", pady=(px(1), 0))
        self._refresh()

    def _edited(self, _=None):
        self._refresh()
        self.on_change()

    def _refresh(self):
        loaded = self.salary() * (1 + self.load / 100.0)
        # The type is the slider's label now, so it is not repeated here.
        self.note.config(text=f"{self.load:g}% benefits  \u00b7  "
                              f"${money(loaded)} loaded")

    # -- values ----------------------------------------------------------
    def title(self):
        return self.f_title.text().strip()

    def start(self):
        try:
            return float(self.f_start.text().strip())
        except ValueError:
            return 0.0

    def salary(self):
        return self.s_salary.get()

    def snapshot(self):
        return (self.title(), self.start(), self.salary())

    def restore(self, snap):
        title, start, salary = snap
        self.f_title.set_text(title)
        self.f_start.set_text(f"{start:g}")
        self.s_salary.set(salary)
        self._refresh()


# ================================================================ slider ====
class Slider(ttk.Frame):
    """Label, value field and a custom-drawn track. Values are display units.

    With track=False only the label and field are drawn: the value is typed,
    not dragged, so it can take a fine step no track could land on."""

    def __init__(self, parent, label, lo, hi, step, initial, kind, on_change,
                 sub=None, style="Card.TFrame", bg=CARD, field="#f1f4f8",
                 track=True):
        super().__init__(parent, style=style)
        self.lo, self.hi, self.step, self.kind = lo, hi, step, kind
        self.on_change = on_change
        self.value = initial
        self._drag = False
        self._hot = False

        head = ttk.Frame(self, style=style)
        head.pack(fill="x")
        ttk.Label(head, text=label,
                  style="Param.TLabel" if bg == CARD else "ParamSub.TLabel"
                  ).pack(side="left")
        self.field = PillEntry(head, 76 if track else 110, self._from_entry,
                               bg=bg, field=field)
        self.field.pack(side="right")
        self.cv = None
        if not track:
            self._sync_field()
            return

        self.kd = px(16)
        self.th = px(5)
        self.ch = px(26)
        self.inset = self.kd // 2 + px(1)
        self._r = self.th // 2 + 1

        self.cv = tk.Canvas(self, height=self.ch, bg=bg, highlightthickness=0,
                            bd=0, cursor="hand2")
        self.cv.pack(fill="x", pady=(px(7), 0))

        y = self.ch // 2
        top, bot = y - self.th // 2, y + self.th - self.th // 2
        self._bg_body = self.cv.create_rectangle(0, top, 0, bot, fill=TRACK,
                                                 outline="")
        self._bg_l = self.cv.create_image(0, y, anchor="w", image=self._cap(
            TRACK, "l"))
        self._bg_r = self.cv.create_image(0, y, anchor="w", image=self._cap(
            TRACK, "r"))
        self._fl_body = self.cv.create_rectangle(0, top, 0, bot, fill=ACCENT,
                                                 outline="")
        self._fl_l = self.cv.create_image(0, y, anchor="w", image=self._cap(
            ACCENT, "l"))
        self._knob = self.cv.create_image(0, y, image=self._knob_img(False))

        foot = ttk.Frame(self, style=style)
        foot.pack(fill="x", pady=(px(3), 0))
        tick = "Tick.TLabel" if bg == CARD else "TickSub.TLabel"
        ttk.Label(foot, text=self._edge(lo), style=tick).pack(side="left")
        if sub:
            ttk.Label(foot, text=sub, style=tick).pack(side="left", padx=px(7))
        ttk.Label(foot, text=self._edge(hi), style=tick).pack(side="right")

        self.cv.bind("<Configure>", lambda e: self._redraw())
        self.cv.bind("<Button-1>", self._press)
        self.cv.bind("<B1-Motion>", self._move)
        self.cv.bind("<ButtonRelease-1>", self._release)
        self.cv.bind("<Enter>", lambda e: self._hover(True))
        self.cv.bind("<Leave>", lambda e: self._hover(False))
        self._sync_field()

    # -- sprites ---------------------------------------------------------
    def _cap(self, colour, side):
        return sprite(("cap", self._r, self.th, colour, side),
                      lambda: cap_image(self._r, self.th, colour, side))

    def _knob_img(self, active):
        return sprite(("knob", self.kd, active),
                      lambda: knob_image(self.kd,
                                         _rgba(ACCENT, 38) if active else None))

    # -- geometry --------------------------------------------------------
    def _redraw(self):
        if self.cv is None:
            return
        w = self.cv.winfo_width()
        if w <= 1:
            return
        y = self.ch // 2
        top, bot = y - self.th // 2, y + self.th - self.th // 2
        x0, x1 = self.inset, w - self.inset
        frac = (self.value - self.lo) / (self.hi - self.lo)
        kx = x0 + frac * (x1 - x0)

        self.cv.coords(self._bg_body, x0, top, x1, bot)
        self.cv.coords(self._bg_l, x0 - self._r, y)
        self.cv.coords(self._bg_r, x1, y)
        self.cv.coords(self._fl_body, x0, top, max(x0, kx), bot)
        self.cv.coords(self._fl_l, x0 - self._r, y)
        self.cv.itemconfig(self._fl_l, state="normal" if kx > x0 else "hidden")
        self.cv.coords(self._knob, kx, y)

    def _hover(self, on):
        self._hot = on
        if not self._drag:
            self.cv.itemconfig(self._knob, image=self._knob_img(on))

    # -- interaction -----------------------------------------------------
    def _from_x(self, x):
        w = self.cv.winfo_width()
        x0, x1 = self.inset, w - self.inset
        f = min(max((x - x0) / max(1, x1 - x0), 0.0), 1.0)
        self._set(self.lo + f * (self.hi - self.lo))

    def _press(self, e):
        self._drag = True
        self.cv.itemconfig(self._knob, image=self._knob_img(True))
        self._from_x(e.x)

    def _move(self, e):
        if self._drag:
            self._from_x(e.x)

    def _release(self, _):
        self._drag = False
        self.cv.itemconfig(self._knob, image=self._knob_img(self._hot))

    def _from_entry(self, _=None):
        raw = self.field.text().replace(",", "").replace("$", "").replace("%", "")
        try:
            self._set(float(raw.strip()))
        except ValueError:
            self._sync_field()

    # -- value -----------------------------------------------------------
    def _snap(self, v):
        v = round(v / self.step) * self.step
        return round(min(max(v, self.lo), self.hi), 6)

    def _set(self, v, notify=True, snap=True):
        # Only snapping rounds. Values restored from a scenario or the
        # workbook keep full precision - the sales mix carries far more
        # than six decimals, and rounding it visibly moved the ROI.
        v = self._snap(v) if snap else min(max(v, self.lo), self.hi)
        changed = v != self.value
        self.value = v
        self._sync_field()
        self._redraw()
        if changed and notify:
            self.on_change()

    def _edge(self, v):
        return f"{v:,.0f}" if self.kind == "money" else f"{v:g}"

    def _sync_field(self):
        v = self.value
        self.field.set_text(f"{v:,.0f}" if self.kind == "money" else f"{v:g}")

    def get(self):
        return self.value

    def set(self, v, snap=False):
        """Set without notifying. Values straight from the workbook are kept
        exactly even when they are off the step grid - commission is 48.6%,
        which no 0.25 grid can land on - so the app opens on the real model
        numbers. Dragging and typing still snap."""
        self._set(v, notify=False, snap=snap)


class SliderGroup(ttk.Frame):
    """One master slider that drives N per-year sliders, revealed on demand."""

    def __init__(self, parent, label, lo, hi, step, default, on_change,
                 first_year=1, count=MAX_YEARS, sub=None, kind="pct",
                 per_year_only=False):
        super().__init__(parent, style="Card.TFrame")
        self.on_change = on_change
        self._pushing = False
        self.count = count
        # A per-year-only series has no lead slider to summarise it: the
        # year rows are always shown and there is nothing to toggle.
        self.per_year_only = per_year_only
        # Payroll growth starts in year 2, so it shows one row fewer than
        # the term; everything else starts in year 1.
        self.offset = first_year - 1
        self.visible = count
        values = as_years(default, count)

        # Deliberately not named `master`: that attribute is Tkinter's
        # parent pointer. Overwriting it makes this group and its lead slider
        # each other's parent, and every _root() walk from a widget inside the
        # group then spins forever - which hangs the whole event loop.
        # Kept even when hidden: snapshots, restore and set_values all
        # still round-trip a lead value through it.
        self.lead = Slider(self, label, lo, hi, step, lead_value(values),
                           kind, self._lead_moved, sub=sub)
        if not per_year_only:
            self.lead.pack(fill="x")

        self.expanded = tk.BooleanVar(value=per_year_only)
        self.toggle = ttk.Checkbutton(
            self, variable=self.expanded, command=self._toggle,
            style="Link.TCheckbutton",
            text=f" set each year individually  ({count} years)")
        if not per_year_only:
            self.toggle.pack(anchor="w", pady=(px(7), 0))

        self.box = ttk.Frame(self, style="Sub.TFrame")
        if per_year_only:
            # With the lead slider gone the rows need a heading of their own.
            ttk.Label(self.box, text=caps(label),
                      style="SubSection.TLabel").pack(
                          anchor="w", padx=px(13), pady=(px(11), px(4)))
        self.rows = []
        for i in range(count):
            s = Slider(self.box, f"Year {first_year + i}", lo, hi, step,
                       values[i], kind, self._child_moved, style="Sub.TFrame",
                       bg=SUBCARD, field="#e9eef4")
            s.pack(fill="x", padx=px(13), pady=px(7))
            self.rows.append(s)

        # A series the workbook varies year by year opens already expanded,
        # otherwise the single lead slider would misrepresent it.
        if per_year_only:
            self.box.pack(fill="x", pady=(0, px(2)))
        elif not is_uniform(values):
            self.expanded.set(True)
            self.box.pack(fill="x", pady=(px(8), px(2)))

    def set_term(self, term):
        """Show only the years that fall inside the contract term."""
        want = max(0, min(self.count, term - self.offset))
        if want == self.visible:
            return
        self.visible = want
        for s in self.rows:               # repack in order so years stay sorted
            s.pack_forget()
        for s in self.rows[:want]:
            s.pack(fill="x", padx=px(13), pady=px(7))
        self.toggle.configure(
            text=f" set each year individually  ({want} years)")

    def _toggle(self):
        if self.expanded.get():
            self.box.pack(fill="x", pady=(px(8), px(2)))
        else:
            self.box.forget()
        self.on_change()

    def _lead_moved(self):
        self._pushing = True
        for s in self.rows:
            s.set(self.lead.get())
        self._pushing = False
        self.on_change()

    def _child_moved(self):
        if not self._pushing:
            self.on_change()

    def values(self):
        n = max(1, self.visible)
        if self.expanded.get():
            return [s.get() for s in self.rows[:n]]
        return [self.lead.get()] * n

    def snapshot(self):
        """Lead value, every per-year value, and whether the panel is open."""
        return (self.lead.get(), [s.get() for s in self.rows],
                bool(self.expanded.get()))

    def restore(self, snap):
        lead, rows, expanded = snap
        self.lead.set(lead)
        for s, v in zip(self.rows, rows):
            s.set(v)
        if self.per_year_only:
            return
        if bool(self.expanded.get()) != expanded:
            self.expanded.set(expanded)
            if expanded:
                self.box.pack(fill="x", pady=(px(8), px(2)))
            else:
                self.box.forget()

    def set_values(self, default):
        values = as_years(default, self.count)
        self.lead.set(lead_value(values))
        for s, v in zip(self.rows, values):
            s.set(v)
        if self.per_year_only:
            return
        want = not is_uniform(values)
        if want != bool(self.expanded.get()):
            self.expanded.set(want)
            if want:
                self.box.pack(fill="x", pady=(px(8), px(2)))
            else:
                self.box.forget()


# ==================================================================== app ====
class App(tk.Tk):
    def __init__(self):
        global K
        super().__init__()
        self.withdraw()

        if sys.platform == "darwin":
            # Tk on macOS reports 72 dpi and handles Retina scaling itself;
            # deriving K from that would shrink the whole UI by a quarter.
            K = 1.0
        else:
            dpi = self.winfo_fpixels("1i")
            K = dpi / 96.0
            self.tk.call("tk", "scaling", dpi / 72.0)  # point fonts -> true size
        resolve_fonts(self)

        self.title(APP_TITLE)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(px(1320), sw - px(40))
        h = min(px(840), sh - px(90))
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 2 - px(20))}")
        # The header row - long title, note field and eight buttons -
        # needs about 1150 logical px before anything starts clipping.
        self.minsize(min(px(1150), sw - px(20)), min(px(680), sh - px(60)))
        self.configure(bg=PAGE)

        self._job = None
        self._fitting = False
        self.active = 0
        self.scenarios = [None] * SCENARIOS
        self._styles()
        self._build()
        self._apply_term()      # year sliders and columns follow the term

        # Every scenario starts from the workbook's opening numbers.
        self.scenarios = [self.snapshot() for _ in range(SCENARIOS)]
        opening = self._roi_text(compute(self.params_from(self.scenarios[0])))
        for snap in self.scenarios:
            snap["roi"] = opening
        self._paint_scenarios()

        self.recalc()
        self.deiconify()

    # ------------------------------------------------------------ styling --
    def _styles(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TFrame", background=PAGE)
        s.configure("Card.TFrame", background=CARD)
        s.configure("Sub.TFrame", background=SUBCARD)
        s.configure("TLabel", background=PAGE, foreground=INK, font=F(10))
        s.configure("Param.TLabel", background=CARD, foreground=INK_2, font=F(10))
        s.configure("ParamSub.TLabel", background=SUBCARD, foreground=INK_2,
                    font=F(10))
        s.configure("Tick.TLabel", background=CARD, foreground=MUTED, font=F(8))
        s.configure("TickSub.TLabel", background=SUBCARD, foreground=MUTED,
                    font=F(8))
        s.configure("Section.TLabel", background=CARD, foreground=MUTED,
                    font=F(8, "sb"))
        s.configure("SubSection.TLabel", background=SUBCARD,
                    foreground=MUTED, font=F(8, "sb"))
        s.configure("Value.TLabel", background=CARD, foreground=INK,
                    font=F(11, "sb"))
        s.configure("Warn.TLabel", background=CARD, foreground=NEG,
                    font=F(8))
        s.configure("H1.TLabel", background=PAGE, foreground=INK, font=F(18, "sl"))
        s.configure("Sub.TLabel", background=PAGE, foreground=MUTED_2, font=F(10))
        s.configure("Link.TCheckbutton", background=CARD, foreground=ACCENT,
                    font=F(8), focuscolor=CARD, indicatorsize=px(9))
        s.map("Link.TCheckbutton", background=[("active", CARD)],
              foreground=[("active", ACCENT)])

        s.configure("Treeview", background=CARD, fieldbackground=CARD,
                    foreground=INK_2, rowheight=px(25), font=F(10),
                    borderwidth=0, relief="flat")
        s.configure("Treeview.Heading", background=CARD, foreground=MUTED,
                    font=F(8, "sb"), relief="flat", borderwidth=0, padding=px(4))
        s.map("Treeview.Heading", background=[("active", CARD)])
        s.map("Treeview", background=[("selected", CARD)],
              foreground=[("selected", INK_2)])

        # A slim, arrow-less scrollbar.
        s.layout("Slim.Vertical.TScrollbar",
                 [("Vertical.Scrollbar.trough",
                   {"sticky": "ns", "children":
                    [("Vertical.Scrollbar.thumb",
                      {"expand": "1", "sticky": "nswe"})]})])
        s.configure("Slim.Vertical.TScrollbar", troughcolor=CARD,
                    background="#dbe2ea", borderwidth=0, relief="flat",
                    width=px(7), arrowsize=0)
        s.map("Slim.Vertical.TScrollbar", background=[("active", "#c3ccd8")])

    # ------------------------------------------------------------- layout --
    def _build(self):
        head = ttk.Frame(self, padding=(px(22), px(16), px(22), px(8)))
        head.pack(fill="x")
        # The buttons are packed BEFORE the title so they claim their space
        # first. Packed the other way round the title took its full width and
        # the leftmost buttons - Reset and Save Across - were squeezed to
        # nothing on a narrower window. The title absorbs the slack instead.
        # Right to left, this reads
        # Reset | Save Across | 1 | 2 | 3 | Save Setup | Load Setup | Export.
        PillButton(head, "Export XLSX", self.export, primary=True).pack(
            side="right", padx=(px(10), 0), pady=(px(7), 0))
        PillButton(head, "Load Setup", self.load_setup).pack(
            side="right", padx=(px(6), 0), pady=(px(7), 0))
        PillButton(head, "Save Setup", self.save_setup).pack(
            side="right", padx=(px(6), 0), pady=(px(7), 0))
        self.scen_btns = []
        for i in reversed(range(SCENARIOS)):
            b = ScenarioButton(head, f"Scenario {i + 1}",
                               lambda n=i: self._select_scenario(n))
            b.pack(side="right", padx=(px(6), 0), pady=(px(2), 0))
            self.scen_btns.insert(0, b)
        PillButton(head, "Save Across", self.save_across).pack(
            side="right", padx=(px(6), 0), pady=(px(7), 0))
        PillButton(head, "Reset", self.reset).pack(
            side="right", padx=(px(6), 0), pady=(px(7), 0))

        title = ttk.Frame(head)
        title.pack(side="left", fill="x", expand=True)
        row = ttk.Frame(title)
        row.pack(anchor="w", fill="x")
        ttk.Label(row, text="Hotel ROI Modeling Tool",
                  style="H1.TLabel").pack(side="left")
        self.note = NoteEntry(row, chars=26)
        self.note.pack(side="left", padx=(px(14), 0), pady=(px(9), 0),
                       fill="x", expand=True)
        ttk.Label(title, text="Contract cash-flow model  ·  App Model.xlsx",
                  style="Sub.TLabel").pack(anchor="w", pady=(px(2), 0))

        body = ttk.Frame(self, padding=(px(15), 0, px(15), px(8)))
        body.pack(fill="both", expand=True)

        self._build_params(body)
        self.right = ttk.Frame(body)
        self.right.pack(side="left", fill="both", expand=True, padx=(px(5), 0))
        self._build_results(self.right)
        self.right.bind("<Configure>", self._fit_chart)

    # ------------------------------------------------- parameters (left) ---
    _sizing_cols = False

    def _build_params(self, parent):
        card = Card(parent, pad=0)
        card.configure(width=px(432))
        card.pack(side="left", fill="y")
        card.pack_propagate(False)

        canvas = tk.Canvas(card.body, bg=CARD, highlightthickness=0, bd=0)
        bar = ttk.Scrollbar(card.body, orient="vertical", command=canvas.yview,
                            style="Slim.Vertical.TScrollbar")
        inner = ttk.Frame(canvas, style="Card.TFrame")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        canvas.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y", padx=(0, px(5)), pady=px(12))
        canvas.pack(side="left", fill="both", expand=True)
        self.bind_all("<MouseWheel>", self._maybe_scroll)
        # X11 reports the wheel as buttons 4 and 5 rather than <MouseWheel>.
        self.bind_all("<Button-4>", self._maybe_scroll)
        self.bind_all("<Button-5>", self._maybe_scroll)
        self._param_canvas = canvas

        r = self.recalc_soon
        d = DEFAULTS

        g = self._group(inner, "Contract Term (Years)", first=True)
        self.s_term = Slider(g, "Contract Term", 1, MAX_YEARS, TERM_STEP,
                             d["term"], "pct", self._term_changed,
                             sub="operating years, quarter-year steps")
        self.s_term.pack(fill="x")

        g = self._group(inner, "Revenue")
        self.s_revenue = Slider(g, "Gross Revenue (Year 0)", 0, 1_000_000_000,
                                1, d["gross_revenue"], "money", r, track=False)
        self.s_revenue.pack(fill="x", pady=(0, px(16)))
        self.g_growth = SliderGroup(g, "Revenue Growth", -10, 30, 0.25,
                                    d["rev_growth"], r, per_year_only=True)
        self.g_growth.pack(fill="x")

        # The share of sales each category represents - one half of the
        # commission blend below, so it is set before the rates it weights.
        g = self._group(inner, "Sales By Category")
        self.mix_lbl = ttk.Label(g, text="", style="Tick.TLabel")
        self.mix_lbl.pack(anchor="w", pady=(0, px(8)))
        self.mix_box = ttk.Frame(g, style="Sub.TFrame")
        self.mix_box.pack(fill="x", pady=(0, px(2)), ipady=px(5))
        self.s_mix = []
        for i, cat in enumerate(CATEGORIES):
            sl = Slider(self.mix_box, cat, 0, 100, 0.25, d["sales_mix"][i],
                        "pct", r, style="Sub.TFrame", bg=SUBCARD,
                        field="#e9eef4")
            sl.pack(fill="x", padx=px(13), pady=px(6))
            self.s_mix.append(sl)

        # Commission is blended from a rate per income category against the
        # sales mix above - the workbook's SUMPRODUCT.
        g = self._group(inner, "Commission")
        head = ttk.Frame(g, style="Card.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text="Blended Commission",
                  style="Param.TLabel").pack(side="left")
        self.comm_lbl = ttk.Label(head, text="-", style="Value.TLabel")
        self.comm_lbl.pack(side="right")

        self.comm_box = ttk.Frame(g, style="Sub.TFrame")
        self.comm_box.pack(fill="x", pady=(px(8), px(2)))
        ttk.Label(self.comm_box, text=caps("Contract Commission Rates"),
                  style="SubSection.TLabel").pack(anchor="w", padx=px(13),
                                                  pady=(px(11), px(4)))
        self.s_rates = []
        for i, cat in enumerate(CATEGORIES):
            sl = Slider(self.comm_box, cat, 0, 100, 0.25,
                        d["commission_rates"][i], "pct", r,
                        style="Sub.TFrame", bg=SUBCARD, field="#e9eef4")
            sl.pack(fill="x", padx=px(13), pady=px(6))
            self.s_rates.append(sl)

        g = self._group(inner, "Cost Structure")
        self.g_cogs = SliderGroup(g, "COGS", 0, 60, 0.25, d["cogs"], r,
                                  sub="% of net revenue")
        self.g_cogs.pack(fill="x", pady=(0, px(16)))
        self.g_expenses = SliderGroup(g, "Expenses", 0, 60, 0.25,
                                      d["expenses"], r,
                                      sub="% of net revenue")
        self.g_expenses.pack(fill="x")

        # Payroll is built from the staffing roster on the workbook's
        # "Payroll Calculation" tab, not entered as a share of revenue.
        g = self._group(inner, "Payroll")
        head = ttk.Frame(g, style="Card.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text="Payroll (Year 1)",
                  style="Param.TLabel").pack(side="left")
        self.pay_lbl = ttk.Label(head, text="-", style="Value.TLabel")
        self.pay_lbl.pack(side="right")
        self.pay_sub = ttk.Label(g, text="", style="Tick.TLabel")
        self.pay_sub.pack(anchor="w", pady=(px(3), 0))

        self.roster_open = tk.BooleanVar(value=False)
        self.roster_link = ttk.Checkbutton(
            g, variable=self.roster_open, command=self._toggle_roster,
            style="Link.TCheckbutton",
            text=f" set the staff roster ({len(STAFF_TITLES)} roles)"
        )
        self.roster_link.pack(anchor="w", pady=(px(7), 0))

        self.roster_box = ttk.Frame(g, style="Sub.TFrame")
        self.roster = []
        for i in range(len(STAFF_TITLES)):
            if i:
                tk.Frame(self.roster_box, bg="#e4e9f0", height=1).pack(
                    fill="x", padx=px(13), pady=px(2))
            row = RosterRow(self.roster_box, STAFF_TYPES[i],
                            d["staff_titles"][i], d["staff_starts"][i],
                            d["staff_salaries"][i], d["staff_loads"][i], r)
            row.pack(fill="x", padx=px(13), pady=(px(9), px(7)))
            self.roster.append(row)

        self.s_merit = Slider(g, "Merit Increase", 0, 15, 0.25, d["merit"],
                              "pct", r, sub="a year, across the roster")
        self.s_merit.pack(fill="x", pady=(px(16), px(16)))
        self.s_incentive = Slider(g, "Incentives & Bonus", 0, 10, 0.25,
                                  d["incentive_pct"], "pct", r,
                                  sub="% of gross revenue")
        self.s_incentive.pack(fill="x", pady=(0, px(16)))
        self.s_profit = Slider(g, "Profit Sharing", 0, 10, 0.25,
                               d["profit_share_pct"], "pct", r,
                               sub="% of net revenue")
        self.s_profit.pack(fill="x")

        # Year 0 is no longer just capex: the workbook adds a month of
        # payroll, a start-up expense and a signing bonus to make the
        # Initial Investment that the ROI is measured against.
        g = self._group(inner, "Capital")
        head = ttk.Frame(g, style="Card.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text="Initial Investment",
                  style="Param.TLabel").pack(side="left")
        self.invest_lbl = ttk.Label(head, text="-", style="Value.TLabel")
        self.invest_lbl.pack(side="right")
        self.invest_sub = ttk.Label(g, text="", style="Tick.TLabel")
        self.invest_sub.pack(anchor="w", pady=(px(3), px(14)))

        self.s_capex0 = Slider(g, "Initial Capex (Year 0)", 0, 150, 0.25,
                               d["capex_initial"], "pct", r,
                               sub="% of year-1 gross revenue")
        self.s_capex0.pack(fill="x", pady=(0, px(16)))
        self.s_year0_exp = Slider(g, "Year-0 Expenses", 0, 10, 0.05,
                                  d["year0_expense_pct"], "pct", r,
                                  sub="% of year-1 gross revenue")
        self.s_year0_exp.pack(fill="x", pady=(0, px(16)))
        # The workbook keeps two bonus inputs: D76 lands in year 0 as part of
        # the Initial Investment, E79:J79 recur and are netted off each year.
        self.s_bonus = Slider(g, "Signing Bonus (Year 0)", 0, 500_000, 1_000,
                              d["signing_bonus"][0], "money", r,
                              sub="part of the initial investment")
        self.s_bonus.pack(fill="x", pady=(0, px(16)))
        self.g_bonus_yr = SliderGroup(g, "Signing Bonus (Yearly)", 0,
                                      500_000, 1_000, d["signing_bonus"][1:],
                                      r, kind="money",
                                      sub="each operating year")
        self.g_bonus_yr.pack(fill="x", pady=(0, px(16)))
        self.g_capex = SliderGroup(g, "Ongoing Capex", 0, 30, 0.25, d["capex"], r)
        self.g_capex.pack(fill="x")

        ttk.Frame(inner, style="Card.TFrame", height=px(20)).pack(fill="x")

        self.groups = {"growth": self.g_growth, "cogs": self.g_cogs,
                       "expenses": self.g_expenses, "capex": self.g_capex,
                       "bonus_yr": self.g_bonus_yr}

    def _toggle_roster(self):
        if self.roster_open.get():
            self.roster_box.pack(fill="x", pady=(px(8), px(2)),
                                 after=self.roster_link)
        else:
            self.roster_box.forget()
        self.recalc_soon()

    def _apply_term(self):
        """Sync the year sliders and the table columns to the term.

        A fractional term ends in a stub year, which still needs a column of
        its own - it just bills part of a year. The header says how many
        months it carries so a short column is never read as a full one.
        """
        term = self.s_term.get()
        n = year_count(term)
        weights = term_weights(term)
        for g in getattr(self, "groups", {}).values():
            g.set_term(n)
        if getattr(self, "tree", None):
            for i in range(1, MAX_YEARS + 1):
                w = weights[i - 1] if i <= n else 1.0
                months = "" if w >= 1.0 else f" ({w * 12:g} MO)"
                self.tree.heading(f"y{i}", text=f"YEAR {i}{months}")
            self.tree.configure(
                displaycolumns=["item"] + [f"y{i}" for i in range(n + 1)])
            self._fit_columns(n)
        return n

    def _fit_columns(self, n=None):
        """Re-share the table width across the visible year columns.

        Treeview only re-runs its stretch on a resize, so changing the term
        would otherwise leave the columns at the widths the old term had
        stretched them to and push the last year off the right edge.
        """
        tree = getattr(self, "tree", None)
        if tree is None or self._sizing_cols:
            return
        if n is None:
            n = year_count(self.s_term.get())
        total = tree.winfo_width()
        if total <= 1:                    # not mapped yet; the bind will retry
            return
        self._sizing_cols = True
        try:
            item = int(tree.column("item", "width"))
            each = max(px(74), (total - item - px(4)) // (n + 1))
            for i in range(MAX_YEARS + 1):
                tree.column(f"y{i}", width=each)
        finally:
            self._sizing_cols = False

    def _term_changed(self):
        self._apply_term()
        self.recalc_soon()

    def _group(self, parent, title, first=False):
        if not first:
            tk.Frame(parent, bg=HAIR, height=1).pack(
                fill="x", padx=px(24), pady=(px(21), 0))
        wrap = ttk.Frame(parent, style="Card.TFrame")
        wrap.pack(fill="x", padx=px(24), pady=(px(19), 0))
        ttk.Label(wrap, text=caps(title), style="Section.TLabel").pack(
            anchor="w", pady=(0, px(14)))
        return wrap

    def _maybe_scroll(self, e):
        """Scroll the parameter column only when the pointer is over it."""
        c = self._param_canvas
        x, y = c.winfo_rootx(), c.winfo_rooty()
        if x <= e.x_root <= x + c.winfo_width() and \
           y <= e.y_root <= y + c.winfo_height():
            c.yview_scroll(_wheel_steps(e), "units")

    # ----------------------------------------------------- results (right) --
    def _build_results(self, parent):
        # ---- hero -------------------------------------------------------
        hero = self.hero_card = Card(parent, bg=HERO, pad=21)
        hero.pack(fill="x")

        left = tk.Frame(hero.body, bg=HERO)
        left.pack(side="left")
        tk.Label(left, text=caps("Return on Investment"), bg=HERO,
                 fg=ACCENT_HI, font=F(8, "sb")).pack(anchor="w")
        self.roi_lbl = tk.Label(left, text="-", bg=HERO, fg=ACCENT_HI,
                                font=F(44, "sl"))
        self.roi_lbl.pack(anchor="w", pady=(px(6), 0))
        self.hero_note = tk.Label(
            left, text="annualized on the year-0 capex", bg=HERO,
            fg="#7c8aa0", font=F(9))
        self.hero_note.pack(anchor="w")

        stats = tk.Frame(hero.body, bg=HERO)
        stats.pack(side="right", pady=(px(12), 0))
        self.stat = {}
        for i, (key, cap) in enumerate((("invest", "Initial Investment"),
                                        ("total", "Total Net Cash"),
                                        ("payback", "Payback"),
                                        ("margin", "Avg NOI Margin"))):
            if i:
                tk.Frame(stats, bg=HERO_2, width=1).pack(
                    side="left", fill="y", padx=px(18), pady=px(4))
            col = tk.Frame(stats, bg=HERO)
            col.pack(side="left")
            tk.Label(col, text=caps(cap), bg=HERO, fg="#69788e",
                     font=F(7, "sb")).pack(anchor="w")
            v = tk.Label(col, text="-", bg=HERO, fg="#e9eef5", font=F(16, "sl"))
            v.pack(anchor="w", pady=(px(4), 0))
            self.stat[key] = v

        # ---- table ------------------------------------------------------
        table = self.table_card = Card(parent, pad=12)
        table.pack(fill="both", expand=True, pady=(px(3), 0))

        # Columns exist for every possible year; the term picks which show.
        cols = ["item"] + [f"y{i}" for i in range(MAX_YEARS + 1)]
        self.tree = ttk.Treeview(table.body, columns=cols, show="headings",
                                 height=12, selectmode="none")
        self.tree.heading("item", text="")
        self.tree.column("item", width=px(158), anchor="w", stretch=False)
        for i in range(MAX_YEARS + 1):
            self.tree.heading(f"y{i}", text=f"YEAR {i}")
            self.tree.column(f"y{i}", width=px(92), minwidth=px(74),
                             anchor="e", stretch=True)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Configure>", lambda e: self._fit_columns())
        self.tree.tag_configure("head", font=F(10, "sb"), foreground=INK)
        self.tree.tag_configure("soft", foreground=MUTED)
        self.tree.tag_configure("cash", font=F(10, "sb"), foreground=INK,
                                background=ACCENT_BG)

        # Row order and formats come from roi_core; only the visual
        # emphasis is a front-end concern.
        tags = {"_growth": "soft", "net_rev": "head", "noi": "head",
                "margin": "soft", "payroll_pct": "soft",
                "net_cash": "cash", "cumulative": "soft"}
        self.rows = [(lbl, key, fmt, tags.get(key, "")) for lbl, key, fmt in ROWS]
        self.year0 = YEAR0
        self.iids = [self.tree.insert("", "end",
                                      values=[r[0]] + [""] * MAX_YEARS,
                                      tags=(r[3],) if r[3] else ())
                     for r in self.rows]
        self.tree.configure(
            displaycolumns=["item"]
            + [f"y{i}" for i in range(DEFAULT_TERM + 1)])

        # ---- chart ------------------------------------------------------
        chart = self.chart_card = Card(parent, pad=11)
        chart.pack(fill="x", pady=(px(3), px(3)))
        tk.Label(chart.body, text=caps("Net cash by year"), bg=CARD, fg=MUTED,
                 font=F(8, "sb")).pack(anchor="w")
        self.chart = tk.Canvas(chart.body, height=px(128), bg=CARD,
                               highlightthickness=0, bd=0)
        self.chart.pack(fill="x", pady=(px(7), 0))
        self.chart.bind("<Configure>", lambda e: self.draw_chart())

    def _fit_chart(self, _=None):
        """Give the chart whatever height is left; drop it if that is too little.

        The hero and the twelve-row table have fixed heights, so on a short
        window the chart is what has to give. Without this it simply ran off
        the bottom edge, taking the Year-0 bar and the axis labels with it.
        """
        if self._fitting or not hasattr(self, "chart_card"):
            return
        total = self.right.winfo_height()
        if total <= 1:
            return
        self._fitting = True
        try:
            spare = (total - self.hero_card.winfo_reqheight()
                     - self.table_card.winfo_reqheight() - px(6))
            chrome = self.chart_card.winfo_reqheight() - int(self.chart["height"])
            want = min(max(spare - chrome, 0), px(150))
            shown = bool(self.chart_card.winfo_manager())
            if want < px(70):
                if shown:
                    self.chart_card.pack_forget()
            else:
                if not shown:
                    self.chart_card.pack(fill="x", pady=(px(3), px(3)))
                if abs(want - int(self.chart["height"])) > 2:
                    self.chart.configure(height=want)
        finally:
            self._fitting = False

    # ------------------------------------------------------------- engine --
    def snapshot(self):
        """Everything the controls currently hold, enough to restore exactly."""
        return {
            "term": float(self.s_term.get()),
            "rev": self.s_revenue.get(),
            "rates": [s.get() for s in self.s_rates],
            "mix": [s.get() for s in self.s_mix],
            "merit": self.s_merit.get(),
            "incentive": self.s_incentive.get(),
            "profit": self.s_profit.get(),
            "roster": [row.snapshot() for row in self.roster],
            "roster_open": bool(self.roster_open.get()),
            "capex0": self.s_capex0.get(),
            "year0_exp": self.s_year0_exp.get(),
            "bonus": self.s_bonus.get(),
            "groups": {name: g.snapshot() for name, g in self.groups.items()},
        }

    def restore(self, snap):
        self.s_term.set(snap["term"])
        self.s_revenue.set(snap["rev"])
        self.s_merit.set(snap["merit"])
        self.s_incentive.set(snap["incentive"])
        self.s_profit.set(snap["profit"])
        self.s_capex0.set(snap["capex0"])
        self.s_year0_exp.set(snap["year0_exp"])
        self.s_bonus.set(snap["bonus"])
        for row, v in zip(self.roster, snap["roster"]):
            row.restore(v)
        if bool(self.roster_open.get()) != snap["roster_open"]:
            self.roster_open.set(snap["roster_open"])
            if snap["roster_open"]:
                self.roster_box.pack(fill="x", pady=(px(8), px(2)),
                                     after=self.roster_link)
            else:
                self.roster_box.forget()
        for sl, v in zip(self.s_rates, snap["rates"]):
            sl.set(v)
        for sl, v in zip(self.s_mix, snap["mix"]):
            sl.set(v)
        for name, g in self.groups.items():
            g.restore(snap["groups"][name])
        self._apply_term()
        self.recalc()

    def params_from(self, snap):
        """Model inputs for a snapshot. A collapsed group contributes its lead
        value for every year, exactly as the live controls do, and only the
        years inside the term are passed on."""
        term = float(snap["term"])

        def series(name):
            lead, rows, expanded = snap["groups"][name]
            # A stub year still needs its own rate, so the series runs to the
            # year count, not to the (possibly fractional) term.
            n = year_count(term)
            return list(rows[:n]) if expanded else [lead] * n

        return {
            "term": term,
            "gross_revenue": snap["rev"],
            "rev_growth": series("growth"),
            "commission_rates": list(snap["rates"]),
            "sales_mix": list(snap["mix"]),
            "cogs": series("cogs"),
            "expenses": series("expenses"),
            "staff_titles": [t for t, _, _ in snap["roster"]],
            "staff_starts": [m for _, m, _ in snap["roster"]],
            "staff_salaries": [s for _, _, s in snap["roster"]],
            "staff_loads": [row.load for row in self.roster],
            "merit": snap["merit"],
            "incentive_pct": snap["incentive"],
            "profit_share_pct": snap["profit"],
            "capex_initial": snap["capex0"],
            "capex": series("capex"),
            "year0_expense_pct": snap["year0_exp"],
            "signing_bonus": [snap["bonus"]] + series("bonus_yr"),
        }

    def params(self):
        return self.params_from(self.snapshot())

    # ---------------------------------------------------------- scenarios --
    @staticmethod
    def _roi_text(r):
        return pct(r["roi"], 2) if r["roi"] is not None else "n/a"

    def _save_active(self):
        """The live controls always belong to the highlighted scenario."""
        snap = self.snapshot()
        snap["roi"] = self._roi_text(compute(self.params_from(snap)))
        self.scenarios[self.active] = snap

    def _paint_scenarios(self):
        for i, b in enumerate(self.scen_btns):
            b.set_active(i == self.active)
            if i != self.active and self.scenarios[i]:
                b.set_roi(self.scenarios[i].get("roi", "-"))

    def save_across(self):
        """Copy what is on screen into all three scenarios.

        Each slot gets its own snapshot rather than a shared reference, so the
        scenarios stay independent the moment you start editing them again.
        """
        text = self._roi_text(compute(self.params_from(self.snapshot())))
        self.scenarios = []
        for _ in range(SCENARIOS):
            snap = self.snapshot()
            snap["roi"] = text
            self.scenarios.append(snap)
        self._paint_scenarios()

    def _select_scenario(self, index):
        """Store what is on screen, then show the scenario that was clicked."""
        self._save_active()
        if index != self.active:
            self.active = index
            self.restore(self.scenarios[index])
        self._paint_scenarios()

    def recalc_soon(self):
        """Coalesce rapid slider events into a single repaint."""
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(15, self.recalc)

    def recalc(self):
        self._job = None
        p = self.params()
        r = self.res = compute(p)

        self.roi_lbl.config(
            text=pct(r["roi"], 2) if r["roi"] is not None else "n/a",
            fg=ACCENT_HI if (r["roi"] or 0) >= 0 else NEG_2)
        if getattr(self, "scen_btns", None):
            self.scen_btns[self.active].set_roi(self._roi_text(r))
        self.stat["invest"].config(text="$" + money(-r["net_cash"][0]))
        self.stat["total"].config(text="$" + money(r["total"]),
                                  fg="#e9eef5" if r["total"] >= 0 else NEG_2)
        self.stat["payback"].config(
            text=f"{r['payback']:.1f} yrs" if r["payback"] is not None else "never")
        term = r["term"]
        # Columns are whole years even when the term is not: a 6.25-year
        # term fills seven of them, the last carrying only its three months.
        years = r["years"]
        self.stat["margin"].config(text=pct(r["avg_margin"]))
        self.hero_note.config(
            text=f"annualized on the year-0 capex, over "
                 f"{term:g} year{'' if term == 1 else 's'}")
        self.comm_lbl.config(text=pct(r["commission_pct"] / 100.0, 2))
        self.invest_lbl.config(text="$" + money(r["invest"][0]))
        self.invest_sub.config(
            text=f"capex ${money(r['capex'][0])}  +  payroll "
                 f"${money(r['year0_payroll'])}  +  expenses "
                 f"${money(r['year0_expense'])}  +  bonus "
                 f"${money(r['year0_bonus'])}")
        self.pay_lbl.config(text="$" + money(r["payroll"][1]))
        share = r["payroll"][1] / r["net_rev"][1] if r["net_rev"][1] else 0.0
        self.pay_sub.config(
            text=f"{pct(share)} of year-1 net revenue  ·  roster "
                 f"${money(r['staff_base'])} before merit")
        total_mix = sum(s.get() for s in self.s_mix)
        self.mix_lbl.config(
            text=f"mix totals {total_mix:g}%",
            style="Tick.TLabel" if abs(total_mix - 100) < 0.005
            else "Warn.TLabel")

        for iid, (label, key, kind, _tag) in zip(self.iids, self.rows):
            series = r[key]
            out = []
            for t in range(MAX_YEARS + 1):
                if (t > years or (t == 0 and key not in self.year0)
                        or (t > 0 and key in ONLY_YEAR0)):
                    out.append("")
                elif kind == "money":
                    out.append(money(series[t]))
                else:
                    out.append(pct(series[t]))
            self.tree.item(iid, values=[label] + out)

        self.draw_chart()

    # -------------------------------------------------------------- chart --
    def draw_chart(self):
        c = self.chart
        c.delete("all")
        if not hasattr(self, "res"):
            return
        vals = self.res["net_cash"]
        w = c.winfo_width() or px(700)
        h = int(c["height"])
        pad_l, pad_r = px(4), px(4)
        pad_t = int(max(px(11), min(px(17), h * 0.13)))
        pad_b = int(max(px(19), min(px(28), h * 0.21)))
        plot_w, plot_h = w - pad_l - pad_r, h - pad_t - pad_b
        if plot_w <= px(20):
            return

        hi = max(max(vals), 0.0)
        lo = min(min(vals), 0.0)
        span = (hi - lo) or 1.0
        zero_y = pad_t + plot_h * (hi / span)
        c.create_line(pad_l, zero_y, w - pad_r, zero_y, fill=HAIR)

        years = len(vals) - 1          # the term drives how many bars
        slot = plot_w / (years + 1)
        bw = int(min(slot * 0.44, px(50)))
        cap_r = max(1, min(px(4), bw // 2))
        for t, v in enumerate(vals):
            cx = pad_l + slot * (t + 0.5)
            y = zero_y - plot_h * (v / span)
            up = v >= 0
            colour = ACCENT_2 if up else NEG_2
            x0 = int(cx - bw / 2)

            if abs(y - zero_y) > cap_r:
                img = sprite(("bar", bw, cap_r, colour, up),
                             lambda bw=bw, r=cap_r, col=colour, u=up:
                             bar_cap_image(bw, r, col, u))
                c.create_image(x0, y if up else y - cap_r, image=img, anchor="nw")
                if up:
                    c.create_rectangle(x0, y + cap_r, x0 + bw, zero_y,
                                       fill=colour, outline="")
                else:
                    c.create_rectangle(x0, zero_y, x0 + bw, y - cap_r,
                                       fill=colour, outline="")
            else:
                c.create_rectangle(x0, min(y, zero_y), x0 + bw, max(y, zero_y),
                                   fill=colour, outline="")

            c.create_text(cx, h - px(10), text=f"Y{t}", fill=MUTED, font=F(8))
            # A negative bar reaches the floor of the plot, so its value
            # goes above the zero line rather than under the bar, where it
            # would sit on top of the axis label.
            c.create_text(cx, y - px(10) if up else zero_y - px(9),
                          text=money(v), fill=INK_2 if up else NEG,
                          font=F(8, "sb"))

    # ------------------------------------------------------------ actions --
    def reset(self):
        d = DEFAULTS
        self.s_revenue.set(d["gross_revenue"])
        self.g_growth.set_values(d["rev_growth"])
        self.s_term.set(d["term"])
        for sl, v in zip(self.s_rates, d["commission_rates"]):
            sl.set(v)
        for sl, v in zip(self.s_mix, d["sales_mix"]):
            sl.set(v)
        self.g_cogs.set_values(d["cogs"])
        self.g_expenses.set_values(d["expenses"])
        self.s_merit.set(d["merit"])
        self.s_incentive.set(d["incentive_pct"])
        self.s_profit.set(d["profit_share_pct"])
        for i, row in enumerate(self.roster):
            row.restore((d["staff_titles"][i], d["staff_starts"][i],
                         d["staff_salaries"][i]))
        self.s_capex0.set(d["capex_initial"])
        self.s_year0_exp.set(d["year0_expense_pct"])
        self.s_bonus.set(d["signing_bonus"][0])
        self.g_bonus_yr.set_values(d["signing_bonus"][1:])
        self.g_capex.set_values(d["capex"])
        self._apply_term()
        self.recalc()

    def save_setup(self):
        """Write every scenario's parameters to a file the app can read back."""
        path = filedialog.asksaveasfilename(
            defaultextension=".json", initialfile="roi_setup.json",
            filetypes=[("Hotel ROI Modeling Tool setup", "*.json")])
        if not path:
            return
        self._save_active()          # the live controls are the active scenario
        data = {
            "format": SETUP_FORMAT,
            "version": SETUP_VERSION,
            "saved": datetime.now().isoformat(timespec="seconds"),
            "note": self.note.get(),
            "active": self.active,
            "scenarios": self.scenarios,
        }
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except OSError as e:
            messagebox.showerror(
                APP_TITLE, "Could not write the setup file:\n" + str(e))
            return
        messagebox.showinfo(
            APP_TITLE,
            "Saved to\n" + path
            + f"\n\nAll {SCENARIOS} scenarios are in this file - every slider, "
              "the staff roster, the note and which scenario was active. "
              "Load Setup puts them all back.")

    def _coerce(self, raw):
        """Rebuild a scenario from file against the current snapshot shape.

        Anything missing keeps the value the app already has, so a setup file
        written by an older build still loads instead of failing outright.
        """
        if not isinstance(raw, dict):
            raise ValueError("a scenario in that file is not an object")
        snap = self.snapshot()
        for key in ("term", "rev", "capex0", "year0_exp", "bonus",
                    "merit", "incentive", "profit"):
            if key in raw:
                snap[key] = float(raw[key])
        if "roster_open" in raw:
            snap["roster_open"] = bool(raw["roster_open"])
        for key in ("rates", "mix"):
            if isinstance(raw.get(key), list):
                vals = [float(v) for v in raw[key][:len(snap[key])]]
                snap[key] = vals + list(snap[key][len(vals):])
        if isinstance(raw.get("roster"), list):
            rows = []
            for i, cur in enumerate(snap["roster"]):
                src = raw["roster"][i] if i < len(raw["roster"]) else None
                if isinstance(src, (list, tuple)) and len(src) >= 3:
                    rows.append((str(src[0]), float(src[1]), float(src[2])))
                else:
                    rows.append(cur)
            snap["roster"] = rows
        if isinstance(raw.get("groups"), dict):
            groups = {}
            for name, cur in snap["groups"].items():
                src = raw["groups"].get(name)
                if isinstance(src, (list, tuple)) and len(src) >= 3:
                    years = [float(v) for v in src[1]][:len(cur[1])]
                    years += list(cur[1][len(years):])
                    groups[name] = (float(src[0]), years, bool(src[2]))
                else:
                    groups[name] = cur
            snap["groups"] = groups
        snap["roi"] = self._roi_text(compute(self.params_from(snap)))
        return snap

    def load_setup(self):
        """Restore every scenario from a file written by Save Setup."""
        path = filedialog.askopenfilename(
            filetypes=[("Hotel ROI Modeling Tool setup", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, ValueError) as e:
            messagebox.showerror(
                APP_TITLE, "Could not read that file:\n" + str(e))
            return
        if not isinstance(data, dict) or data.get("format") != SETUP_FORMAT:
            messagebox.showerror(
                APP_TITLE,
                "That is not a Hotel ROI Modeling Tool setup file.\n\nPick a file written "
                "by Save Setup.")
            return
        raw = data.get("scenarios")
        if not isinstance(raw, list) or not raw:
            messagebox.showerror(APP_TITLE, "That setup file has no scenarios.")
            return
        try:
            loaded = [self._coerce(one) for one in raw[:SCENARIOS]]
        except (TypeError, ValueError, KeyError) as e:
            messagebox.showerror(
                APP_TITLE, "That setup file could not be read:\n" + str(e))
            return
        while len(loaded) < SCENARIOS:      # a short file keeps what is here
            loaded.append(self.snapshot())

        self.scenarios = loaded
        self.active = max(0, min(SCENARIOS - 1, int(data.get("active", 0) or 0)))
        self.note.set(data.get("note", ""))
        self.restore(self.scenarios[self.active])
        self._paint_scenarios()
        messagebox.showinfo(
            APP_TITLE,
            "Loaded\n" + path
            + f"\n\nAll {SCENARIOS} scenarios restored; "
              f"Scenario {self.active + 1} is showing.")

    def export(self):
        """Write a live workbook: formulas, not values, so the reader can
        re-drive the model in Excel."""
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile="roi_model.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        self._save_active()          # the live controls are the active scenario
        try:
            write_workbook(path,
                           [self.params_from(snap) for snap in self.scenarios],
                           active=self.active, note=self.note.get())
        except OSError as e:
            messagebox.showerror(
                APP_TITLE,
                "Could not write the workbook:\n" + str(e)
                + "\n\nIf that file is open in Excel, close it and try again.")
            return
        messagebox.showinfo(
            APP_TITLE,
            "Saved to\n" + path
            + f"\n\nAll {SCENARIOS} scenarios are included, each on its own "
              "sheet with live formulas. Edit any yellow input and Excel "
              "recalculates everything, ROI included.")


if __name__ == "__main__":
    _declare_dpi_aware()
    App().mainloop()
