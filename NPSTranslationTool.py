import json
import random
import re
import sys
import tempfile
from pathlib import Path

# ── tkinterdnd2: must wrap Tk BEFORE creating the root window ────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    tk = None


APP_VERSION   = "1.5.2"

def _get_aliases_file() -> Path:
    import os
    candidates = []
    try:
        candidates.append(Path(sys.argv[0]).resolve().parent)
    except Exception:
        pass
    try:
        candidates.append(Path(__file__).resolve().parent)
    except Exception:
        pass
    for folder in candidates:
        try:
            candidate = folder / "nps_speaker_aliases.json"
            candidate.touch(exist_ok=True)
            return candidate
        except Exception:
            pass
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    cfg_dir = base / "NPSTranslationTool"
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir / "nps_speaker_aliases.json"

_ALIASES_FILE = _get_aliases_file()


def _resource_path(name: str) -> Path:
    """Resolve bundled resource path (PyInstaller) or local file path."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def load_aliases() -> dict:
    try:
        if _ALIASES_FILE.exists():
            return json.loads(_ALIASES_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_aliases(aliases: dict):
    try:
        _ALIASES_FILE.write_text(
            json.dumps(aliases, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        try:
            log = _ALIASES_FILE.with_name("nps_aliases_error.log")
            log.write_text("Save error: " + str(e) + "\nPath: " + str(_ALIASES_FILE) + "\n", encoding="utf-8")
        except Exception:
            pass

# ─────────────────────────────────────────────────────────────────────────────
# Transliteration: Latin → Ukrainian
# ─────────────────────────────────────────────────────────────────────────────
_TRANSLIT_PAIRS = sorted({
    "Shch": "Щ", "shch": "щ",
    "Ye": "Є", "Zh": "Ж", "Yi": "Ї", "Kh": "Х",
    "Ts": "Ц", "Ch": "Ч", "Sh": "Ш", "Shh": "Щ",
    "Yu": "Ю", "Ya": "Я",
    "ye": "є", "zh": "ж", "yi": "ї", "kh": "х",
    "ts": "ц", "ch": "ч", "sh": "ш", "shh": "щ",
    "yu": "ю", "ya": "я",
    "A": "А", "B": "Б", "C": "С", "D": "Д", "E": "Е",
    "F": "Ф", "G": "Г", "H": "Г", "I": "І", "J": "Й",
    "K": "К", "L": "Л", "M": "М", "N": "Н", "O": "О",
    "P": "П", "Q": "К", "R": "Р", "S": "С", "T": "Т",
    "U": "У", "V": "В", "W": "В", "X": "КС", "Y": "И",
    "Z": "З",
    "a": "а", "b": "б", "c": "с", "d": "д", "e": "е",
    "f": "ф", "g": "г", "h": "г", "i": "і", "j": "й",
    "k": "к", "l": "л", "m": "м", "n": "н", "o": "о",
    "p": "п", "q": "к", "r": "р", "s": "с", "t": "т",
    "u": "у", "v": "в", "w": "в", "x": "кс", "y": "и",
    "z": "з",
}.items(), key=lambda x: -len(x[0]))


def transliterate_latin_to_ua(text: str) -> str:
    result, i = [], 0
    while i < len(text):
        matched = False
        for lat, cyr in _TRANSLIT_PAIRS:
            if text[i:i + len(lat)] == lat:
                result.append(cyr)
                i += len(lat)
                matched = True
                break
        if not matched:
            result.append(text[i])
            i += 1
    return "".join(result)


# ─────────────────────────────────────────────────────────────────────────────
# NPS parsing
# ─────────────────────────────────────────────────────────────────────────────
VOICE_LINE_RE  = re.compile(r'(<voice\b[^>]*\bname="([^"]+)"[^>]*>)(.*)$', re.IGNORECASE)
CHOICE_LINE_RE = re.compile(r'(?i)^(?!\s*//)(.*)(<CHOICE\b[^>]*\bTEXT="([^"]*)"[^>]*></A>//?.*)$')
_TRAILING_TAG_RE = re.compile(r'(\s*)(<[^>]+>)(\s*)$')


def split_text_and_tail_tags(text: str):
    """
    Split a line tail into translatable text and trailing control tags.
    Keep italic tags (<I>, </I>, <II>) inside the translatable text.
    """
    keep_inline = {"i", "ii"}
    working = text
    tail_parts = []

    while True:
        m = _TRAILING_TAG_RE.search(working)
        if not m:
            break
        tag = m.group(2)
        nm = re.match(r'</?\s*([A-Za-z0-9_:-]+)', tag)
        if not nm:
            break
        tag_name = nm.group(1).lower()
        if tag_name in keep_inline:
            break

        consumed = m.group(1) + tag + m.group(3)
        tail_parts.insert(0, consumed)
        working = working[:m.start()]

    return working.strip(), "".join(tail_parts)


def split_voice_line(line: str):
    m = VOICE_LINE_RE.search(line)
    if not m:
        return None
    head_full, speaker, rest = m.groups()
    text, tail_tags = split_text_and_tail_tags(rest)
    head = head_full + rest[: len(rest) - len(rest.lstrip())]
    return speaker, head, text, tail_tags


def split_choice_line(line: str):
    stripped = line.strip()
    if stripped.startswith("//"):
        return None
    m = re.search(r'(?i)(<CHOICE\b[^>]*\bTEXT="([^"]*)"[^>]*></A>//?.*)$', line)
    if not m:
        return None
    tag_and_rest = m.group(1)
    text         = m.group(2)
    pre          = line[:m.start(1)]
    return pre, tag_and_rest, text


def split_narration_line(line: str):
    stripped = line.strip()
    if not stripped or stripped.startswith('//'):
        return None
    if stripped.startswith('<') and not re.match(r'(?i)^<\s*/?\s*i\b|^<\s*ii\b', stripped):
        return None

    m = re.match(r'(\s*)(.*)$', line)
    if not m:
        return None
    head_ws, body = m.groups()

    text, tail_tags = split_text_and_tail_tags(body)
    if not text or re.fullmatch(r'(?:\s*<[^>]+>\s*)+', text):
        return None
    return head_ws, text, tail_tags


def build_entries_from_nps(nps_path: Path):
    lines = nps_path.read_text(encoding="utf-8").splitlines()
    entries, entry_id = [], 0
    for lineno, line in enumerate(lines, start=1):
        vd = split_voice_line(line)
        if vd:
            speaker, _h, text, _t = vd
            if text:
                entries.append({"id": entry_id, "type": "voice", "line_no": lineno,
                                "speaker": speaker, "original": text, "translation": ""})
                entry_id += 1
            continue
        cd = split_choice_line(line)
        if cd:
            _pre, _tag, text = cd
            if text:
                entries.append({"id": entry_id, "type": "choice", "line_no": lineno,
                                "speaker": "CHOICE", "original": text, "translation": ""})
                entry_id += 1
            continue
        nd = split_narration_line(line)
        if nd:
            _h, text, _t = nd
            if text:
                entries.append({"id": entry_id, "type": "narration", "line_no": lineno,
                                "speaker": "", "original": text, "translation": ""})
                entry_id += 1
    return entries


def export_to_json_data(nps_path: Path, entries: list) -> dict:
    return {"source_file": str(nps_path), "entries": entries}


def apply_translations_json(nps_path: Path, json_path: Path, out_path):
    data = json.loads(json_path.read_text(encoding="utf-8"))
    entries = sorted(data.get("entries", []), key=lambda r: r.get("id", 0))
    it = iter(entries)
    lines = nps_path.read_text(encoding="utf-8").splitlines()
    new_lines = []
    try:
        current = next(it)
    except StopIteration:
        current = None
    for _lineno, line in enumerate(lines, start=1):
        if current is None:
            new_lines.append(line)
            continue
        vd = split_voice_line(line)
        if vd:
            _, head, _text, tail = vd
            if current and current.get("type") == "voice":
                new_text = (current.get("translation") or "").strip() or current.get("original", "")
                line = f"{head}{new_text}{tail}"
                try:
                    current = next(it)
                except StopIteration:
                    current = None
        else:
            cd = split_choice_line(line)
            if cd:
                pre, tag_and_rest, old_text = cd
                if current and current.get("type") == "choice":
                    new_text = (current.get("translation") or "").strip() or old_text
                    tag_replaced = re.sub(
                        r'TEXT="[^"]*"',
                        'TEXT="' + new_text + '"',
                        tag_and_rest,
                        count=1,
                        flags=re.IGNORECASE,
                    )
                    line = pre + tag_replaced
                    try:
                        current = next(it)
                    except StopIteration:
                        current = None
            else:
                nd = split_narration_line(line)
                if nd:
                    head, _text, tail = nd
                    if current and current.get("type") == "narration":
                        new_text = (current.get("translation") or "").strip() or current.get("original", "")
                        line = f"{head}{new_text}{tail}"
                        try:
                            current = next(it)
                        except StopIteration:
                            current = None
        new_lines.append(line)
    target = out_path if out_path is not None else nps_path
    target.write_text("\n".join(new_lines), encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# Import translation from plain NPS file (no JSON)
# Matches lines by position/order of translatable entries
# ─────────────────────────────────────────────────────────────────────────────
def import_translations_from_nps(translated_nps_path: Path, original_entries: list) -> tuple:
    """
    Parse the translated .nps file and match its translatable lines
    to the original entries by position (index order).
    Returns (updated_entries, matched_count, total_count).
    """
    translated_entries = build_entries_from_nps(translated_nps_path)

    matched = 0
    total = min(len(original_entries), len(translated_entries))

    updated = [e.copy() for e in original_entries]

    for i, orig in enumerate(updated):
        if i >= len(translated_entries):
            break
        tr = translated_entries[i]
        # Only import if the translated text differs from original
        tr_text = tr.get("original", "").strip()
        orig_text = orig.get("original", "").strip()
        if tr_text and tr_text != orig_text:
            orig["translation"] = tr_text
            matched += 1

    return updated, matched, total


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────
def run_gui(initial_path: Path = None):
    if tk is None:
        print("tkinter unavailable.")
        return

    if _DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app_icon_ref = [None]
    try:
        icon_path = _resource_path("Locus-logo.png")
        if icon_path.exists():
            app_icon_ref[0] = tk.PhotoImage(file=str(icon_path))
            root.iconphoto(True, app_icon_ref[0])
    except Exception:
        pass

    root.title("NPS Translation Tool")
    root.minsize(900, 600)

    # ── Colour palette ───────────────────────────────────────────────────────
    BG      = "#0d1117"
    PANEL   = "#161b22"
    SURFACE = "#21262d"
    BORDER  = "#30363d"
    FG      = "#c9d1d9"
    FG_DIM  = "#8b949e"
    ACCENT  = "#2f81f7"
    ACCENT2 = "#238636"
    GREEN   = "#3fb950"
    ORANGE  = "#d29922"
    SEL_BG  = "#1f2a3a"
    BTN_HO  = "#30363d"
    TEAL    = "#1f6feb"
    UI_FONT = "Trebuchet MS"

    root.configure(bg=BG)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background=SURFACE, foreground=FG,
                    fieldbackground=SURFACE, rowheight=25, borderwidth=0,
                    font=(UI_FONT, 10))
    style.map("Treeview",
              background=[("selected", SEL_BG)],
              foreground=[("selected", FG)])
    style.configure("Treeview.Heading", background=PANEL, foreground=FG_DIM,
                    relief="flat", font=(UI_FONT, 10, "bold"))
    style.map("Treeview.Heading", background=[("active", BORDER)])
    for ori in ("Vertical", "Horizontal"):
        style.configure(f"{ori}.TScrollbar", background=PANEL, troughcolor=BG,
                        bordercolor=BG, arrowcolor=FG_DIM, relief="flat")
    style.configure("Accent.Horizontal.TProgressbar",
                    troughcolor=SURFACE, background=ACCENT,
                    bordercolor=SURFACE, lightcolor=ACCENT, darkcolor=ACCENT)

    # ── Widget helpers ───────────────────────────────────────────────────────
    def mk_btn(parent, text, cmd, color=ACCENT, width=None, tooltip=None):
        btn_fg = "#ffffff" if color in (ACCENT, ACCENT2, TEAL, GREEN) else FG
        kw = dict(text=text, command=cmd, bg=color, fg=FG,
                  activebackground=BTN_HO, activeforeground=btn_fg,
                  relief="flat", bd=0, cursor="hand2",
                  padx=10, pady=5, font=(UI_FONT, 10))
        kw["fg"] = btn_fg
        if width:
            kw["width"] = width
        b = tk.Button(parent, **kw)
        if tooltip:
            _add_tooltip(b, tooltip)
        return b

    def mk_label(parent, text, **kw):
        return tk.Label(parent, text=text,
                        bg=kw.pop("bg", BG), fg=kw.pop("fg", FG_DIM),
                        font=(UI_FONT, 10), **kw)

    def mk_entry(parent, **kw):
        # Some Tk builds do not support Entry undo/maxundo options.
        with_undo = dict(kw)
        with_undo.setdefault("undo", True)
        with_undo.setdefault("maxundo", 200)
        try:
            return tk.Entry(parent, **with_undo)
        except tk.TclError:
            return tk.Entry(parent, **kw)

    def _add_tooltip(widget, tip_text):
        tip_win = [None]
        def show(e):
            tip_win[0] = tk.Toplevel(widget)
            tip_win[0].wm_overrideredirect(True)
            tip_win[0].wm_geometry(f"+{e.x_root+12}+{e.y_root+8}")
            tk.Label(tip_win[0], text=tip_text, bg="#30363d", fg="#c9d1d9",
                     relief="solid", bd=1, font=(UI_FONT, 9), padx=4, pady=2).pack()
        def hide(e):
            if tip_win[0]:
                tip_win[0].destroy()
                tip_win[0] = None
        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

    # ── Speaker aliases ──────────────────────────────────────────────────────
    _aliases_box = [load_aliases()]

    def get_display_speaker(original: str) -> str:
        if not original:
            return "narrator"
        return _aliases_box[0].get(original, original)

    # ── App state ────────────────────────────────────────────────────────────
    state = {
        "data": None, "entries": [], "json_path": None,
        "nps_path": None, "current_id": None, "modified": False,
        "entry_history": {}, "history_busy": False,
        "filter": "all",
    }

    def _entry_matches_filter(entry: dict) -> bool:
        mode = state.get("filter", "all")
        has_tr = bool((entry.get("translation") or "").strip())
        if mode == "all":
            return True
        if mode == "todo":
            return not has_tr
        if mode == "done":
            return has_tr
        if mode == "voice":
            return entry.get("type") == "voice"
        if mode == "narration":
            return entry.get("type") == "narration"
        if mode == "choice":
            return entry.get("type") == "choice"
        return True

    def _refresh_window_title():
        fname = ""
        if state.get("nps_path"):
            fname = state["nps_path"].name
        elif state.get("json_path"):
            fname = state["json_path"].name
        dirty = " *" if state.get("modified") else ""
        suffix = f" - {fname}" if fname else ""
        root.title(f"NPS Translation Tool{dirty}{suffix}")

    # ════════════════════════════════════════════════════════════════════════
    # LAYOUT
    # ════════════════════════════════════════════════════════════════════════

    toolbar = tk.Frame(root, bg=PANEL, pady=6)
    toolbar.pack(fill=tk.X)
    tb_left = tk.Frame(toolbar, bg=PANEL)
    tb_left.pack(side=tk.LEFT, padx=10)
    tk.Label(toolbar, text=f"v{APP_VERSION}", bg=PANEL, fg=FG_DIM,
             font=(UI_FONT, 9)).pack(side=tk.RIGHT, padx=10)

    tk.Frame(root, bg=BORDER, height=1).pack(fill=tk.X)
    lbl_status = tk.Label(
        root,
        text="No file loaded — drag a .nps or .json here, or use the buttons above",
        bg=PANEL, fg=FG_DIM, font=(UI_FONT, 9), anchor="w", padx=10, pady=5)
    lbl_status.pack(fill=tk.X)

    main_pane = tk.PanedWindow(root, orient=tk.VERTICAL, bg=BG,
                               sashwidth=5, sashrelief="flat", sashpad=2)
    main_pane.pack(fill=tk.BOTH, expand=True)

    # ── Top: search + table ──────────────────────────────────────────────────
    top_pane = tk.Frame(main_pane, bg=BG)
    main_pane.add(top_pane, stretch="always", minsize=200)

    search_bar = tk.Frame(top_pane, bg=PANEL, pady=5)
    search_bar.pack(fill=tk.X)
    mk_label(search_bar, "🔍", bg=PANEL, fg=FG_DIM).pack(side=tk.LEFT, padx=(8, 2))
    search_var = tk.StringVar()
    search_entry = mk_entry(search_bar, textvariable=search_var, bg=SURFACE, fg=FG,
                            insertbackground=FG, relief="flat", bd=0,
                            font=(UI_FONT, 10), highlightthickness=1,
                            highlightbackground=BORDER, highlightcolor=ACCENT)
    search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
    mk_btn(search_bar, "◀", lambda: select_match(-1), SURFACE).pack(side=tk.LEFT, padx=2)
    mk_btn(search_bar, "▶", lambda: select_match(1),  SURFACE).pack(side=tk.LEFT, padx=2)
    mk_btn(search_bar, "✕",
           lambda: (search_var.set(""), rebuild_search_matches()), SURFACE
           ).pack(side=tk.LEFT, padx=(2, 8))
    search_count_var = tk.StringVar(value="Matches: 0")
    tk.Label(search_bar, textvariable=search_count_var, bg=PANEL, fg=FG_DIM,
             font=(UI_FONT, 9, "bold")).pack(side=tk.RIGHT, padx=(0, 10))

    filter_bar = tk.Frame(top_pane, bg=BG)
    filter_bar.pack(fill=tk.X, padx=8, pady=(6, 4))
    mk_label(filter_bar, "Show:", bg=BG, fg=FG_DIM).pack(side=tk.LEFT, padx=(0, 6))
    filter_buttons = {}

    def _set_filter(mode: str):
        state["filter"] = mode
        _update_filter_buttons()
        rebuild_tree()
        set_status(f"Filter: {mode}")

    def _update_filter_buttons():
        current = state.get("filter", "all")
        for mode, btn in filter_buttons.items():
            if mode == current:
                btn.configure(bg=ACCENT, fg="#ffffff")
            else:
                btn.configure(bg=SURFACE, fg=FG)

    for mode, title in [
        ("all", "All"),
        ("todo", "Untranslated"),
        ("done", "Translated"),
        ("voice", "Voice"),
        ("narration", "Narration"),
        ("choice", "Choices"),
    ]:
        btn = mk_btn(filter_bar, title, lambda m=mode: _set_filter(m), SURFACE)
        btn.pack(side=tk.LEFT, padx=(0, 4))
        filter_buttons[mode] = btn
    _update_filter_buttons()

    tree_frame = tk.Frame(top_pane, bg=BG)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    columns = ("id", "speaker", "original", "translation", "type", "line_no")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                        selectmode="browse", height=15)
    for col, heading, w, anch, stretch in [
        ("id",          "ID",          45,  "center", False),
        ("speaker",     "Speaker",     130, "w",      False),
        ("original",    "Original",    420, "w",      True),
        ("translation", "Translation", 420, "w",      True),
        ("type",        "Type",        75,  "center", False),
        ("line_no",     "Line",        55,  "center", False),
    ]:
        tree.heading(col, text=heading, anchor=anch)
        tree.column(col, width=w, anchor=anch, stretch=stretch)

    tree.tag_configure("odd",      background="#0d1117")
    tree.tag_configure("even",     background="#161b22")
    tree.tag_configure("done",     foreground=GREEN)
    tree.tag_configure("empty",    foreground=FG_DIM)
    tree.tag_configure("narrator", foreground=ORANGE)
    tree.tag_configure("choice",   foreground="#58a6ff")

    _SPEAKER_COLOURS = [
        "#c792ea", "#5a9cf8", "#e8a24f", "#f07178", "#89ddff",
        "#c3e88d", "#ff9cac", "#82aaff", "#ffcb6b", "#b2ccd6",
        "#d4a5ff", "#ff869a", "#80cbc4", "#f78c6c", "#a6accd",
    ]
    _speaker_colour_map: dict = {}

    vsb = ttk.Scrollbar(tree_frame, orient="vertical",
                        command=tree.yview, style="Vertical.TScrollbar")
    tree.configure(yscroll=vsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    tree_frame.rowconfigure(0, weight=1)
    tree_frame.columnconfigure(0, weight=1)

    # ── Bottom: editor ───────────────────────────────────────────────────────
    bot_pane = tk.Frame(main_pane, bg=BG)
    main_pane.add(bot_pane, stretch="never", minsize=190)

    speaker_row = tk.Frame(bot_pane, bg=PANEL, pady=5)
    speaker_row.pack(fill=tk.X)
    mk_label(speaker_row, "Speaker:", bg=PANEL, fg=FG_DIM).pack(side=tk.LEFT, padx=(10, 4))
    speaker_var = tk.StringVar()
    tk.Label(speaker_row, textvariable=speaker_var, bg=PANEL, fg=ACCENT2,
             font=(UI_FONT, 10, "bold"), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
    mk_btn(speaker_row, "✏ Rename Speaker", lambda: open_alias_editor(), SURFACE,
           tooltip="Set display alias for speakers (only visible in app)"
           ).pack(side=tk.RIGHT, padx=8)

    progress_row = tk.Frame(bot_pane, bg=BG)
    progress_row.pack(fill=tk.X, padx=8, pady=(2, 0))
    lbl_info = tk.Label(progress_row, text="", bg=BG, fg=FG_DIM,
                        font=(UI_FONT, 9), anchor="w")
    lbl_info.pack(side=tk.LEFT)
    lbl_progress = tk.Label(progress_row, text="", bg=BG, fg=GREEN,
                            font=(UI_FONT, 9, "bold"), anchor="e")
    lbl_progress.pack(side=tk.RIGHT)
    progress_bar = ttk.Progressbar(bot_pane, mode="determinate", maximum=100,
                                   style="Accent.Horizontal.TProgressbar")
    progress_bar.pack(fill=tk.X, padx=8, pady=(3, 0))

    # Original text block
    orig_frame = tk.Frame(bot_pane, bg=BG)
    orig_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 2))
    orig_header = tk.Frame(orig_frame, bg=BG)
    orig_header.pack(fill=tk.X)
    mk_label(orig_header, "Original", bg=BG, fg=FG_DIM).pack(side=tk.LEFT)
    mk_btn(orig_header, "⧉ Copy",
           lambda: _copy_original_plain(), SURFACE,
           tooltip="Copy original text to clipboard"
           ).pack(side=tk.RIGHT, padx=(4, 0))
    mk_btn(orig_header, "🔤 Translit → Translation",
           lambda: _translit_to_translation_field(), SURFACE,
           tooltip="Transliterate Latin→Ukrainian and insert into Translation field"
           ).pack(side=tk.RIGHT, padx=(4, 0))

    original_text = tk.Text(orig_frame, height=3, wrap="word",
                            bg=SURFACE, fg=FG, insertbackground=FG,
                            relief="flat", bd=0, font=(UI_FONT, 10),
                            padx=8, pady=6, highlightthickness=1,
                            highlightbackground=BORDER, highlightcolor=BORDER,
                            selectbackground=SEL_BG)
    original_text.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
    original_text.configure(state="disabled")

    # Translation text block
    tr_frame = tk.Frame(bot_pane, bg=BG)
    tr_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))
    tr_header = tk.Frame(tr_frame, bg=BG)
    tr_header.pack(fill=tk.X)
    mk_label(tr_header, "Translation", bg=BG, fg=FG_DIM).pack(side=tk.LEFT)
    tk.Label(tr_header, text="↑↓ navigate  •  Enter → next  •  -- → —",
             bg=BG, fg=FG_DIM, font=(UI_FONT, 9)).pack(side=tk.LEFT, padx=(10, 0))

    mk_btn(tr_header, "⧉ Copy",
           lambda: _copy_translation(), SURFACE,
           tooltip="Copy translation to clipboard"
           ).pack(side=tk.RIGHT, padx=(4, 0))
    mk_btn(tr_header, "✕ Clear",
           lambda: translation_text.delete("1.0", tk.END), SURFACE,
           tooltip="Clear translation field"
           ).pack(side=tk.RIGHT, padx=(4, 0))
    mk_btn(tr_header, "↷", lambda: redo_action(), SURFACE, width=3,
           tooltip="Redo (Ctrl+Y / Ctrl+Shift+Z)"
           ).pack(side=tk.RIGHT, padx=(4, 0))
    mk_btn(tr_header, "↶", lambda: undo_action(), SURFACE, width=3,
           tooltip="Undo (Ctrl+Z), up to 200 actions"
           ).pack(side=tk.RIGHT, padx=(4, 0))

    translation_text = tk.Text(tr_frame, height=3, wrap="word",
                               bg=SURFACE, fg=FG, insertbackground=FG,
                               relief="flat", bd=0, font=(UI_FONT, 10),
                               padx=8, pady=6, highlightthickness=1,
                               highlightbackground=BORDER, highlightcolor=ACCENT,
                               selectbackground=SEL_BG)
    translation_text.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

    # ════════════════════════════════════════════════════════════════════════
    # ── Em-dash autocorrect: -- → — ─────────────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════
    def _on_translation_key(event):
        """Replace '--' with '—' when the second '-' is typed."""
        if event.char != "-":
            return
        # Check if the character just before the insertion point is also '-'
        idx = translation_text.index("insert")
        # Get one character before current cursor
        if idx == "1.0":
            return
        prev_idx = translation_text.index(f"insert - 1 chars")
        prev_char = translation_text.get(prev_idx, idx)
        if prev_char == "-":
            # Delete the two dashes and insert em-dash
            # We use edit_separator to keep this as one undoable unit
            translation_text.edit_separator()
            translation_text.delete(prev_idx, idx)
            translation_text.insert(prev_idx, "—")
            translation_text.edit_separator()
            return "break"  # suppress the second '-' being inserted normally

    def _on_translation_separator(event):
        if event.char in (" ", "\t", "\n") or event.keysym in (
                "BackSpace", "Delete", "Return", "KP_Enter"):
            root.after_idle(apply_translation_from_widget)

    translation_text.bind("<Key>", _on_translation_key)
    translation_text.bind("<KeyRelease>", _on_translation_separator)

    # ════════════════════════════════════════════════════════════════════════
    # LOGIC
    # ════════════════════════════════════════════════════════════════════════

    search_state = {"matches": [], "index": -1}

    def set_status(text: str):
        lbl_status.config(text=text)

    def update_progress():
        entries = state["entries"]
        if not entries:
            lbl_progress.config(text="")
            progress_bar["value"] = 0
            return
        done  = sum(1 for e in entries if (e.get("translation") or "").strip())
        total = len(entries)
        pct = int(done / total * 100)
        lbl_progress.config(text=f"Translated: {done}/{total}  ({pct}%)")
        progress_bar["value"] = pct

    def set_entries_from_data(data, json_path, nps_path):
        entries = sorted(data.get("entries", []), key=lambda r: r.get("id", 0))
        state.update(data=data, entries=entries, json_path=json_path,
                     nps_path=nps_path, modified=False, current_id=None)
        state["entry_history"] = {
            e.get("id"): {"undo": [e.get("translation", "")], "redo": []}
            for e in entries
        }
        state["history_busy"] = False
        rebuild_tree()
        fname = (nps_path.name if nps_path else None) or (json_path.name if json_path else "")
        _refresh_window_title()
        set_status(f"Loaded: {fname}  |  {len(entries)} entries")
        update_progress()

    # ── Navigation ───────────────────────────────────────────────────────────
    def navigate_entry(direction: int):
        """Move to the previous (-1) or next (+1) entry in the tree."""
        apply_translation_from_widget()
        children = tree.get_children()
        if not children:
            return
        sel = tree.selection()
        if sel:
            try:
                idx = list(children).index(sel[0])
            except ValueError:
                idx = -1
        else:
            idx = -1

        new_idx = idx + direction
        if new_idx < 0:
            new_idx = 0
        elif new_idx >= len(children):
            new_idx = len(children) - 1

        if new_idx == idx:
            return

        iid = children[new_idx]
        tree.selection_set(iid)
        tree.focus(iid)
        tree.see(iid)
        on_tree_select()
        update_progress()

    # ── File loading ─────────────────────────────────────────────────────────
    def load_json(path=None):
        if path is None:
            fn = filedialog.askopenfilename(
                title="Open JSON",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
            if not fn:
                return
            path = Path(fn)
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read JSON:\n{e}")
            return
        src = data.get("source_file")
        set_entries_from_data(data, path, Path(src) if src else None)

    def open_file_dialog():
        fn = filedialog.askopenfilename(
            title="Open .nps or .json",
            filetypes=[
                ("NPS / JSON files", "*.nps *.json"),
                ("NPS files", "*.nps"),
                ("JSON files", "*.json"),
                ("All files", "*.*"),
            ])
        if not fn:
            return
        open_any_file(Path(fn))

    def load_nps(path=None):
        if path is None:
            fn = filedialog.askopenfilename(
                title="Open .nps",
                filetypes=[("NPS files", "*.nps"), ("All files", "*.*")])
            if not fn:
                return
            path = Path(fn)
        try:
            entries = build_entries_from_nps(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read NPS:\n{e}")
            return
        json_path = path.with_suffix(".json")
        if json_path.exists():
            try:
                saved     = json.loads(json_path.read_text(encoding="utf-8"))
                saved_map = {e["id"]: e for e in saved.get("entries", []) if "id" in e}
                for e in entries:
                    se = saved_map.get(e["id"])
                    if se:
                        e["translation"] = se.get("translation", "")
            except Exception as ex:
                messagebox.showwarning("Warning", f"Couldn't read existing JSON:\n{ex}")
        data = export_to_json_data(path, entries)
        try:
            json_path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write JSON:\n{e}")
            return
        set_entries_from_data(data, json_path, path)

    def open_any_file(path: Path):
        suf = path.suffix.lower()
        if suf == ".json":
            load_json(path)
        elif suf == ".nps":
            load_nps(path)
        else:
            messagebox.showwarning(
                "Unsupported", f"Cannot open: {path.name}\nSupported: .nps, .json")

    # ── Import translation from plain NPS ────────────────────────────────────
    def import_translation_from_nps():
        """
        Open a translated .nps file (no JSON), match entries by position,
        and fill in translation fields. Shows a preview/confirmation dialog.
        """
        if not state["entries"]:
            messagebox.showinfo("Import Translation",
                                "Please open an original .nps or .json file first.")
            return

        fn = filedialog.askopenfilename(
            title="Open translated .nps file",
            filetypes=[("NPS files", "*.nps"), ("All files", "*.*")])
        if not fn:
            return
        tr_path = Path(fn)

        try:
            updated, matched, total = import_translations_from_nps(tr_path, state["entries"])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import translations:\n{e}")
            return

        if matched == 0:
            messagebox.showinfo(
                "Import Translation",
                f"No differences found between the files.\n"
                f"Compared {total} entries — all texts are identical.\n\n"
                f"Make sure you selected the TRANSLATED file, not the original.")
            return

        # Preview dialog
        win = tk.Toplevel(root)
        win.title("Import Translation — Preview")
        win.configure(bg=BG)
        win.grab_set()
        win.resizable(True, True)
        win.minsize(700, 400)

        tk.Label(win,
                 text=f"Found {matched} translated lines out of {total} compared entries.",
                 bg=BG, fg=GREEN, font=("Segoe UI", 10, "bold")
                 ).pack(padx=16, pady=(12, 2))
        tk.Label(win,
                 text="Preview of changes (Original → Translation). Scroll to review.",
                 bg=BG, fg=FG_DIM, font=("Segoe UI", 9)
                 ).pack(padx=16, pady=(0, 6))
        tk.Frame(win, bg=BORDER, height=1).pack(fill=tk.X, padx=16)

        # Scrollable preview
        pf = tk.Frame(win, bg=BG)
        pf.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        preview_tree = ttk.Treeview(pf,
                                    columns=("id", "original", "translation"),
                                    show="headings", height=14)
        preview_tree.heading("id",          text="ID",          anchor="center")
        preview_tree.heading("original",    text="Original",    anchor="w")
        preview_tree.heading("translation", text="Translation", anchor="w")
        preview_tree.column("id",          width=50,  anchor="center", stretch=False)
        preview_tree.column("original",    width=310, anchor="w",      stretch=True)
        preview_tree.column("translation", width=310, anchor="w",      stretch=True)
        preview_tree.tag_configure("changed", foreground=GREEN)
        preview_tree.tag_configure("same",    foreground=FG_DIM)

        pvb = ttk.Scrollbar(pf, orient="vertical", command=preview_tree.yview)
        preview_tree.configure(yscroll=pvb.set)
        preview_tree.grid(row=0, column=0, sticky="nsew")
        pvb.grid(row=0, column=1, sticky="ns")
        pf.rowconfigure(0, weight=1)
        pf.columnconfigure(0, weight=1)

        for orig_e, upd_e in zip(state["entries"], updated):
            tr = upd_e.get("translation", "").strip()
            orig_txt = orig_e.get("original", "")
            tag = "changed" if tr and tr != orig_txt else "same"
            preview_tree.insert("", tk.END,
                                values=(orig_e.get("id"), orig_txt, tr or "—"),
                                tags=(tag,))

        tk.Frame(win, bg=BORDER, height=1).pack(fill=tk.X, padx=16)

        # Overwrite options
        opt_frame = tk.Frame(win, bg=BG)
        opt_frame.pack(fill=tk.X, padx=16, pady=6)
        overwrite_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            opt_frame,
            text="Overwrite existing translations (uncheck = only fill empty fields)",
            variable=overwrite_var,
            bg=BG, fg=FG, selectcolor=SURFACE,
            activebackground=BG, activeforeground=FG,
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        confirmed = [False]

        def apply_import():
            confirmed[0] = True
            overwrite = overwrite_var.get()
            apply_translation_from_widget()  # save current editor state first

            for orig_e, upd_e in zip(state["entries"], updated):
                new_tr = upd_e.get("translation", "").strip()
                if not new_tr:
                    continue
                existing = (orig_e.get("translation") or "").strip()
                if existing and not overwrite:
                    continue
                orig_e["translation"] = new_tr

            state["modified"] = True
            state["history_busy"] = True
            try:
                rebuild_tree()
            finally:
                state["history_busy"] = False
            update_progress()
            _refresh_window_title()
            set_status(f"Import complete — {matched} translations applied from {tr_path.name}")
            win.destroy()

        btn_row = tk.Frame(win, bg=BG)
        btn_row.pack(pady=10)
        mk_btn(btn_row, "✔ Apply Import", apply_import, GREEN).pack(side=tk.LEFT, padx=6)
        mk_btn(btn_row, "✕ Cancel",       win.destroy,  SURFACE).pack(side=tk.LEFT, padx=6)

        win.update_idletasks()
        w = max(win.winfo_reqwidth(), 740)
        h = max(win.winfo_reqheight(), 480)
        rx = root.winfo_rootx() + (root.winfo_width()  - w) // 2
        ry = root.winfo_rooty() + (root.winfo_height() - h) // 2
        win.geometry(f"{w}x{h}+{rx}+{ry}")

    # ── Drag & Drop ──────────────────────────────────────────────────────────
    def _on_drop(event):
        raw = event.data.strip()
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        first = next((a or b for a, b in paths), None)
        if first:
            open_any_file(Path(first))

    if _DND_AVAILABLE:
        for w in (root, tree, original_text, translation_text):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", _on_drop)
            except Exception:
                pass
    else:
        set_status(
            "Drag & drop not available.  "
            "Run:  pip install tkinterdnd2  then restart the app.")

    # ── Quick save ───────────────────────────────────────────────────────────
    def quick_save_json():
        if not state["data"]:
            messagebox.showinfo("Info", "No file loaded.")
            return
        if state["json_path"] is None:
            fn = filedialog.asksaveasfilename(
                title="Save JSON", defaultextension=".json",
                filetypes=[("JSON files", "*.json")])
            if not fn:
                return
            state["json_path"] = Path(fn)
        apply_translation_from_widget()
        state["data"]["entries"] = state["entries"]
        state["json_path"].write_text(
            json.dumps(state["data"], ensure_ascii=False, indent=2), encoding="utf-8")
        state["modified"] = False
        _refresh_window_title()
        set_status(f"Saved → {state['json_path'].name}")
        update_progress()

    # ── Save translated NPS ──────────────────────────────────────────────────
    def save_translated_nps():
        if not state["nps_path"] or not state["entries"]:
            messagebox.showinfo("Info", "Open a .nps file first.")
            return
        apply_translation_from_widget()
        nps_path        = state["nps_path"]
        translated_path = nps_path.with_name(f"translated_{nps_path.name}")
        tmp             = Path("__tmp_nps_translator__.json")
        try:
            data = export_to_json_data(nps_path, state["entries"])
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            apply_translations_json(nps_path, tmp, translated_path)
            if state["json_path"] is None:
                state["json_path"] = nps_path.with_suffix(".json")
            state["data"] = data
            state["json_path"].write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            state["modified"] = False
            _refresh_window_title()
            set_status(f"Saved → {translated_path.name}")
            messagebox.showinfo("Done", f"Saved as:\n{translated_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            if tmp.exists():
                try:
                    tmp.unlink()
                except Exception:
                    pass

    # ── Counter ──────────────────────────────────────────────────────────────
    def run_counter():
        filenames = filedialog.askopenfilenames(
            title="Select .nps files",
            filetypes=[("NPS files", "*.nps"), ("All files", "*.*")])
        if not filenames:
            return
        total_e = total_w = 0
        per_file = []
        for name in filenames:
            p = Path(name)
            try:
                ents = build_entries_from_nps(p)
            except Exception as ex:
                messagebox.showwarning("Warning", f"Failed: {p.name}\n{ex}")
                continue
            fl = len(ents)
            fw = sum(len(e.get("original", "").split()) for e in ents)
            total_e += fl
            total_w += fw
            per_file.append((p.name, fl, fw, ents))
        if not total_e:
            messagebox.showinfo("Result", "No translatable lines found.")
            return
        out_name = filedialog.asksaveasfilename(
            title="Save report", defaultextension=".txt",
            filetypes=[("Text files", "*.txt")])
        if not out_name:
            return
        lines = [f"Total lines: {total_e}", f"Total words: {total_w}", ""]
        for fname, fl, fw, ents in per_file:
            lines.append(f"[{fname}]  lines={fl}, words={fw}")
            for e in ents:
                lines.append(
                    f"{e.get('id')}\t{e.get('speaker','') or 'narrator'}\t{e.get('original','')}")
            lines.append("")
        Path(out_name).write_text("\n".join(lines), encoding="utf-8")
        messagebox.showinfo("Done", f"Lines: {total_e}\nWords: {total_w}")

    # ── Speaker alias editor ─────────────────────────────────────────────────
    def open_alias_editor():
        entries = state["entries"]
        if not entries:
            messagebox.showinfo("Info", "No file loaded.")
            return
        speakers_orig = sorted(
            {e.get("speaker", "") for e in entries if e.get("speaker", "")})
        if not speakers_orig:
            messagebox.showinfo("Info", "No named speakers found (only narration lines).")
            return

        win = tk.Toplevel(root)
        win.title("Rename Speakers")
        win.configure(bg=BG)
        win.grab_set()
        win.resizable(False, False)

        tk.Label(
            win,
            text="Speaker display aliases  —  only visible in this app, never saved to file",
            bg=BG, fg=FG_DIM, font=("Segoe UI", 9)
        ).pack(padx=16, pady=(12, 2))
        tk.Label(
            win,
            text="Original name → Display alias   (leave blank to keep original)",
            bg=BG, fg="#555566", font=("Segoe UI", 8)
        ).pack(padx=16, pady=(0, 6))
        tk.Frame(win, bg=BORDER, height=1).pack(fill=tk.X, padx=16)

        list_frame = tk.Frame(win, bg=BG)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        alias_vars: dict = {}
        for sp in speakers_orig:
            row = tk.Frame(list_frame, bg=BG)
            row.pack(fill=tk.X, pady=4)
            tk.Label(row, text=sp, bg=BG, fg=FG,
                     font=("Segoe UI", 9), width=30, anchor="w"
                     ).pack(side=tk.LEFT, padx=(0, 8))
            var = tk.StringVar(value=_aliases_box[0].get(sp, ""))
            mk_entry(row, textvariable=var, bg=SURFACE, fg=FG,
                     insertbackground=FG, relief="flat",
                     font=("Segoe UI", 9), width=30,
                     highlightthickness=1,
                     highlightbackground=BORDER,
                     highlightcolor=ACCENT
                     ).pack(side=tk.LEFT)
            alias_vars[sp] = var

        tk.Frame(win, bg=BORDER, height=1).pack(fill=tk.X, padx=16)

        def apply_aliases():
            aliases = _aliases_box[0]
            for sp, var in alias_vars.items():
                alias = var.get().strip()
                if alias:
                    aliases[sp] = alias
                else:
                    aliases.pop(sp, None)
            save_aliases(aliases)
            rebuild_tree()
            on_tree_select()
            win.destroy()

        btn_row = tk.Frame(win, bg=BG)
        btn_row.pack(pady=10)
        mk_btn(btn_row, "✔ Apply",  apply_aliases, ACCENT).pack(side=tk.LEFT, padx=6)
        mk_btn(btn_row, "✕ Cancel", win.destroy,   SURFACE).pack(side=tk.LEFT, padx=6)

        win.update_idletasks()
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()
        rx = root.winfo_rootx() + (root.winfo_width()  - w) // 2
        ry = root.winfo_rooty() + (root.winfo_height() - h) // 2
        win.geometry(f"{max(w, 520)}x{max(h, 120)}+{rx}+{ry}")

    # ── Copy / Translit helpers ───────────────────────────────────────────────
    def _get_original_text() -> str:
        return original_text.get("1.0", tk.END).rstrip("\n")

    def _copy_original_plain():
        txt = _get_original_text()
        root.clipboard_clear()
        root.clipboard_append(txt)
        set_status("Copied original to clipboard")

    def _translit_to_translation_field():
        txt = _get_original_text()
        if not txt:
            return
        result = transliterate_latin_to_ua(txt)
        translation_text.delete("1.0", tk.END)
        translation_text.insert(tk.END, result)
        translation_text.focus_set()
        set_status("Transliteration inserted into Translation field")

    def _copy_translation():
        txt = translation_text.get("1.0", tk.END).rstrip("\n")
        root.clipboard_clear()
        root.clipboard_append(txt)
        set_status("Copied translation to clipboard")

    EASTER_EGG_CHANCE = 0.02

    _saya_popup = {
        "img": None,
        "win": None,
        "label": None,
        "anim_after": None,
        "trans_key": "#01ff01",
        "load_error_shown": False,
        "tmp_paths": [],
    }

    def _load_saya_image():
        png_pic = _resource_path("saya-pic.png")
        gif_pic = _resource_path("saya-pic.gif")
        # Prefer PNG because it preserves alpha transparency.
        pic = png_pic if png_pic.exists() else gif_pic
        if not pic.exists():
            return None, "saya-pic.gif / saya-pic.png not found"

        errors = []

        try:
            # Prefer native Tk image loader first.
            img = tk.PhotoImage(file=str(pic.resolve()))
            max_dim = max(img.width(), img.height())
            if max_dim > 140:
                factor = max(1, max_dim // 120)
                img = img.subsample(factor, factor)
            return img, None
        except Exception as tk_err:
            errors.append(f"Tk PNG: {tk_err}")

        # Fallback: convert via Pillow to formats Tk can decode reliably.
        # GIF/PPM work on older Tk builds where PNG support is unavailable.
        try:
            from PIL import Image  # type: ignore
            pil_img = Image.open(pic).convert("RGBA")
            pil_img.thumbnail((140, 140))

            temp_root = Path(tempfile.gettempdir())
            for ext, fmt, save_img in (
                ("gif", "GIF", pil_img.convert("P", palette=Image.ADAPTIVE)),
                ("ppm", "PPM", pil_img.convert("RGB")),
                ("png", "PNG", pil_img),
            ):
                out = temp_root / f"nps_saya_popup_tmp.{ext}"
                try:
                    save_img.save(out, format=fmt)
                    img = tk.PhotoImage(file=str(out))
                    _saya_popup["tmp_paths"].append(out)
                    return img, None
                except Exception as fmt_err:
                    errors.append(f"Tk {ext.upper()}: {fmt_err}")
        except Exception as pil_err:
            errors.append(f"PIL open/convert: {pil_err}")

        # Last resort: no image support; caller will show text badge.
        return None, "; ".join(errors)

    def _cancel_saya_anim():
        aid = _saya_popup.get("anim_after")
        if aid is not None:
            try:
                root.after_cancel(aid)
            except Exception:
                pass
            _saya_popup["anim_after"] = None

    def _ensure_saya_overlay():
        if _saya_popup["win"] is not None:
            return
        win = tk.Toplevel(root)
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        trans_key = _saya_popup["trans_key"]
        win.configure(bg=trans_key)
        try:
            win.wm_attributes("-transparentcolor", trans_key)
        except Exception:
            # Fallback for Tk builds without transparent color support.
            win.configure(bg=BG)

        label = tk.Label(
            win,
            bg=trans_key,
            fg=FG,
            bd=0,
            relief="flat",
            highlightthickness=0,
            padx=0,
            pady=0,
            font=(UI_FONT, 9, "bold"),
        )
        label.pack()
        _saya_popup["win"] = win
        _saya_popup["label"] = label

    def maybe_show_saya_popup(has_translation: bool):
        # Mini easter egg: appears only sometimes after Enter on non-empty text.
        if not has_translation or random.random() > EASTER_EGG_CHANCE:
            return
        if _saya_popup["img"] is None:
            img, err = _load_saya_image()
            _saya_popup["img"] = img
            if err and not _saya_popup["load_error_shown"]:
                _saya_popup["load_error_shown"] = True
                set_status(f"Saya image unavailable, using text badge ({err})")

        _ensure_saya_overlay()
        win = _saya_popup["win"]
        lbl = _saya_popup["label"]
        if win is None or lbl is None:
            return

        trans_key = _saya_popup["trans_key"]
        if _saya_popup["img"] is not None:
            lbl.configure(image=_saya_popup["img"], text="", bg=trans_key)
            lbl.image = _saya_popup["img"]
            sprite_w = _saya_popup["img"].width()
            sprite_h = _saya_popup["img"].height()
        else:
            lbl.configure(image="", text="Saya", bg=SURFACE, bd=1, relief="solid", padx=6, pady=2)
            lbl.image = None
            win.update_idletasks()
            sprite_w = max(48, lbl.winfo_reqwidth())
            sprite_h = max(18, lbl.winfo_reqheight())

        _cancel_saya_anim()

        root.update_idletasks()
        rx = root.winfo_rootx()
        ry = root.winfo_rooty()
        rw = root.winfo_width()
        rh = root.winfo_height()

        min_x = rx + 10
        max_x = rx + max(10, rw - sprite_w - 10)
        x = random.randint(min_x, max_x)
        start_y = ry + rh + sprite_h + 4
        peak_y = max(ry + 14, ry + rh - sprite_h - 120)
        end_y = ry + rh + sprite_h + 28

        win.geometry(f"{sprite_w}x{sprite_h}+{x}+{start_y}")
        win.deiconify()

        jump_frames = 10
        fall_frames = 14
        frame_ms = 16

        def ease_out_quad(t: float) -> float:
            return 1.0 - (1.0 - t) * (1.0 - t)

        def ease_in_quad(t: float) -> float:
            return t * t

        def step_jump(i: int):
            t = i / max(1, jump_frames)
            k = ease_out_quad(t)
            y = int(start_y + (peak_y - start_y) * k)
            win.geometry(f"{sprite_w}x{sprite_h}+{x}+{y}")
            if i < jump_frames:
                _saya_popup["anim_after"] = root.after(frame_ms, step_jump, i + 1)
            else:
                _saya_popup["anim_after"] = root.after(frame_ms, step_fall, 0)

        def step_fall(i: int):
            t = i / max(1, fall_frames)
            k = ease_in_quad(t)
            y = int(peak_y + (end_y - peak_y) * k)
            win.geometry(f"{sprite_w}x{sprite_h}+{x}+{y}")
            if i < fall_frames:
                _saya_popup["anim_after"] = root.after(frame_ms, step_fall, i + 1)
            else:
                _saya_popup["anim_after"] = None
                win.withdraw()

        step_jump(0)

    # ── Tree ─────────────────────────────────────────────────────────────────
    def _ensure_speaker_tag(orig_spk: str):
        if not orig_spk:
            return
        tag = f"spk_{orig_spk}"
        if tag not in _speaker_colour_map:
            idx    = len(_speaker_colour_map) % len(_SPEAKER_COLOURS)
            colour = _SPEAKER_COLOURS[idx]
            _speaker_colour_map[tag] = colour
            tree.tag_configure(tag, foreground=colour)

    def rebuild_tree():
        current_id = state.get("current_id")
        selected_iid = None
        first_iid = None
        for item in tree.get_children():
            tree.delete(item)
        for i, e in enumerate(state["entries"]):
            if not _entry_matches_filter(e):
                continue
            orig_spk  = e.get("speaker", "")
            disp_spk  = get_display_speaker(orig_spk)
            has_tr    = bool((e.get("translation") or "").strip())
            stripe    = "even" if i % 2 == 0 else "odd"

            if has_tr:
                color_tag = "done"
            elif not orig_spk:
                color_tag = "narrator"
            elif orig_spk == "CHOICE":
                color_tag = "choice"
            else:
                _ensure_speaker_tag(orig_spk)
                color_tag = f"spk_{orig_spk}"

            iid = tree.insert("", tk.END,
                              values=(e.get("id"), disp_spk, e.get("original", ""),
                                      e.get("translation", ""), e.get("type", ""),
                                      e.get("line_no", "")),
                              tags=(stripe, color_tag))
            if first_iid is None:
                first_iid = iid
            if current_id is not None and e.get("id") == current_id:
                selected_iid = iid

        if selected_iid is not None:
            tree.selection_set(selected_iid)
            tree.focus(selected_iid)
            tree.see(selected_iid)
        elif first_iid is not None:
            tree.selection_set(first_iid)
            tree.focus(first_iid)
            tree.see(first_iid)
        else:
            state["current_id"] = None
        on_tree_select()
        update_progress()

    def _set_entry_translation(entry_id: int, txt: str):
        entries = state["entries"]
        for e in entries:
            if e.get("id") == entry_id:
                if e.get("translation") != txt:
                    e["translation"] = txt
                    state["modified"] = True
                    _refresh_window_title()
                    for iid in tree.get_children():
                        vals = tree.item(iid, "values")
                        if int(vals[0]) == entry_id:
                            tree.set(iid, "translation", txt)
                            stripe   = tree.item(iid, "tags")[0]
                            orig_spk = e.get("speaker", "")
                            if txt.strip():
                                ctag = "done"
                            elif not orig_spk:
                                ctag = "narrator"
                            elif orig_spk == "CHOICE":
                                ctag = "choice"
                            else:
                                _ensure_speaker_tag(orig_spk)
                                ctag = f"spk_{orig_spk}"
                            tree.item(iid, tags=(stripe, ctag))
                            break
                break

    def _ensure_entry_history(entry_id: int, initial_text: str = ""):
        hist_map = state["entry_history"]
        if entry_id not in hist_map:
            hist_map[entry_id] = {"undo": [initial_text], "redo": []}
        if not hist_map[entry_id]["undo"]:
            hist_map[entry_id]["undo"] = [initial_text]

    def apply_translation_from_widget():
        if state["history_busy"]:
            return
        entries = state["entries"]
        cur_id  = state.get("current_id")
        if not entries or cur_id is None:
            return
        txt = translation_text.get("1.0", tk.END).rstrip("\n")
        for e in entries:
            if e.get("id") == cur_id:
                old = e.get("translation") or ""
                if old != txt:
                    _ensure_entry_history(cur_id, old)
                    hist = state["entry_history"][cur_id]
                    hist["undo"].append(txt)
                    if len(hist["undo"]) > 200:
                        hist["undo"].pop(0)
                    hist["redo"].clear()
                    _set_entry_translation(cur_id, txt)
                break

    def _apply_history_value(entry_id: int, new_txt: str):
        state["history_busy"] = True
        try:
            _set_entry_translation(entry_id, new_txt)
            if state.get("current_id") == entry_id:
                translation_text.delete("1.0", tk.END)
                translation_text.insert(tk.END, new_txt)
                translation_text.focus_set()
        finally:
            state["history_busy"] = False
        update_progress()

    def undo_action():
        apply_translation_from_widget()
        cur_id = state.get("current_id")
        if cur_id is None:
            return "break"
        current_text = next(
            (e.get("translation", "") for e in state["entries"] if e.get("id") == cur_id),
            "",
        )
        _ensure_entry_history(cur_id, current_text)
        hist = state["entry_history"][cur_id]
        if len(hist["undo"]) <= 1:
            return "break"
        current = hist["undo"].pop()
        hist["redo"].append(current)
        _apply_history_value(cur_id, hist["undo"][-1])
        set_status("Undo applied")
        return "break"

    def redo_action():
        apply_translation_from_widget()
        cur_id = state.get("current_id")
        if cur_id is None:
            return "break"
        current_text = next(
            (e.get("translation", "") for e in state["entries"] if e.get("id") == cur_id),
            "",
        )
        _ensure_entry_history(cur_id, current_text)
        hist = state["entry_history"][cur_id]
        if not hist["redo"]:
            return "break"
        next_text = hist["redo"].pop()
        hist["undo"].append(next_text)
        if len(hist["undo"]) > 200:
            hist["undo"].pop(0)
        _apply_history_value(cur_id, next_text)
        set_status("Redo applied")
        return "break"

    def on_tree_select(_event=None):
        apply_translation_from_widget()
        entries = state["entries"]
        if not entries:
            lbl_info.config(text="")
            speaker_var.set("")
            original_text.configure(state="normal")
            original_text.delete("1.0", tk.END)
            original_text.configure(state="disabled")
            translation_text.delete("1.0", tk.END)
            return
        sel = tree.selection()
        if not sel:
            state["current_id"] = None
            lbl_info.config(text="No row selected")
            speaker_var.set("")
            original_text.configure(state="normal")
            original_text.delete("1.0", tk.END)
            original_text.configure(state="disabled")
            translation_text.delete("1.0", tk.END)
            return
        vals     = tree.item(sel[0], "values")
        entry_id = int(vals[0])
        state["current_id"] = entry_id
        e = next((x for x in entries if x.get("id") == entry_id), None)
        if not e:
            return
        lbl_info.config(
            text=f"ID: {e.get('id')}  |  type: {e.get('type')}  |  line: {e.get('line_no')}")
        orig_spk = e.get("speaker", "")
        speaker_var.set(get_display_speaker(orig_spk) if orig_spk else "narrator")
        original_text.configure(state="normal")
        original_text.delete("1.0", tk.END)
        original_text.insert(tk.END, e.get("original", ""))
        original_text.configure(state="disabled")
        translation_text.delete("1.0", tk.END)
        translation_text.insert(tk.END, e.get("translation") or "")
        _ensure_entry_history(entry_id, e.get("translation") or "")
        translation_text.focus_set()

    # ── Search ───────────────────────────────────────────────────────────────
    def rebuild_search_matches():
        q = search_var.get().strip().lower()
        search_state["matches"] = []
        search_state["index"]   = -1
        if not q:
            search_count_var.set("Matches: 0")
            set_status("Search cleared")
            return
        matches = [
            iid for iid in tree.get_children()
            if q in " ".join(str(v) for v in tree.item(iid, "values")[1:4]).lower()
        ]
        search_state["matches"] = matches
        search_count_var.set(f"Matches: {len(matches)}")
        set_status(
            f"Found {len(matches)} match(es) for: '{q}'"
            if matches else f"No matches for: '{q}'")

    def select_match(direction: int):
        if not search_state["matches"]:
            rebuild_search_matches()
        if not search_state["matches"]:
            return
        search_state["index"] = (
            (search_state["index"] + direction) % len(search_state["matches"]))
        iid = search_state["matches"][search_state["index"]]
        tree.selection_set(iid)
        tree.focus(iid)
        tree.see(iid)
        on_tree_select()

    search_entry.bind("<KeyRelease>", lambda e: rebuild_search_matches())
    search_entry.bind("<Return>",
                      lambda e: (rebuild_search_matches(), select_match(1)) or "break")

    def focus_search(_event=None):
        search_entry.focus_set()
        search_entry.selection_range(0, tk.END)
        return "break"

    def clear_search(_event=None):
        if search_var.get():
            search_var.set("")
            rebuild_search_matches()
            return "break"
        return None

    def find_next_match(_event=None):
        select_match(-1 if (_event and _event.state & 0x1) else 1)
        return "break"

    # ── Keyboard shortcuts ───────────────────────────────────────────────────
    def on_ctrl_key(event):
        if not (event.state & 0x4):
            return
        w    = root.focus_get()
        code = event.keycode
        key  = (event.keysym or "").lower()
        if isinstance(w, (tk.Text, tk.Entry)):
            for k, ev in ((67, "<<Copy>>"), (86, "<<Paste>>"), (88, "<<Cut>>")):
                if code == k:
                    w.event_generate(ev)
                    return "break"
            if code == 65 or key == "a":
                if isinstance(w, tk.Text):
                    w.tag_add("sel", "1.0", "end-1c")
                else:
                    w.selection_range(0, tk.END)
                return "break"
        if code == 83 or key == "s":      # Ctrl+S
            quick_save_json()
            return "break"
        # Ctrl+Z / Ctrl+Shift+Z / Ctrl+Y — native text undo/redo
        if code == 90 or key == "z":
            if event.state & 0x1:  # Shift held
                return redo_action()
            return undo_action()
        if code == 89 or key == "y":
            return redo_action()

    root.bind_all("<KeyPress>", on_ctrl_key)
    root.bind("<Control-f>", focus_search)
    root.bind("<Escape>", clear_search)
    root.bind("<F3>", find_next_match)

    # ── Arrow key navigation ─────────────────────────────────────────────────
    # Up/Down in the translation Text widget → move between entries
    def on_up_in_translation(event):
        navigate_entry(-1)
        return "break"

    def on_down_in_translation(event):
        navigate_entry(1)
        return "break"

    translation_text.bind("<Up>",   on_up_in_translation)
    translation_text.bind("<Down>", on_down_in_translation)

    # Up/Down on the tree → sync editor panel after Treeview moves its cursor
    def on_tree_arrow(event):
        root.after(10, on_tree_select)

    tree.bind("<Up>",   on_tree_arrow)
    tree.bind("<Down>", on_tree_arrow)

    def on_enter_in_translation(event):
        has_translation = bool(translation_text.get("1.0", tk.END).strip())
        maybe_show_saya_popup(has_translation)
        navigate_entry(1)
        return "break"

    translation_text.bind(
        "<FocusOut>", lambda e: (apply_translation_from_widget(), update_progress()))
    translation_text.bind("<Return>",   on_enter_in_translation)
    translation_text.bind("<KP_Enter>", on_enter_in_translation)
    tree.bind("<<TreeviewSelect>>", on_tree_select)

    # ── Toolbar buttons ──────────────────────────────────────────────────────
    mk_btn(tb_left, "📂 Open File", open_file_dialog,              ACCENT
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "💾 Quick Save", quick_save_json,             ACCENT2,
           tooltip="Ctrl+S"
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "📄 Save .nps",  save_translated_nps,         SURFACE
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "📥 Import Translation", import_translation_from_nps, TEAL,
           tooltip="Import translations from a plain translated .nps file"
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "🔢 Counter",    run_counter,                 SURFACE
           ).pack(side=tk.LEFT, padx=(0, 4))

    # ── Window close ─────────────────────────────────────────────────────────
    def on_close():
        if state["modified"] and messagebox.askyesno(
                "Exit", "Unsaved changes. Quick Save before exit?"):
            quick_save_json()
        _cancel_saya_anim()
        if _saya_popup.get("win") is not None:
            try:
                _saya_popup["win"].destroy()
            except Exception:
                pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    if initial_path is not None:
        open_any_file(initial_path)

    root.mainloop()


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    initial = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    run_gui(initial)