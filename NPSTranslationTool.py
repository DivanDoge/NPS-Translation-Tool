import json
import re
import sys
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


APP_VERSION   = "1.2.0"

def _get_aliases_file() -> Path:
    """
    Returns a stable path for nps_speaker_aliases.json.
    For PyInstaller --onefile, sys.argv[0] is the real .exe path.
    For plain .py, __file__ is used. Falls back to APPDATA.
    """
    import os

    candidates = []

    # 1. Real exe location (works for both PyInstaller --onefile and plain launch)
    try:
        candidates.append(Path(sys.argv[0]).resolve().parent)
    except Exception:
        pass

    # 2. Script file location
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

    # 3. Fallback: platform config dir
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    cfg_dir = base / "NPSTranslationTool"
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir / "nps_speaker_aliases.json"

_ALIASES_FILE = _get_aliases_file()


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
        # Write error to a log file next to aliases file so user can inspect
        try:
            log = _ALIASES_FILE.with_name("nps_aliases_error.log")
            log.write_text("Save error: " + str(e) + "\nPath: " + str(_ALIASES_FILE) + "\n", encoding="utf-8")
        except Exception:
            pass

# ─────────────────────────────────────────────────────────────────────────────
# Transliteration: Latin → Ukrainian
# ─────────────────────────────────────────────────────────────────────────────
_TRANSLIT_PAIRS = sorted({
    # ── 4-char sequences (must come before shorter ones) ──────────────────
    "Shch": "Щ", "shch": "щ",
    # ── 2-char uppercase ──────────────────────────────────────────────────
    "Ye": "Є", "Zh": "Ж", "Yi": "Ї", "Kh": "Х",
    "Ts": "Ц", "Ch": "Ч", "Sh": "Ш", "Shh": "Щ",
    "Yu": "Ю", "Ya": "Я",
    # ── 2-char lowercase ──────────────────────────────────────────────────
    "ye": "є", "zh": "ж", "yi": "ї", "kh": "х",
    "ts": "ц", "ch": "ч", "sh": "ш", "shh": "щ",
    "yu": "ю", "ya": "я",
    # ── 1-char uppercase (every letter A-Z) ───────────────────────────────
    "A": "А", "B": "Б", "C": "С", "D": "Д", "E": "Е",
    "F": "Ф", "G": "Г", "H": "Г", "I": "І", "J": "Й",
    "K": "К", "L": "Л", "M": "М", "N": "Н", "O": "О",
    "P": "П", "Q": "К", "R": "Р", "S": "С", "T": "Т",
    "U": "У", "V": "В", "W": "В", "X": "КС", "Y": "И",
    "Z": "З",
    # ── 1-char lowercase (every letter a-z) ───────────────────────────────
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
# Matches: <CHOICE HREF="..." TEXT="the text" OPERATOR="..."></A>//
# Skips lines commented out with leading //
CHOICE_LINE_RE = re.compile(r'(?i)^(?!\s*//)(.*)(<CHOICE\b[^>]*\bTEXT="([^"]*)"[^>]*></A>//?.*)$')


def split_voice_line(line: str):
    m = VOICE_LINE_RE.search(line)
    if not m:
        return None
    head_full, speaker, rest = m.groups()
    m2 = re.match(r'(.*?)(\s*(?:<[^>]*>\s*)*)$', rest)
    if not m2:
        return None
    text_part, tail_tags = m2.groups()
    text = text_part.strip()
    head = head_full + rest[: len(rest) - len(rest.lstrip())]
    return speaker, head, text, tail_tags


def split_choice_line(line: str):
    """
    Returns (pre, tag_with_text, text, post) if this is an active CHOICE line.
    'pre' is everything before <CHOICE, 'tag_with_text' is the full tag,
    'text' is the TEXT= value, 'post' is everything after the tag.
    Returns None for commented-out lines or non-CHOICE lines.
    """
    stripped = line.strip()
    # Skip commented-out choice lines
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
    if not stripped or stripped.startswith('<') or stripped.startswith('//'):
        return None
    m = re.match(r'(\s*)(.*?)(\s*(?:<[^>]*>\s*)*)$', line)
    if not m:
        return None
    head_ws, text_part, tail_tags = m.groups()
    text = text_part.strip()
    if not text:
        return None
    return head_ws, text, tail_tags


def build_entries_from_nps(nps_path: Path):
    lines = nps_path.read_text(encoding="utf-8").splitlines()
    entries, entry_id = [], 0
    for lineno, line in enumerate(lines, start=1):
        # 1. Voice line
        vd = split_voice_line(line)
        if vd:
            speaker, _h, text, _t = vd
            if text:
                entries.append({"id": entry_id, "type": "voice", "line_no": lineno,
                                "speaker": speaker, "original": text, "translation": ""})
                entry_id += 1
            continue
        # 2. Choice line
        cd = split_choice_line(line)
        if cd:
            _pre, _tag, text = cd
            if text:
                entries.append({"id": entry_id, "type": "choice", "line_no": lineno,
                                "speaker": "CHOICE", "original": text, "translation": ""})
                entry_id += 1
            continue
        # 3. Narration
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
                    # Replace only the TEXT="..." value inside the tag
                    # Replace only the TEXT="..." value inside the tag
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
# GUI
# ─────────────────────────────────────────────────────────────────────────────
def run_gui(initial_path: Path = None):
    if tk is None:
        print("tkinter unavailable.")
        return

    # IMPORTANT: TkinterDnD.Tk must be used from the start for DnD to work
    if _DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    root.title("NPS Translation Tool")
    root.minsize(900, 600)

    # ── Colour palette ───────────────────────────────────────────────────────
    BG      = "#16161a"
    PANEL   = "#1f1f23"
    SURFACE = "#27272c"
    BORDER  = "#38383f"
    FG      = "#e8e8f0"
    FG_DIM  = "#888899"
    ACCENT  = "#7c6af7"
    ACCENT2 = "#5a9cf8"
    GREEN   = "#4ec97b"
    ORANGE  = "#e8a24f"
    SEL_BG  = "#3a3460"
    BTN_HO  = "#3c3c44"

    root.configure(bg=BG)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background=SURFACE, foreground=FG,
                    fieldbackground=SURFACE, rowheight=24, borderwidth=0,
                    font=("Segoe UI", 9))
    style.map("Treeview",
              background=[("selected", SEL_BG)],
              foreground=[("selected", FG)])
    style.configure("Treeview.Heading", background=PANEL, foreground=FG_DIM,
                    relief="flat", font=("Segoe UI", 9, "bold"))
    style.map("Treeview.Heading", background=[("active", BORDER)])
    for ori in ("Vertical", "Horizontal"):
        style.configure(f"{ori}.TScrollbar", background=PANEL, troughcolor=BG,
                        bordercolor=BG, arrowcolor=FG_DIM, relief="flat")

    # ── Widget helpers ───────────────────────────────────────────────────────
    def mk_btn(parent, text, cmd, color=ACCENT, width=None, tooltip=None):
        kw = dict(text=text, command=cmd, bg=color, fg=FG,
                  activebackground=BTN_HO, activeforeground=FG,
                  relief="flat", bd=0, cursor="hand2",
                  padx=10, pady=5, font=("Segoe UI", 9))
        if width:
            kw["width"] = width
        b = tk.Button(parent, **kw)
        if tooltip:
            _add_tooltip(b, tooltip)
        return b

    def mk_label(parent, text, **kw):
        return tk.Label(parent, text=text,
                        bg=kw.pop("bg", BG), fg=kw.pop("fg", FG_DIM),
                        font=("Segoe UI", 9), **kw)

    def _add_tooltip(widget, tip_text):
        tip_win = [None]
        def show(e):
            tip_win[0] = tk.Toplevel(widget)
            tip_win[0].wm_overrideredirect(True)
            tip_win[0].wm_geometry(f"+{e.x_root+12}+{e.y_root+8}")
            tk.Label(tip_win[0], text=tip_text, bg="#2a2a30", fg=FG,
                     relief="solid", bd=1, font=("Segoe UI", 8), padx=4, pady=2).pack()
        def hide(e):
            if tip_win[0]:
                tip_win[0].destroy()
                tip_win[0] = None
        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

    # ── Speaker aliases (persisted across sessions via JSON file) ──────────
    # Wrapped in a list so nested functions can mutate it without 'nonlocal'
    _aliases_box = [load_aliases()]   # _aliases_box[0] is the live dict

    def get_display_speaker(original: str) -> str:
        if not original:
            return "narrator"
        return _aliases_box[0].get(original, original)

    # ── App state ────────────────────────────────────────────────────────────
    state = {
        "data": None, "entries": [], "json_path": None,
        "nps_path": None, "current_id": None, "modified": False,
    }

    # ════════════════════════════════════════════════════════════════════════
    # LAYOUT
    # ════════════════════════════════════════════════════════════════════════

    toolbar = tk.Frame(root, bg=PANEL, pady=6)
    toolbar.pack(fill=tk.X)
    tb_left = tk.Frame(toolbar, bg=PANEL)
    tb_left.pack(side=tk.LEFT, padx=10)
    tk.Label(toolbar, text=f"v{APP_VERSION}", bg=PANEL, fg="#444455",
             font=("Segoe UI", 8)).pack(side=tk.RIGHT, padx=10)

    tk.Frame(root, bg=BORDER, height=1).pack(fill=tk.X)
    lbl_status = tk.Label(
        root,
        text="No file loaded — drag a .nps or .json here, or use the buttons above",
        bg=PANEL, fg=FG_DIM, font=("Segoe UI", 8), anchor="w", padx=8, pady=3)
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
    search_entry = tk.Entry(search_bar, textvariable=search_var, bg=SURFACE, fg=FG,
                            insertbackground=FG, relief="flat", bd=0,
                            font=("Segoe UI", 9), highlightthickness=1,
                            highlightbackground=BORDER, highlightcolor=ACCENT)
    search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
    mk_btn(search_bar, "◀", lambda: select_match(-1), SURFACE).pack(side=tk.LEFT, padx=2)
    mk_btn(search_bar, "▶", lambda: select_match(1),  SURFACE).pack(side=tk.LEFT, padx=2)
    mk_btn(search_bar, "✕",
           lambda: (search_var.set(""), rebuild_search_matches()), SURFACE
           ).pack(side=tk.LEFT, padx=(2, 8))

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

    tree.tag_configure("odd",      background="#1d1d22")
    tree.tag_configure("even",     background="#23232a")
    tree.tag_configure("done",     foreground=GREEN)
    tree.tag_configure("empty",    foreground=FG_DIM)
    tree.tag_configure("narrator", foreground=ORANGE)
    tree.tag_configure("choice",   foreground="#89ddff")  # cyan for choice lines

    # Speaker colour palette (no green — reserved for "done")
    _SPEAKER_COLOURS = [
        "#c792ea",  # lavender
        "#5a9cf8",  # sky blue
        "#e8a24f",  # amber  (= ORANGE, used for narrator too)
        "#f07178",  # coral red
        "#89ddff",  # cyan
        "#c3e88d",  # lime (light, not bright green)
        "#ff9cac",  # pink
        "#82aaff",  # periwinkle
        "#ffcb6b",  # gold
        "#b2ccd6",  # steel blue-grey
        "#d4a5ff",  # violet
        "#ff869a",  # salmon
        "#80cbc4",  # teal
        "#f78c6c",  # peach-orange
        "#a6accd",  # cool grey
    ]
    # Maps original speaker name → colour string; built in rebuild_tree
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
             font=("Segoe UI", 9, "bold"), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
    mk_btn(speaker_row, "✏ Rename Speaker", lambda: open_alias_editor(), SURFACE,
           tooltip="Set display alias for speakers (only visible in app)"
           ).pack(side=tk.RIGHT, padx=8)

    progress_row = tk.Frame(bot_pane, bg=BG)
    progress_row.pack(fill=tk.X, padx=8, pady=(2, 0))
    lbl_info = tk.Label(progress_row, text="", bg=BG, fg=FG_DIM,
                        font=("Segoe UI", 8), anchor="w")
    lbl_info.pack(side=tk.LEFT)
    lbl_progress = tk.Label(progress_row, text="", bg=BG, fg=GREEN,
                            font=("Segoe UI", 8, "bold"), anchor="e")
    lbl_progress.pack(side=tk.RIGHT)

    # Original text block
    orig_frame = tk.Frame(bot_pane, bg=BG)
    orig_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 2))
    orig_header = tk.Frame(orig_frame, bg=BG)
    orig_header.pack(fill=tk.X)
    mk_label(orig_header, "Original", bg=BG, fg=FG_DIM).pack(side=tk.LEFT)
    # Buttons reference functions defined below — closures capture by reference, so this is fine
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
                            relief="flat", bd=0, font=("Segoe UI", 10),
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
    mk_btn(tr_header, "⧉ Copy",
           lambda: _copy_translation(), SURFACE,
           tooltip="Copy translation to clipboard"
           ).pack(side=tk.RIGHT, padx=(4, 0))
    mk_btn(tr_header, "✕ Clear",
           lambda: translation_text.delete("1.0", tk.END), SURFACE,
           tooltip="Clear translation field"
           ).pack(side=tk.RIGHT, padx=(4, 0))

    translation_text = tk.Text(tr_frame, height=3, wrap="word",
                               bg=SURFACE, fg=FG, insertbackground=FG,
                               relief="flat", bd=0, font=("Segoe UI", 10),
                               padx=8, pady=6, highlightthickness=1,
                               highlightbackground=BORDER, highlightcolor=ACCENT,
                               selectbackground=SEL_BG)
    translation_text.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

    # ════════════════════════════════════════════════════════════════════════
    # LOGIC (all functions that reference widgets defined above)
    # ════════════════════════════════════════════════════════════════════════

    search_state = {"matches": [], "index": -1}

    def set_status(text: str):
        lbl_status.config(text=text)

    def update_progress():
        entries = state["entries"]
        if not entries:
            lbl_progress.config(text="")
            return
        done  = sum(1 for e in entries if (e.get("translation") or "").strip())
        total = len(entries)
        lbl_progress.config(text=f"Translated: {done}/{total}  ({int(done/total*100)}%)")

    def set_entries_from_data(data, json_path, nps_path):
        entries = sorted(data.get("entries", []), key=lambda r: r.get("id", 0))
        state.update(data=data, entries=entries, json_path=json_path,
                     nps_path=nps_path, modified=False, current_id=None)
        rebuild_tree()
        fname = (nps_path.name if nps_path else None) or (json_path.name if json_path else "")
        root.title(f"NPS Translation Tool — {fname}")
        set_status(f"Loaded: {fname}  |  {len(entries)} entries")
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

    # ── Drag & Drop ──────────────────────────────────────────────────────────
    def _on_drop(event):
        raw = event.data.strip()
        # tkinterdnd2 wraps paths that contain spaces in {braces}
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        first = next((a or b for a, b in paths), None)
        if first:
            open_any_file(Path(first))

    if _DND_AVAILABLE:
        # Register the root window and every major widget
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
            tk.Entry(row, textvariable=var, bg=SURFACE, fg=FG,
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
            save_aliases(aliases)   # persist to disk
            rebuild_tree()
            on_tree_select()
            win.destroy()

        btn_row = tk.Frame(win, bg=BG)
        btn_row.pack(pady=10)
        mk_btn(btn_row, "✔ Apply",  apply_aliases, ACCENT).pack(side=tk.LEFT, padx=6)
        mk_btn(btn_row, "✕ Cancel", win.destroy,   SURFACE).pack(side=tk.LEFT, padx=6)

        # Centre window over root after content is laid out
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
        """Copy original to clipboard (no translit)."""
        txt = _get_original_text()
        root.clipboard_clear()
        root.clipboard_append(txt)
        set_status("Copied original to clipboard")

    def _translit_to_translation_field():
        """
        Transliterate Latin → Ukrainian and INSERT result into the
        Translation field (replaces its current content).
        Does NOT touch the clipboard at all.
        """
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

    # ── Tree ─────────────────────────────────────────────────────────────────
    def _ensure_speaker_tag(orig_spk: str):
        """Register a Treeview tag for this speaker if not already done."""
        if not orig_spk:
            return  # narrator uses its own static tag
        tag = f"spk_{orig_spk}"
        if tag not in _speaker_colour_map:
            idx   = len(_speaker_colour_map) % len(_SPEAKER_COLOURS)
            colour = _SPEAKER_COLOURS[idx]
            _speaker_colour_map[tag] = colour
            tree.tag_configure(tag, foreground=colour)

    def rebuild_tree():
        for item in tree.get_children():
            tree.delete(item)
        for i, e in enumerate(state["entries"]):
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

            tree.insert("", tk.END,
                        values=(e.get("id"), disp_spk, e.get("original", ""),
                                e.get("translation", ""), e.get("type", ""),
                                e.get("line_no", "")),
                        tags=(stripe, color_tag))
        on_tree_select()
        update_progress()

    def apply_translation_from_widget():
        entries = state["entries"]
        cur_id  = state.get("current_id")
        if not entries or cur_id is None:
            return
        txt = translation_text.get("1.0", tk.END).rstrip("\n")
        for e in entries:
            if e.get("id") == cur_id:
                if e.get("translation") != txt:
                    e["translation"] = txt
                    state["modified"] = True
                    for iid in tree.get_children():
                        vals = tree.item(iid, "values")
                        if int(vals[0]) == cur_id:
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
        translation_text.focus_set()

    # ── Search ───────────────────────────────────────────────────────────────
    def rebuild_search_matches():
        q = search_var.get().strip().lower()
        search_state["matches"] = []
        search_state["index"]   = -1
        if not q:
            set_status("Search cleared")
            return
        matches = [
            iid for iid in tree.get_children()
            if q in " ".join(str(v) for v in tree.item(iid, "values")[1:4]).lower()
        ]
        search_state["matches"] = matches
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

    # ── Keyboard shortcuts ───────────────────────────────────────────────────
    def on_ctrl_key(event):
        if not (event.state & 0x4):
            return
        w    = root.focus_get()
        code = event.keycode
        if isinstance(w, (tk.Text, tk.Entry)):
            for k, ev in ((67, "<<Copy>>"), (86, "<<Paste>>"), (88, "<<Cut>>")):
                if code == k:
                    w.event_generate(ev)
                    return "break"
            if code == 65:
                if isinstance(w, tk.Text):
                    w.tag_add("sel", "1.0", "end-1c")
                else:
                    w.selection_range(0, tk.END)
                return "break"
        if code == 83:      # Ctrl+S
            quick_save_json()
            return "break"

    root.bind_all("<KeyPress>", on_ctrl_key)

    def on_enter_in_translation(event):
        apply_translation_from_widget()
        children = tree.get_children()
        if not children:
            return "break"
        sel = tree.selection()
        idx = children.index(sel[0]) if sel else -1
        nxt = idx + 1
        if nxt < len(children):
            iid = children[nxt]
            tree.selection_set(iid)
            tree.focus(iid)
            tree.see(iid)
            on_tree_select()
        update_progress()
        return "break"

    translation_text.bind(
        "<FocusOut>", lambda e: (apply_translation_from_widget(), update_progress()))
    translation_text.bind("<Return>",   on_enter_in_translation)
    translation_text.bind("<KP_Enter>", on_enter_in_translation)
    tree.bind("<<TreeviewSelect>>", on_tree_select)

    # ── Toolbar buttons ──────────────────────────────────────────────────────
    mk_btn(tb_left, "📂 Open .nps",  load_nps,            ACCENT
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "📂 Open .json", load_json,           SURFACE
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "💾 Quick Save", quick_save_json,     ACCENT2, tooltip="Ctrl+S"
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "📄 Save .nps",  save_translated_nps, SURFACE
           ).pack(side=tk.LEFT, padx=(0, 4))
    mk_btn(tb_left, "🔢 Counter",    run_counter,         SURFACE
           ).pack(side=tk.LEFT, padx=(0, 4))

    # ── Window close ─────────────────────────────────────────────────────────
    def on_close():
        if state["modified"] and messagebox.askyesno(
                "Exit", "Unsaved changes. Quick Save before exit?"):
            quick_save_json()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    if initial_path is not None:
        open_any_file(initial_path)

    root.mainloop()


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    initial = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    run_gui(initial)