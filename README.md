# NPS Translation Tool

A desktop GUI tool for translating `.nps` script files used in NPS-based visual novel engines (e.g. *The Song of Saya* / *Saya no Uta*).

Supports voice lines, narration, and choice branches. Saves progress to a sidecar JSON file so you can pick up where you left off.

---

## Features

- **Open `.nps` and `.json` files** via buttons or drag & drop
- **Three line types supported:**
  - `<voice name="...">` — character voice lines
  - Plain narration lines
  - `<CHOICE ... TEXT="...">` — choice/branch lines (translates only the `TEXT` value)
- **Per-row colour coding** — each speaker gets a unique colour; narrator is amber; choice lines are cyan; translated lines turn green
- **Speaker aliases** — rename any speaker for display convenience; aliases are never written to the source file and persist across sessions
- **Transliteration** — one-click Latin → Ukrainian transliteration inserted directly into the translation field
- **Search** — live search across speaker, original, and translation columns with Prev/Next navigation
- **Progress tracking** — live counter showing how many lines are translated out of total
- **Quick Save** — saves translation progress to a `.json` sidecar file (Ctrl+S)
- **Save .nps** — writes a translated copy of the original file as `translated_<filename>.nps`
- **Counter** — batch word/line count report across multiple `.nps` files
- **Keyboard navigation** — press Enter in the translation field to save and jump to the next line

---

## Requirements

### Running from source

- Python 3.10+
- `tkinterdnd2` *(optional — required for drag & drop)*

```bash
pip install tkinterdnd2
```

### Running as `.exe`

No Python required. Download the prebuilt `.exe` from [Releases](../../releases).

---

## Building the `.exe`

```bash
pip install pyinstaller tkinterdnd2
```

Create a `build.bat` in the same folder as the script:

```bat
@echo off
set SCRIPT_NAME=NPSTranslationTool.py
set EXE_NAME=NPSTranslationTool
py -3 -m PyInstaller --noconfirm --onefile --windowed --name "%EXE_NAME%" "%SCRIPT_NAME%"
echo Done. Executable is in the "dist" folder.
pause
```

Run `build.bat`. The `.exe` will appear in the `dist\` folder.

> Use `--console` instead of `--windowed` during development if you need to see error output.

---

## Usage

### Opening a file

- Click **📂 Open .nps** to parse a new script file
- Click **📂 Open .json** to reopen a previously saved session
- Or drag and drop a `.nps` or `.json` file onto the window *(requires `tkinterdnd2`)*

When a `.nps` file is opened, a sidecar `.json` is automatically created next to it. If one already exists, saved translations are restored automatically.

### Translating

1. Click a row in the table to select it
2. The original text appears in the **Original** panel
3. Type your translation in the **Translation** panel
4. Press **Enter** to save and move to the next line, or click another row

### Buttons in the editor panel

| Button | Action |
|---|---|
| ⧉ Copy | Copy the original text to clipboard |
| 🔤 Translit → Translation | Transliterate Latin characters to Ukrainian and insert into the translation field |
| ⧉ Copy *(translation)* | Copy the translation to clipboard |
| ✕ Clear | Clear the translation field |

### Speaker aliases

Click **✏ Rename Speaker** to open the alias editor. Set a display name for any speaker — it will show in the table and speaker label instead of the original name.

- Aliases are **never written to the `.nps` or `.json` file**
- They are saved to `nps_speaker_aliases.json` next to the `.exe` / script and restored on next launch
- Aliases apply across all files — if two files share a speaker name, the alias carries over

### Saving

- **💾 Quick Save** (or Ctrl+S) — saves translation progress to the sidecar `.json`
- **📄 Save .nps** — writes the translated script as `translated_<original_filename>.nps` in the same folder as the source file. Untranslated lines fall back to the original text.

---

## File format

The sidecar `.json` looks like this:

```json
{
  "source_file": "C:/path/to/script.nps",
  "entries": [
    {
      "id": 0,
      "type": "voice",
      "line_no": 12,
      "speaker": "Fuminori",
      "original": "It's nothing.",
      "translation": "Нічого."
    },
    {
      "id": 1,
      "type": "choice",
      "line_no": 34,
      "speaker": "CHOICE",
      "original": "...want it all back.",
      "translation": "...хочу повернути все."
    }
  ]
}
```

`type` is one of `voice`, `narration`, or `choice`.

---

## Transliteration table

The **🔤 Translit** button converts Latin characters to Ukrainian using this mapping:

| Latin | Ukrainian | Latin | Ukrainian |
|---|---|---|---|
| shch | щ | sh | ш |
| ch | ч | zh | ж |
| kh | х | ts | ц |
| ya | я | yu | ю |
| ye | є | yi | ї |
| a | а | b | б |
| v | в | g | г |
| h | г | d | д |
| e | е | z | з |
| y | и | i | і |
| j | й | k | к |
| l | л | m | м |
| n | н | o | о |
| p | п | r | р |
| s | с | t | т |
| u | у | f | ф |
| w | в | x | кс |
| q | к | c | с |

Multi-character sequences are matched before single characters (longest match first).

---

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+S | Quick Save |
| Enter *(in translation field)* | Save current line and move to next |
| Ctrl+C / V / X / A | Standard clipboard shortcuts in text fields |

---

## License

MIT