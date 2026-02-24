# NPS Translation Tool

**NPS Translation Tool** is a utility for extracting dialogue from `.nps` visual novel script files, translating it, and writing translations back into the game.

---

# UA Українська версія

## 🚀 Запуск

Просто запустіть:

```
NpsTranslator.exe
```

---

## 📂 Робочий процес

### 1️⃣ Open .nps

- Відкриває `.nps` файл
- Парсить усі репліки (`voice` + `narration`)
- Автоматично створює або оновлює `.json` файл поруч
- Якщо JSON вже існує — підтягує раніше збережені переклади

---

### 2️⃣ Quick Save

- Зберігає поточні переклади у JSON
- Файл має ту ж назву, що й `.nps`, але з розширенням `.json`
- Рекомендується використовувати регулярно

---

### 3️⃣ Save .nps

Створює новий файл:

```
translated_<оригінальна_назва>.nps
```

- Усі технічні теги (`<voice>`, службові команди, і т.д.) зберігаються
- Змінюється лише текст реплік

---

### 4️⃣ Counter

Створює звіт з:

- загальною кількістю рядків
- загальною кількістю слів
- переліком усіх реплік для кожного файлу

---

## 🧾 Таблиця реплік

| Колонка | Опис |
|----------|------|
| № | внутрішній ID |
| Speaker | ім’я персонажа (`narrator`, якщо відсутнє) |
| Original | оригінальний текст |
| Translation | переклад |
| type | `voice` або `narration` |
| line | номер рядка у файлі |

При виборі рядка:

- оригінал показується у блоці **Original (selected line)**
- переклад редагується у **Translation (selected line)**

---

## ✍️ Редагування перекладу

1. Оберіть рядок у таблиці
2. Введіть переклад
3. Натисніть **Enter** — переклад збережеться і перейде до наступного рядка
4. При перемиканні мишкою переклад попереднього рядка зберігається автоматично

Щоб зберегти переклади між сесіями — використовуйте **Quick Save**.

---

## 🔎 Пошук

Пошук без урахування регістру в:

- Speaker
- Original
- Translation

Кнопки:

- **Prev**
- **Next**
- **Clear**

Enter у полі пошуку = Next.

---

## ⌨️ Гарячі клавіші

- Ctrl+C — копіювати
- Ctrl+V — вставити
- Ctrl+X — вирізати
- Ctrl+A — виділити все

---

## 📄 Формат JSON

```json
{
  "source_file": "path_to_original.nps",
  "entries": [
    {
      "id": 1,
      "type": "voice",
      "line_no": 123,
      "speaker": "Character",
      "original": "Original text",
      "translation": "Translated text"
    }
  ]
}
```

---

# EN English Version

## 🚀 Launch

Simply run:


NpsTranslator.exe


---

## 📂 Workflow

### 1️⃣ Open .nps

- Opens a `.nps` file
- Parses all dialogue lines (`voice` + `narration`)
- Automatically creates or updates a `.json` file next to it
- If a JSON file already exists — previously saved translations are loaded automatically

---

### 2️⃣ Quick Save

- Saves current translations to JSON
- File has the same name as the `.nps`, but with a `.json` extension
- Recommended to use regularly to avoid losing progress

---

### 3️⃣ Save .nps

Creates a new file:


translated_<original_name>.nps


- All technical tags (`<voice>`, script commands, etc.) are preserved
- Only the dialogue text is replaced

---

### 4️⃣ Counter

Generates a report containing:

- total number of translatable lines
- total word count
- full list of lines for each selected file

---

## 🧾 Lines Table

| Column | Description |
|--------|------------|
| № | internal line ID |
| Speaker | character name (`narrator` if empty) |
| Original | original text |
| Translation | translated text |
| type | `voice` or `narration` |
| line | line number inside the file |

When selecting a row:

- original text is displayed in **Original (selected line)**
- translation can be edited in **Translation (selected line)**

---

## ✍️ Editing Translation

1. Select a row in the table
2. Enter your translation
3. Press **Enter** — the translation is saved and selection moves to the next row
4. Switching rows with the mouse saves the previous translation automatically

Translations are stored in memory while the program is running.  
Use **Quick Save** to persist them between sessions.

---

## 🔎 Search

Case-insensitive search in:

- Speaker
- Original
- Translation

Buttons:

- **Prev**
- **Next**
- **Clear**

Enter in the search field works as **Next**.

---

## ⌨️ Hotkeys

- Ctrl+C — copy
- Ctrl+V — paste
- Ctrl+X — cut
- Ctrl+A — select all

Works in all text fields and in the table.

---

## 📄 JSON Format

```json
{
  "source_file": "path_to_original.nps",
  "entries": [
    {
      "id": 1,
      "type": "voice",
      "line_no": 123,
      "speaker": "Character",
      "original": "Original text",
      "translation": "Translated text"
    }
  ]
}

