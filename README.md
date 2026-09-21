# AI Assistant for Word

**Talk to Microsoft Word in plain English. Your AI assistant for documents.**

AI Assistant for Word is a lightweight Windows companion that sits beside Microsoft Word and carries out your instructions the way a skilled assistant would — drafting letters, refining grammar, formatting pages, researching the internet, and handling professional document tasks. No technical skill required: if you can use Word and chat, you can use AI Assistant for Word.

![Platform](https://img.shields.io/badge/platform-Windows-0078D6)
![Word](https://img.shields.io/badge/Word-2013+-2B579A)
![Python](https://img.shields.io/badge/Python-3.10+-3776AB)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Why AI Assistant for Word

| Traditional way | With AI Assistant for Word |
|---|---|
| Memorize ribbon menus and shortcuts | Type what you want in plain English |
| Copy-paste between browser AI and Word | Research, draft, and insert without leaving Word |
| Repeat the same formatting by hand | One command formats pages, headings, tables |
| Hope the AI understood you | Every action shows a preview — nothing runs until you approve |

---

## Get started in 60 seconds

**You need:** Windows 10/11, Microsoft Word 2013 or newer, and one free API key.

1. **Download** the latest `AI Assistant for Word` package from [Releases](../../releases) and unzip it anywhere (e.g. Desktop).
2. **Open your Word document.**
3. **Double-click `Start AI Assistant`.** A small window opens beside Word.
4. **First run only:** tap `Get Groq key` (free, 2 minutes), paste the key, tap `Save & Start`. Done forever.
5. **Type a request**, press Send, check the preview, tap **Yes, do it** — watch Word change instantly.

> Do not open anything inside the engine folder — it contains engine files. `Start AI Assistant` is the only entry point.

### About the free keys

AI Assistant for Word runs on free cloud AI — no subscription, no card.

| Key | Free allowance | Role |
|---|---|---|
| **Groq** (recommended) | Thousands of requests/day | Primary engine |
| **Google Gemini** | ~20 requests/day | Automatic backup |

Add or rotate keys anytime via the **Keys** button in the app. Keys are stored only on your PC (`%APPDATA%\word-cloud-agent\settings.json`) and are never transmitted anywhere except the provider's API.

---

## What you can say

Just write naturally. Name the scope — **this highlighted text**, **the first page**, or **the whole document**.

**Writing**
- *"Write a leave request letter for Friday"*
- *"Draft a two-paragraph welcome note at the end"*

**Editing & refining**
- *"Fix the grammar of this document"*
- *"Make this sound more professional"* (highlight the text first)
- *"Replace YOLD with YNEW"*

**Research**
- *"Find current fuel prices in Nigeria and add them"*
- *"Summarise this document"*

**Formatting & structure**
- *"Underline everything on the first page"*
- *"Make this a Heading 1"* / *"Center this"* / *"Make it 14pt blue"*
- *"Insert a 3x4 table"* / *"Add bullets"* / *"Make it landscape"*
- *"Add page numbers"* / *"Set footer Confidential"*

**Find & navigate**
- *"Find the word refund"* — reports count and pages, jumps to the match
- *"Go to page 3"* / *"Delete page 2"*

**Finish**
- *"How many words is this?"* / *"Save as PDF"* / *"Undo that"*

---

## Capabilities (26 tools)

**Write & Research**
`draft` — compose letters, paragraphs, lists · `research` — internet research inserted into the doc · `ask` — Q&A and summaries (document unchanged)

**Edit**
`replace_text` · `delete_para_containing` · `delete_page` · `refine` (grammar/tone) · `set_full_text` · `insert_at_end` · `insert_at_cursor` · `undo_last`

**Format**
`format_text` (underline/bold/italic, page-scoped) · `font_size` · `font_color` · `align_paras` · `line_spacing` · `make_list` (bullets/numbers)

**Structure**
`apply_heading` (Heading 1–3) · `insert_table` · `add_header` · `add_footer` · `add_page_numbers` · `set_orientation`

**Find & Reports**
`find_text` (counts, pages, jump-to-match) · `select_text` · `go_to_page` · `word_count` · `export_pdf`

---

## Safety by design

- **Nothing runs without approval.** Every request shows a plain-English preview; Word changes only when you tap **Yes, do it**.
- **Your keys stay yours.** Stored locally, used only for provider API calls.
- **Built-in guardrails.** Busy-Write detection (open Word dialogs pause actions with a Retry option), invalid plans are refused with guidance, and `undo_last` reverses mistakes.
- **No document lock-in.** Your files remain ordinary `.docx` files — usable with or without the tool.

---

## Requirements

- Windows 10/11 (64-bit)
- Microsoft Word 2013 or newer (desktop)
- Internet connection (cloud AI + optional research)
- One free API key (Groq recommended)

> macOS, web Word, and mobile Word are not supported — the tool drives the Windows desktop Word engine.

---

## Run from source (developers)

```bash
pip install -r requirements.txt
python sidebar.py
```

To build the double-click Windows package:

```bash
pip install pyinstaller
pyinstaller -y --onedir --windowed --name AIAssistantForWord --hidden-import win32com.client sidebar.py
```

Project layout:

```text
word-ai-service/
├── sidebar.py      # User interface (chat window, setup, approvals)
├── agent.py        # AI planning: Groq primary, Gemini fallback, rotation
├── word_agent.py   # Word automation: 26 tools over Word COM
├── requirements.txt
└── Start AI Assistant.cmd  # Launcher (packaged alongside the build)
```

---

## Roadmap

- Multi-document awareness and batch jobs
- Custom reusable prompts ("my report style")
- Track-changes-friendly edits
- Additional free providers for even higher limits

Feature requests and bug reports are welcome via [Issues](../../issues).

---

## Contributing

1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/your-idea`).
3. Keep Word 2013 compatibility (COM only, no Office.js).
4. Never commit API keys — `.env` is git-ignored by design.
5. Open a pull request with a clear description and test notes.

---

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Ayoola Damisile Ayooluwa.

Built for everyone who writes in Word and deserves a coworker, not a maze of menus.
