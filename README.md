<div align="center">

# 📄 AI Assistant for Word
### The 100% Free, Open-Source Alternative to Microsoft 365 Copilot

**Talk to Microsoft Word in plain English. Fast, private, and works offline.**

[![GitHub Stars](https://img.shields.io/github/stars/Ayoola-tech2024/ai-assistant-for-word?style=social)](https://github.com/Ayoola-tech2024/ai-assistant-for-word)
[![Platform](https://img.shields.io/badge/Platform-Windows_10%20%2F%2011-0078D6?logo=windows&logoColor=white)](https://github.com/Ayoola-tech2024/ai-assistant-for-word)
[![Word](https://img.shields.io/badge/Microsoft_Word-2013_--_365-2B579A?logo=microsoftword&logoColor=white)](https://github.com/Ayoola-tech2024/ai-assistant-for-word)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Cost](https://img.shields.io/badge/Cost-100%25_Free_Forever-success)](https://github.com/Ayoola-tech2024/ai-assistant-for-word)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org)

<br />

[**⬇️ Download Portable App (.ZIP)**](https://github.com/Ayoola-tech2024/ai-assistant-for-word/releases) • [**📖 How To Use**](HOW_TO_USE.txt) • [**✨ Features**](#-core-features) • [**💡 Why Star This?**](#-why-star-this-project)

</div>

---

## ⚡ Why Pay $360/Year for Copilot?

Microsoft charges **$20 to $30 per month** ($360/year) for Microsoft 365 Copilot. **AI Assistant for Word** gives you the exact same conversational coworker experience inside your desktop Word application for **$0**, powered by high-speed free AI engines and a lightning-fast offline automation core.

| Feature | Microsoft 365 Copilot | AI Assistant for Word |
| :--- | :---: | :---: |
| **Annual Cost** | **$240 – $360 / year** | **$0 (100% Free Forever)** |
| **Account Requirement** | M365 Business / Pro Sub | Any Word 2013, 2016, 2019, 2021, or 365 |
| **Offline Execution** | ❌ None (Cloud only) | ✅ **Instant Local Heuristics Engine (0.001s)** |
| **Safety & Control** | Overwrites without approval | ✅ **Strict Preview & Confirm ("Yes, do it")** |
| **Botanical / Scientific Styler** | ❌ Manual search & format | ✅ **Automatic Latin Binomial Styler in 1 Click** |
| **ChatGPT Markdown Cleaner** | ❌ Pastes raw `#` and `**` | ✅ **Auto-strips hashtags, asterisks & placeholders** |
| **Voice Dictation** | Basic dictation | ✅ **Groq Whisper Large v3 (Fast Speech-to-Text)** |
| **Charts & Visuals** | Limited | ✅ **Native Bar, Line & Pie charts inserted into Word** |
| **Privacy** | Enterprise cloud | ✅ **Local API keys, zero tracking, your PC only** |

---

## ✨ Core Features

### 1. ⚡ Instant Offline Engine (0ms Latency, Zero AI Quota)
Common document edits don't need expensive cloud AI. The assistant includes a built-in deterministic offline engine that executes in milliseconds:
- **🌿 Botanical & Scientific Names**: Finds all Latin biological species and taxa (*Chiroptera*, *Homo sapiens*, *Desmodus rotundus*) and italicizes/underlines them document-wide.
- **🧹 Markdown Artifact Cleaner**: Strips unsightly raw hashtags (`#`, `##`, `###`), bold asterisks (`**`), and bracket placeholders (`[Insert Table Here]`).
- **🔤 Font Resizing**: Instant proportional delta adjustments (`-2pt`, `+2pt`) across selection or entire document.
- **📄 1-Click PDF Export**: Direct COM export revealing the PDF in File Explorer.
- **📊 Document Reading Analytics**: Computes word counts, estimated reading time, speaking pace, and Flesch readability grade level.

### 2. 🤖 Cloud AI Coworker (Powered by Groq & Gemini)
When you need deep creative writing or research, the assistant taps into free high-speed models:
- **✍️ Deep Research & Drafting**: Ask Word to research the web on any topic and synthesize a structured section with cited sources.
- **💬 Margin Review Notes**: Automated peer-review that places constructive review balloons directly in the document margins.
- **🌐 Polyglot Translation**: Translates selections or entire documents into French, Spanish, German, Yoruba, and 50+ languages while preserving structure.
- **📈 Data Visualizations**: Insert styled data tables and generated charts (Bar, Line, Pie).
- **🎙️ Voice Dictation**: Press `Ctrl + Space` or click `🎙️ Voice` to speak your instructions naturally via Groq's high-speed Whisper model.

### 3. 🛡️ Safety-First Preview Loop
Word is **NEVER** touched automatically. Every single command shows you a plain-English preview of what will happen. You decide whether to click **"Yes, do it"** or **"No, cancel"**, backed by a 1-click **Undo** button.

---

## 🚀 Quick Start (Non-Technical Users)

You do **not** need Python or technical experience:

1. **Download**: Grab `AI-Assistant-for-Word-v1.0.zip` from [Releases](https://github.com/Ayoola-tech2024/ai-assistant-for-word/releases).
2. **Extract**: Right-click the `.zip` file and select **Extract All**.
3. **Launch**:
   - Open your document in Microsoft Word.
   - Double-click **`AIAssistantForWord.exe`** (or `Start AI Assistant.cmd`).
4. **First Time Setup (Takes 30 seconds)**:
   - Offline tools (PDF export, Stats, Botanical names, Clean markdown) work **instantly with zero keys**.
   - For AI drafting, click **Keys**, tap **"Get Groq key (free)"**, paste your key, and click **Save**. Done!

---

## 💬 Things You Can Say

Just type naturally in plain English:

- *"Make a comprehensive research about bats and draft a report"*
- *"Write all botanical names well, italise them and underline them"*
- *"Reduce the font and remove those placeholders and sharps you put"*
- *"Export this document to PDF and reveal it in Explorer"*
- *"Turn on track changes and review this contract for risks"*
- *"Insert a 4-column competitor comparison table with pricing"*
- *"Translate the selected text into professional French"*
- *"Add a diagonal watermark CONFIDENTIAL and set header to Internal Use"*
- *"Generate a pie chart showing Sales: Q1 40, Q2 30, Q3 20, Q4 10"*

---

## 🛠️ Developer Setup & Architecture

For developers who want to inspect the source or build custom extensions:

```bash
# 1. Clone the repository
git clone https://github.com/Ayoola-tech2024/ai-assistant-for-word.git
cd ai-assistant-for-word

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the application
python sidebar.py
```

### Architecture Overview

```text
word-ai-service/
├── sidebar.py        # Tkinter floating UI, snap-to-Word logic, voice dictation & events
├── agent.py          # Dual brain: fast_local_plan() offline heuristic matcher + Groq/Gemini cloud planner
├── word_agent.py     # 35+ native Microsoft Word COM dispatch routines via win32com
├── AIAssistantForWord.spec # PyInstaller standalone Windows distribution configuration
└── HOW_TO_USE.txt    # Beginner-friendly guide packaged with the portable executable
```

To build the standalone Windows binary:
```bash
pyinstaller AIAssistantForWord.spec --noconfirm
```

---

## ⭐ Why Star This Project?

If you believe that powerful AI tools shouldn't be locked behind expensive monthly subscriptions, please give this repository a **Star ⭐️**!

- **It helps more students, writers, and researchers discover free tools.**
- **It motivates ongoing active development and new features.**
- **It keeps the project 100% open-source and free forever.**

---

## 📜 License

Distributed under the **MIT License**. See [LICENSE](LICENSE) for details.

Developed with ❤️ by **[Ayoola Damisile Ayooluwa](https://github.com/Ayoola-tech2024)**.
