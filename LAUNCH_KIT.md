# 🚀 Launch Marketing & Promotion Kit
## AI Assistant for Word — Open-Source Alternative to Microsoft 365 Copilot

This kit contains ready-to-publish posts, article drafts, and outreach strategies tailored for LinkedIn, Reddit, Dev.to, Twitter/X, and Product Hunt to help you gain GitHub stars, user feedback, and professional recognition.

---

## 1. LinkedIn Launch Post

**Target Audience:** Professionals, researchers, students, developers, tech recruiters.  
**Best Time to Post:** Tuesday – Thursday, between 8:00 AM – 11:00 AM local time.

### Post Copy:
> 💡 Why pay $360/year for Microsoft 365 Copilot when you can build a faster, 100% free, and private alternative?
>
> Over the past few weeks, I’ve been working on **AI Assistant for Word** — a lightweight, open-source Windows companion that sits right beside Microsoft Word to give you conversational AI powers directly inside your documents.
>
> Here’s why I built it:
> ❌ Copilot costs $20–$30/user/month (too expensive for students and small teams).  
> ❌ Web AI tools force you to constantly switch tabs and copy-paste ugly `#` hashtags and markdown artifacts into Word.  
> ❌ Cloud-only tools fail when offline and overwrite documents unpredictably.
>
> What AI Assistant for Word does differently:
> ⚡ **Hybrid 0ms Offline Engine**: Instant local execution for font resizing, markdown cleaning, PDF export, and readability stats with ZERO internet or API keys.  
> 🌿 **Botanical & Scientific Nomenclature Styler**: Automatically finds Latin binomial species (*Chiroptera*, *Desmodus rotundus*) document-wide and formats them with italics and underlines in 1 click.  
> 🧹 **ChatGPT Cleaner**: Automatically strips raw `#` sharps, asterisks, and bracket placeholders before they hit your document.  
> 🛡️ **Safety-First**: Word is NEVER touched without your explicit approval via the interactive preview box.  
> 🎙️ **Voice Dictation**: Built-in Groq Whisper speech-to-text.  
> 💸 **100% Free Forever**: Runs with free Groq & Google Gemini keys (thousands of free queries a day).
>
> It’s completely open-source, and you don’t even need Python installed — you can download the standalone Windows app, unzip, and start using it immediately.
>
> 🔗 Check out the project and give it a star ⭐️ on GitHub:  
> https://github.com/Ayoola-tech2024/ai-assistant-for-word
>
> I’d love your feedback and thoughts! What feature would you like to see next?
>
> #OpenSource #Python #ArtificialIntelligence #Productivity #MicrosoftWord #Innovation #SoftwareEngineering

---

## 2. Reddit Posts

### A. For `r/Python` & `r/opensource`
**Title:** I built a free, open-source alternative to Microsoft Copilot for Word using Python & Word COM (with 0ms offline heuristics)

**Body:**
> Hey r/Python!
>
> Like many of you, I find Microsoft 365 Copilot's $30/month pricing prohibitive, especially for students, academics, and indie creators.
>
> I built **AI Assistant for Word**, an open-source Windows companion app that controls native desktop Word via `win32com` with a dual-layer architecture:
>
> **1. Fast Offline Heuristic Core (0ms latency, zero quota):**  
> Simple operations (font delta adjustments, 1-click PDF exports, table formatting, and watermark generation) are intercepted before touching any LLM. It also features an automated botanical/scientific name parser that recognizes Latin binomial nomenclature (*Genus species*) and higher biological taxa suffixes, styling them with native Word formatting instantly.
>
> **2. Free LLM Layer (Groq + Gemini rotation):**  
> For creative drafting, internet research, and document reviews, it connects to Groq (`whisper-large-v3`, `gpt-oss-20b`, `qwen`) with an automatic failover to Google Gemini. It includes ready-first key rotation and cooldown management.
>
> **3. Safe-by-Design Preview:**  
> The agent validates all tool plans and presents an interactive diff preview to the user. No change is committed to the active document until the user taps "Yes, do it".
>
> It's packaged as a portable, standalone executable (zero Python required for end users).
>
> Code & releases: https://github.com/Ayoola-tech2024/ai-assistant-for-word  
>
> Would love your feedback on the architecture and any contributions!

---

### B. For `r/SideProject` & `r/productivity`
**Title:** Tired of paying $360/year for Copilot? I made a 100% free, private AI Assistant for Word that works offline

**Body:**
> If you write reports, academic papers, or documents in Microsoft Word, you probably know the pain of copying and pasting from ChatGPT, only to spend 15 minutes fixing broken `#` hashtags, bold asterisks, and misaligned fonts.
>
> I created **AI Assistant for Word** to solve this:
> - Sits right beside Word as a floating assistant.
> - Cleans messy ChatGPT markdown automatically.
> - Runs common formatting (font resizing, PDF export, document analytics) 100% offline with zero internet.
> - Automatically formats plant and animal scientific names with italics and underline.
> - 100% free with no subscriptions.
>
> It's available as a free download on GitHub:  
> https://github.com/Ayoola-tech2024/ai-assistant-for-word
>
> Let me know what you think!

---

## 3. Dev.to / Medium Article Draft

**Title:** How I Built a Free, Open-Source Alternative to Microsoft 365 Copilot for Word  
**Tags:** `#python` `#ai` `#opensource` `#productivity`

### Outline & Content:

#### Introduction: The Problem with Office AI
Microsoft Copilot costs $360 per user per year. For millions of students, researchers, and developers, this makes AI assistance in office productivity inaccessible. Furthermore, using browser-based LLMs results in tedious tab-switching and messy formatting artifacts.

#### The Goal
Build a lightweight, safe, and completely free desktop assistant for Microsoft Word that:
1. Costs $0 to run.
2. Doesn't overwrite documents without user permission.
3. Operates offline for standard formatting tasks.
4. Requires zero technical setup for non-developers.

#### Key Engineering Highlights:
- **Direct Word COM Dispatch**: Using `win32com.client`, the assistant talks directly to Microsoft Word's COM interface, enabling rich paragraph styling, table creation, and page manipulation.
- **The Dual-Brain Architecture**: Rather than sending every simple prompt to a cloud LLM, deterministic commands (font adjustments, markdown cleaning, botanical Latin styling) are handled by a local heuristic matcher (`fast_local_plan()`) in under 30 milliseconds.
- **Smart Key Rotation & Cooldowns**: Built-in round-robin rotation for free Groq and Gemini API keys with automatic backoff on 429 rate limits.
- **Voice Dictation**: Uses native Windows audio capture (`winmm.dll`) paired with Groq's high-speed Whisper model for real-time dictation.

#### Try It Out
Check out the repository and release packages on GitHub:  
https://github.com/Ayoola-tech2024/ai-assistant-for-word

---

## 4. Twitter / X Launch Thread

**Tweet 1 (Hook):**  
> Microsoft Copilot costs $360/year.
> 
> So I spent the last few weeks building a 100% FREE, open-source alternative for Microsoft Word that runs locally, works offline, and respects your privacy.
> 
> Introducing AI Assistant for Word 🧵👇

**Tweet 2:**  
> The #1 issue with AI in Word: Hallucinations and unwanted overwrites.
> 
> AI Assistant for Word uses a strict "Check before it runs" preview loop. Nothing in your document ever changes until you click "Yes, do it", backed by 1-click Undo.

**Tweet 3:**  
> It also features an instant 0ms offline engine:
> 🌿 Auto-styles botanical/scientific names in 1 click  
> 🧹 Auto-cleans ugly ChatGPT `#` hashtags & asterisks  
> 🔤 Instant font resizing & 1-click PDF export  
> Zero internet or API keys required for offline tasks.

**Tweet 4:**  
> Powered by free Groq & Google Gemini keys for deep web research, polyglot translation, and voice dictation.
> 
> Best of all: Non-technical users don't need Python installed. Just download the ZIP and run!
> 
> ⭐️ Star the project on GitHub:  
> https://github.com/Ayoola-tech2024/ai-assistant-for-word
