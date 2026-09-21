"""Word Cloud Agent — coworker loop. PLANS ONLY, never executes.

plan_action(user_text) -> dict:
  {"ok": True, "tool": ..., "args": {...}, "explain": ..., "doc": {...}}
  optionally {"answer_only": True, "answer": ...} for ask (no doc change).
Execution happens only in sidebar.py after your 'Yes, do it' tap,
via execute_action(plan).
"""
import json
import os
import pathlib
import time

import requests
from google import genai
from google.genai import types

import word_agent

BASE = pathlib.Path(__file__).resolve().parent
SETTINGS_FILE = pathlib.Path(os.environ.get("APPDATA", str(BASE))) \
    / "word-cloud-agent" / "settings.json"


def _split_keys(raw):
    """One key per line or comma. Accepts a string or a saved list.
    Returns de-duplicated list."""
    if isinstance(raw, list):
        parts = []
        for item in raw:
            parts.extend(str(item).replace(",", "\n").splitlines())
    else:
        parts = str(raw or "").replace(",", "\n").splitlines()
    out = []
    for part in parts:
        part = part.strip().strip('"').strip("'").strip("[]")
        if len(part) >= 20 and part not in out:
            out.append(part)
    return out


def load_config():
    cfg = {"GEMINI_API_KEY": "", "GEMINI_KEYS": [], "GROQ_KEYS": [],
           "GEMINI_MODEL": "gemini-3.6-flash"}
    # 1) one-click settings file (written by the app's setup screen)
    try:
        if SETTINGS_FILE.exists():
            saved = json.loads(SETTINGS_FILE.read_text())
            if saved.get("GEMINI_MODEL"):
                cfg["GEMINI_MODEL"] = str(saved["GEMINI_MODEL"]).strip()
            keys = _split_keys(saved.get("GEMINI_KEYS") or
                               saved.get("GEMINI_API_KEY"))
            if keys:
                cfg["GEMINI_KEYS"] = keys
            gkeys = _split_keys(saved.get("GROQ_KEYS") or
                                saved.get("GROQ_API_KEY"))
            if gkeys:
                cfg["GROQ_KEYS"] = gkeys
    except Exception:
        pass
    # 2) developer .env fallback (only fills blanks)
    env = BASE / ".env"
    if env.exists():
        for line in env.read_text().splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k == "GEMINI_API_KEY" and v and not cfg["GEMINI_KEYS"]:
                keys = _split_keys(v)
                if keys:
                    cfg["GEMINI_KEYS"] = keys
            elif k in ("GROQ_API_KEY", "GROQ_KEYS") and v \
                    and not cfg["GROQ_KEYS"]:
                keys = _split_keys(v)
                if keys:
                    cfg["GROQ_KEYS"] = keys
            elif k in cfg and not cfg[k] and v:
                cfg[k] = v
    if cfg["GEMINI_KEYS"] and not cfg["GEMINI_API_KEY"]:
        cfg["GEMINI_API_KEY"] = cfg["GEMINI_KEYS"][0]
    return cfg


def save_api_key(key, kind="gemini"):
    """Persist key(s) — one per line accepted — so pasting is once.
    kind is 'gemini' or 'groq'. Returns total saved of that kind."""
    keys = _split_keys(key)
    if not keys:
        raise ValueError("No valid key found in what was pasted.")
    slot = "GROQ_KEYS" if kind == "groq" else "GEMINI_KEYS"
    legacy = "GROQ_API_KEY" if kind == "groq" else "GEMINI_API_KEY"
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    data = {}
    try:
        if SETTINGS_FILE.exists():
            data = json.loads(SETTINGS_FILE.read_text())
    except Exception:
        data = {}
    old = _split_keys(data.get(slot) or data.get(legacy))
    for k in keys:
        if k not in old:
            old.append(k)
    data[slot] = old
    data.pop(legacy, None)
    SETTINGS_FILE.write_text(json.dumps(data))
    return len(old)


TOOLS = ["add_footer", "replace_text", "delete_para_containing",
         "insert_at_end", "draft", "refine", "ask", "research",
         "format_text", "find_text", "select_text", "go_to_page",
         "delete_page", "apply_heading", "insert_table", "make_list",
         "align_paras", "font_size", "font_color", "line_spacing",
         "set_orientation", "add_header", "add_page_numbers", "word_count",
         "export_pdf", "undo_last"]

REQUIRED_ARGS = {
    "add_footer": ("text",),
    "replace_text": ("old", "new"),
    "delete_para_containing": ("text",),
    "insert_at_end": ("text",),
    "draft": ("instruction",),
    "refine": ("instruction",),
    "ask": ("question",),
    "research": ("query",),
    "format_text": ("style",),
    "find_text": ("text",),
    "select_text": ("text",),
    "go_to_page": ("page",),
    "delete_page": ("page",),
    "apply_heading": ("text",),
    "insert_table": ("rows", "cols"),
    "make_list": (),
    "align_paras": (),
    "font_size": ("size",),
    "font_color": ("color",),
    "line_spacing": (),
    "set_orientation": (),
    "add_header": ("text",),
    "add_page_numbers": (),
    "word_count": (),
    "export_pdf": (),
    "undo_last": (),
}
OPTIONAL_ARGS = {
    "draft": ("place",),
    "research": ("instruction",),
    "format_text": ("page",),
    "apply_heading": ("level",),
    "insert_table": ("place",),
    "make_list": ("kind",),
    "align_paras": ("how",),
    "line_spacing": ("spacing",),
    "set_orientation": ("orientation",),
    "undo_last": ("steps",),
}

RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "tool": {"type": "STRING", "enum": TOOLS},
        "args": {
            "type": "OBJECT",
            "properties": {
                "text": {"type": "STRING"},
                "old": {"type": "STRING"},
                "new": {"type": "STRING"},
                "instruction": {"type": "STRING"},
                "question": {"type": "STRING"},
                "query": {"type": "STRING"},
                "place": {"type": "STRING",
                          "enum": ["cursor", "end"]},
                "style": {"type": "STRING",
                          "enum": ["underline", "bold", "italic"]},
                "page": {"type": "NUMBER"},
                "level": {"type": "NUMBER"},
                "rows": {"type": "NUMBER"},
                "cols": {"type": "NUMBER"},
                "size": {"type": "NUMBER"},
                "steps": {"type": "NUMBER"},
                "kind": {"type": "STRING",
                         "enum": ["bullets", "numbers"]},
                "how": {"type": "STRING",
                        "enum": ["left", "center", "right", "justify"]},
                "color": {"type": "STRING"},
                "spacing": {"type": "STRING",
                            "enum": ["single", "1.5", "double"]},
                "orientation": {"type": "STRING",
                                "enum": ["portrait", "landscape"]},
            },
        },
        "explain": {"type": "STRING"},
    },
    "required": ["tool", "args"],
}

SYSTEM = (
    "You are a coworker inside Microsoft Word. Pick ONE tool for the user's "
    "request and extract its arguments. Output JSON only, matching the schema.\n"
    "Small edits:\n"
    "- add_footer {text}: set the footer on every section.\n"
    "- replace_text {old, new}: swap every occurrence of old with new.\n"
    "- delete_para_containing {text}: delete paragraphs containing text.\n"
    "- insert_at_end {text}: append exact user-given text at the end.\n"
    "- format_text {style, page?}: underline, bold, or italic. Use page when "
    "the user names one ('first page' -> 1), highlight when they say "
    "'this/selected', else whole document.\n"
    "Find & move:\n"
    "- find_text {text}: locate text — reports count + pages and jumps the "
    "cursor there. Use for 'find/where is'. Never use select_text just to "
    "report.\n"
    "- select_text {text}: highlight a match in the Word window.\n"
    "- go_to_page {page}: jump the cursor ('go to page 2', 'first page' -> 1).\n"
    "Structure:\n"
    "- delete_page {page}: delete a whole page by number.\n"
    "- apply_heading {text, level?=1}: make the paragraph containing text a "
    "Heading (level 1-3).\n"
    "- insert_table {rows, cols, place?=end}: empty table.\n"
    "- make_list {kind?=bullets}: bullets|numbers on highlight or document.\n"
    "- align_paras {how?=left}: left|center|right|justify.\n"
    "- font_size {size}: number like 14, on highlight or document.\n"
    "- font_color {color}: red, blue, green, black, orange, purple, grey.\n"
    "- line_spacing {spacing?=single}: single|1.5|double.\n"
    "- set_orientation {orientation?=portrait}: portrait|landscape.\n"
    "- add_header {text}: header on every section.\n"
    "- add_page_numbers {}: 'Page X' in the footer.\n"
    "Reports & safety:\n"
    "- word_count {}: counts (nothing changes).\n"
    "- export_pdf {}: save a PDF next to the document.\n"
    "- undo_last {steps?=1}: undo recent change(s).\n"
    "Coworker jobs:\n"
    "- draft {instruction, place=cursor|end}: WRITE new content from scratch "
    "(full letters, paragraphs, lists). place=cursor if the user said "
    "'here'/'at cursor', else end.\n"
    "- refine {instruction}: IMPROVE the selected text (or whole document if "
    "scope=document) — fix grammar, tone, clarity per instruction.\n"
    "- ask {question}: answer about the document (summaries, explanations, "
    "questions). NEVER use ask for writing or changing the document.\n"
    "- research {query, instruction}: find CURRENT info from the internet and "
    "add it to the document. query = short search terms, instruction = what "
    "to write with the findings.\n"
    "Rules: writing new content -> draft. Improving existing text -> refine. "
    "Questions/summaries -> ask. Anything needing the internet -> research. "
    "If unsure between draft and insert_at_end: user-supplied exact text -> "
    "insert_at_end; content the AI must compose -> draft. "
    "Explain mismatches in 'explain'. Never invent tools. "
    "Example: {\"tool\": \"add_footer\", \"args\": {\"text\": \"Confidential\"}} "
    "— always fill every required args field with the user's exact words."
)


def _validate(plan):
    tool = plan.get("tool")
    if tool not in REQUIRED_ARGS:
        return {"ok": False, "error": "Unknown tool: %r" % (tool,)}
    args = plan.get("args") or plan.get("arguments") or {}
    missing = [a for a in REQUIRED_ARGS[tool]
               if not str(args.get(a, "")).strip()]
    if missing:
        return {"ok": False,
                "error": "Missing info for %s: %s"
                         % (tool, ", ".join(missing)),
                "tool": tool, "args": args}
    clean = {a: str(args[a]).strip() for a in REQUIRED_ARGS[tool]}
    for a in OPTIONAL_ARGS.get(tool, ()):
        if str(args.get(a, "")).strip():
            clean[a] = str(args[a]).strip()
    if tool == "draft" and clean.get("place") not in ("cursor", "end"):
        clean["place"] = "end"
    if tool == "format_text":
        s = clean.get("style", "").lower()
        if s.startswith("underlin"):
            clean["style"] = "underline"
        elif "bold" in s:
            clean["style"] = "bold"
        elif "ital" in s:
            clean["style"] = "italic"
    return {"ok": True, "tool": tool, "args": clean,
            "explain": str(plan.get("explain", ""))[:300]}


# -- Groq free engine (primary: thousands of free requests/day) --
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODELS = [
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant",
    "openai/gpt-oss-120b",
    "openai/gpt-oss-20b",
]


def _groq_call(cfg, system_text, user_text, want_json, note_fn=None):
    """Returns (text, tag): tag None on success, 'limited' if all keys
    rate-limited, or 'failed:<msg>' for other errors."""
    keys = cfg.get("GROQ_KEYS", [])
    if not keys:
        return None, "no-key"
    limited = 0
    last = "unknown"
    for i, key in enumerate(keys):
        for model in GROQ_MODELS:
            body = {
                "model": model,
                "messages": [
                    {"role": "system", "content": system_text},
                    {"role": "user", "content": user_text},
                ],
                "temperature": 0.2,
            }
            if want_json:
                body["response_format"] = {"type": "json_object"}
            try:
                r = requests.post(
                    GROQ_URL,
                    headers={"Authorization": "Bearer " + key,
                             "Content-Type": "application/json"},
                    json=body, timeout=60)
            except Exception as e:
                last = "network: " + str(e)[:150]
                break  # network issue — same for all keys, go to fallback
            if r.status_code == 429:
                limited += 1
                if note_fn and len(keys) > 1:
                    note_fn("Groq key %d of %d is busy — switching..."
                            % (i + 1, len(keys)))
                break  # next key
            if r.status_code == 401:
                last = "Groq key %d rejected (check the key)" % (i + 1)
                break  # bad key — next key
            if r.status_code != 200:
                last = "Groq %s: %s" % (r.status_code,
                                        r.text[:150].replace("\n", " "))
                continue  # try next model
            try:
                return r.json()["choices"][0]["message"]["content"], None
            except Exception as e:
                last = "bad reply: " + str(e)[:150]
    if limited >= len(keys) > 0:
        return None, "limited"
    return None, "failed:" + last


def _clients(cfg):
    """One SDK client per key (fresh each operation; avoids stale pools)."""
    return [genai.Client(api_key=k) for k in cfg.get("GEMINI_KEYS", [])]


def _is_retryable(msg):
    m = msg.lower()
    return ("503" in msg or "unavailable" in m
            or "overloaded" in m or "timeout" in m)


def _retry_delay(msg, cap=90):
    """Google 429s carry a retryDelay like '38s'. Honor it (capped)."""
    import re
    m = re.search(r"retry in (\d+)", msg, re.IGNORECASE)
    if m:
        return min(int(m.group(1)) + 2, cap)
    return 0


def _call_json(cfg, prompt, note_fn=None):
    """Strict-JSON call across keys: rotate on 429, quick 503 retries,
    one patient wait only when EVERY key is limited. note_fn(str) gets
    human progress updates (sidebar shows them instead of hanging).
    Groq free engine is tried first; Gemini rotation is the fallback.
    Returns (obj, error)."""
    if cfg.get("GROQ_KEYS"):
        if note_fn:
            note_fn("Asking the free engine...")
        text, tag = _groq_call(cfg, SYSTEM, prompt, True, note_fn)
        if tag is None:
            try:
                return json.loads(text), None
            except Exception:
                pass  # fall through to Gemini
        elif tag != "limited":
            pass  # keys bad/network — fall through to Gemini
    clients = _clients(cfg)
    if not clients:
        return None, ("No free key saved yet. Open Setup: add a Groq key "
                      "(recommended) or a Gemini key, paste, Save.")
    last = "unknown"
    waits = {}  # key index -> unix time when usable again
    tried_wait = False
    for _round in range(3):
        for i, client in enumerate(clients):
            if waits.get(i, 0) > time.time():
                continue
            for attempt in range(2):
                try:
                    resp = client.models.generate_content(
                        model=cfg["GEMINI_MODEL"] or "gemini-3.6-flash",
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=RESPONSE_SCHEMA,
                        ),
                    )
                    return json.loads(resp.text), None
                except Exception as e:
                    last = str(e)[:300]
                    if "429" in str(e):
                        waits[i] = time.time() + (_retry_delay(str(e)) or 45)
                        if note_fn and len(clients) > 1:
                            note_fn("Key %d of %d hit Google's free limit — "
                                    "switching key..." % (i + 1, len(clients)))
                        break  # next key, don't waste retries on 429
                    if attempt < 1 and _is_retryable(str(e)):
                        time.sleep(2)
                        continue
                    break
        # all keys cooling? wait once for the earliest, then try again
        if waits and all(w > time.time() for w in waits.values()) \
                and not tried_wait:
            wait = min(max(int(w - time.time()) + 1, 1) for w in waits.values())
            wait = min(wait, 90)
            if note_fn:
                note_fn("Google's free limit is busy — waiting %ds, "
                        "then retrying..." % wait)
            time.sleep(wait)
            tried_wait = True
            continue
        break
    if waits:
        return None, ("Google's free daily limit is reached on all %d key(s). "
                      "Add another free key in Setup (doubles it), or wait a "
                      "while / tomorrow — nothing was changed in Word."
                      % len(clients))
    return None, "Helper call failed: " + last


def _call_text(cfg, prompt, max_chars=4000, note_fn=None):
    """Free-text generation: Groq first, same rotation policy.
    Returns (text, error)."""
    if cfg.get("GROQ_KEYS"):
        text, tag = _groq_call(cfg, "You write text for Word documents.",
                               prompt, False, note_fn)
        if tag is None and text:
            return text.strip()[:max_chars], None
    clients = _clients(cfg)
    if not clients:
        if cfg.get("GROQ_KEYS"):
            return None, ("Free engine failed right now — try again in a "
                          "minute. Nothing was changed in Word.")
        return None, ("No free key saved yet. Open Setup: add a Groq key "
                      "(recommended) or a Gemini key, paste, Save.")
    last = "unknown"
    waits = {}
    tried_wait = False
    for _round in range(3):
        for i, client in enumerate(clients):
            if waits.get(i, 0) > time.time():
                continue
            for attempt in range(2):
                try:
                    resp = client.models.generate_content(
                        model=cfg["GEMINI_MODEL"] or "gemini-3.6-flash",
                        contents=prompt,
                    )
                    return (resp.text or "").strip()[:max_chars], None
                except Exception as e:
                    last = str(e)[:300]
                    if "429" in str(e):
                        waits[i] = time.time() + (_retry_delay(str(e)) or 45)
                        if note_fn and len(clients) > 1:
                            note_fn("Key %d of %d hit Google's free limit — "
                                    "switching key..." % (i + 1, len(clients)))
                        break
                    if attempt < 1 and _is_retryable(str(e)):
                        time.sleep(2)
                        continue
                    break
        if waits and all(w > time.time() for w in waits.values()) \
                and not tried_wait:
            wait = min(max(int(w - time.time()) + 1, 1) for w in waits.values())
            wait = min(wait, 90)
            if note_fn:
                note_fn("Google's free limit is busy — waiting %ds, "
                        "then retrying..." % wait)
            time.sleep(wait)
            tried_wait = True
            continue
        break
    if waits:
        return None, ("Google's free daily limit is reached on all %d key(s). "
                      "Add another free key in Setup (doubles it), or wait a "
                      "while / tomorrow — nothing was changed in Word."
                      % len(clients))
    return None, "Helper call failed: " + last


def plan_action(user_text, note_fn=None):
    """Plan only. Returns plan dict, never touches the document."""
    user_text = (user_text or "").strip()
    if not user_text:
        return {"ok": False, "error": "Empty request."}
    cfg = load_config()
    if not cfg.get("GEMINI_KEYS") and not cfg.get("GROQ_KEYS"):
        return {"ok": False,
                "error": "No free key saved yet. Open Setup: add a Groq key "
                         "(recommended, thousands of free uses/day) or a "
                         "Gemini key, paste it, Save."}
    ctx = word_agent.get_context()
    if not ctx.get("ok"):
        return ctx  # e.g. Word busy / no document — pass through
    prompt = (
        SYSTEM
        + "\nSelected-or-open text scope: %s\nText:\n%s\n\nUser request: %s"
        % (ctx.get("scope"), ctx.get("text", ""), user_text)
    )
    plan, err = _call_json(cfg, prompt, note_fn)
    if err:
        return {"ok": False, "error": err}
    v = _validate(plan)
    if not v.get("ok"):
        # one repair retry: small models sometimes leave args empty or
        # return a null tool — ask once for the complete JSON
        problem = v.get("error", "unknown tool")
        plan2, err2 = _call_json(
            cfg, prompt + "\n\nYour previous reply was incomplete: %s. "
            "Send the complete JSON again with a valid tool and every "
            "required field filled." % problem, note_fn)
        if not err2:
            plan, v = plan2, _validate(plan2)
    if not v.get("ok"):
        if v.get("tool") is None:
            return {"ok": False,
                    "error": "I couldn't turn that into a Word action. Try "
                             "simpler words, e.g. 'underline everything' or "
                             "'write a leave letter'."}
        return v
    v["doc"] = ctx["doc"]
    # ask is answered now (read-only) so Approve just shows it
    if v["tool"] == "ask":
        answer, err = _call_text(
            cfg,
            "Using this Word %s text:\n%s\n\nAnswer (plain text, concise): %s"
            % (ctx.get("scope"), ctx.get("text", ""),
               v["args"]["question"]),
            note_fn=note_fn)
        if err:
            return {"ok": False, "error": err}
        v["answer_only"] = True
        v["answer"] = answer
    return v


# -- internet: free, no key needed --
UA = {"User-Agent": "AIAssistantForWord/1.0 (contact: local app; +https://aistudio.google.com/)"}


def _web_search(query, max_chars=3000):
    """DuckDuckGo + Wikipedia fallback. Returns text (may be thin)."""
    bits = []
    try:
        r = requests.get("https://api.duckduckgo.com/",
                         params={"q": query, "format": "json",
                                 "no_html": 1, "skip_disambig": 1},
                         headers=UA, timeout=15)
        d = r.json()
        if d.get("AbstractText"):
            bits.append("Source: %s\n%s" % (d.get("AbstractSource", "web"),
                                            d["AbstractText"]))
        for t in (d.get("RelatedTopics") or [])[:5]:
            if isinstance(t, dict) and t.get("Text"):
                bits.append(t["Text"])
    except Exception as e:
        bits.append("(web search hiccup: %s)" % str(e)[:100])
    if not bits:
        try:
            r = requests.get("https://en.wikipedia.org/w/api.php",
                             params={"action": "opensearch", "search": query,
                                     "limit": 3, "format": "json"},
                             headers=UA, timeout=15).json()
            for title, link in zip(r[1], r[3]):
                try:
                    s = requests.get(
                        "https://en.wikipedia.org/api/rest_v1/page/summary/"
                        + requests.utils.quote(title),
                        headers=UA, timeout=15).json()
                    bits.append("%s: %s (%s)"
                                % (title, s.get("extract", "")[:800], link))
                except Exception:
                    continue
        except Exception as e:
            bits.append("(encyclopedia hiccup: %s)" % str(e)[:100])
    text = "\n\n".join(bits)[:max_chars]
    return text or "No internet results found for this query."


def execute_action(plan, note_fn=None):
    """Run a validated plan. Call ONLY after user 'Yes, do it' (sidebar)."""
    if not plan.get("ok"):
        return {"ok": False, "error": "Refusing to run an invalid plan."}
    tool, args = plan.get("tool"), plan.get("args", {})
    cfg = load_config()
    try:
        if tool == "ask" and plan.get("answer_only"):
            return {"ok": True, "answer_only": True,
                    "answer": plan.get("answer", "")}
        if tool == "draft":
            text, err = _call_text(
                cfg, "Write this for a Word document (plain text, no "
                     "markdown, ready to paste): %s"
                     % args["instruction"], note_fn=note_fn)
            if err:
                return {"ok": False, "error": err}
            if args.get("place") == "cursor":
                return word_agent.insert_at_cursor(text)
            return word_agent.insert_at_end(text)
        if tool == "refine":
            ctx = word_agent.get_context()
            if not ctx.get("ok"):
                return ctx
            new_text, err = _call_text(
                cfg, "Rewrite this Word %s text per this instruction: %s\n"
                     "Return ONLY the rewritten text, same structure, plain "
                     "text, no commentary:\n%s"
                     % (ctx.get("scope"), args["instruction"],
                        ctx.get("text", "")),
                note_fn=note_fn)
            if err:
                return {"ok": False, "error": err}
            if ctx.get("scope") == "selection":
                return word_agent.insert_at_cursor(new_text)
            return word_agent.set_full_text(new_text)
        if tool == "research":
            if note_fn:
                note_fn("Searching the internet...")
            found = _web_search(args["query"])
            brief, err = _call_text(
                cfg, "Using these internet findings:\n%s\n\nWrite a short "
                     "Word-ready section (plain text, no markdown%s): %s"
                     % (found, "; end with [Sources: <names>]"
                        if "source" in found.lower() else "",
                        args.get("instruction") or
                        "summarise the key facts"),
                note_fn=note_fn)
            if err:
                return {"ok": False, "error": err}
            return word_agent.insert_at_end(brief)
        fn = word_agent.TOOLS.get(tool)
        if not fn:
            return {"ok": False, "error": "Unknown tool."}
        return fn(**args)
    except TypeError as e:
        return {"ok": False, "error": "Bad details: " + str(e)[:200]}


if __name__ == "__main__":
    import sys
    cmd = " ".join(sys.argv[1:]) or "summarise this document"
    out = plan_action(cmd)
    print(json.dumps(out, indent=2)[:2000])
