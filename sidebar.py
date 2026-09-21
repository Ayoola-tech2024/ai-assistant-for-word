"""Word Cloud Agent — chat with Word in plain English.

No tech skills needed:
  1. Open your Word document.
  2. Double-click WordCloudAgent.exe (or run: python sidebar.py).
  3. First time only: tap 'Get free key', paste it, Save.
  4. Type what you want ("add footer Confidential") and press Send,
     then 'Yes, do it' — watch Word change instantly.

Word is NEVER touched until you tap 'Yes, do it'.
"""
import ctypes
from ctypes import wintypes
import json
import threading
import tkinter as tk
import webbrowser
from tkinter import ttk
import win32gui

import agent
import word_agent

KEY_URL_GROQ = "https://console.groq.com/keys"
KEY_URL = "https://aistudio.google.com/apikey"

ARG_LABELS = {
    "add_footer": (("text", "Footer text"),),
    "replace_text": (("old", "Find this"), ("new", "Replace with")),
    "delete_para_containing": (("text", "Delete paragraphs containing"),),
    "insert_at_end": (("text", "Add at the end"),),
    "draft": (("instruction", "Write"), ("place", "Where")),
    "refine": (("instruction", "Improve like this"),),
    "ask": (("question", "Question"),),
    "research": (("query", "Look up"), ("instruction", "Write it as")),
    "review": (("instruction", "Review focus"),),
    "format_text": (("style", "Style"), ("page", "Page (blank=all)")),
    "find_text": (("text", "Find"),),
    "select_text": (("text", "Highlight"),),
    "go_to_page": (("page", "Page"),),
    "delete_page": (("page", "Page"),),
    "apply_heading": (("text", "Paragraph with"), ("level", "Level 1-3")),
    "insert_table": (("rows", "Rows"), ("cols", "Columns")),
    "insert_data_table": (("headers", "Columns"), ("rows", "Data"), ("style", "Style")),
    "make_list": (("kind", "Kind"),),
    "align_paras": (("how", "Align"),),
    "font_size": (("size", "Size"),),
    "font_color": (("color", "Color"),),
    "line_spacing": (("spacing", "Spacing"),),
    "set_orientation": (("orientation", "Orientation"),),
    "add_header": (("text", "Header text"),),
    "add_page_numbers": (),
    "word_count": (),
    "export_pdf": (),
    "undo_last": (("steps", "Steps"),),
    "toggle_track_changes": (("enabled", "Track Changes"),),
    "add_comment": (("text", "Comment"), ("target", "Target phrase")),
    "insert_toc": (("place", "Position"),),
    "add_watermark": (("text", "Watermark text"),),
}

TOOL_NAMES = {
    "add_footer": "Set footer",
    "replace_text": "Find & replace",
    "delete_para_containing": "Delete paragraphs",
    "insert_at_end": "Add text at end",
    "draft": "Write new content",
    "refine": "Improve the writing",
    "ask": "Answer (Word stays unchanged)",
    "research": "Research the internet + add to doc",
    "review": "Review document (margin comments)",
    "format_text": "Underline / bold / italic",
    "find_text": "Find in document",
    "select_text": "Highlight match",
    "go_to_page": "Go to page",
    "delete_page": "Delete page",
    "apply_heading": "Make heading",
    "insert_table": "Insert table",
    "insert_data_table": "Insert smart data table",
    "make_list": "Bullets / numbers",
    "align_paras": "Align text",
    "font_size": "Font size",
    "font_color": "Font color",
    "line_spacing": "Line spacing",
    "set_orientation": "Portrait / landscape",
    "add_header": "Set header",
    "add_page_numbers": "Add page numbers",
    "word_count": "Count words",
    "export_pdf": "Save as PDF",
    "undo_last": "Undo",
    "toggle_track_changes": "Toggle Track Changes",
    "add_comment": "Add margin comment",
    "insert_toc": "Insert Table of Contents",
    "add_watermark": "Add watermark",
}


class Sidebar:
    def __init__(self, root):
        self.root = root
        root.title("AI Assistant for Word")
        root.geometry("400x640")
        root.attributes("-topmost", True)
        self.pending = None
        self.busy_action = None  # ("plan", cmd) or ("execute", plan)
        self._start_global_hotkey()
        cfg = agent.load_config()
        if not cfg.get("GEMINI_KEYS") and not cfg.get("GROQ_KEYS"):
            self._build_setup()
        else:
            self._build_main()
        self.root.after(300, self.snap_to_word)

    # -- first-run setup: paste key once --
    def _build_setup(self):
        self.setup_frame = ttk.Frame(self.root, padding=20)
        self.setup_frame.pack(fill="both", expand=True)
        ttk.Label(self.setup_frame, text="Welcome! One quick setup.",
                  font=("Segoe UI", 12, "bold")).pack(pady=(10, 4))
        ttk.Label(self.setup_frame,
                  text="You need ONE free key. Groq is recommended\n"
                       "(thousands of free uses a day).",
                  font=("Segoe UI", 10)).pack(pady=4)
        ttk.Label(self.setup_frame, text="Option A (recommended): Groq key",
                  font=("Segoe UI", 10, "bold")).pack(pady=(12, 2))
        ttk.Button(self.setup_frame, text="Get Groq key (free)",
                   command=lambda: webbrowser.open(KEY_URL_GROQ)).pack()
        self.groq_entry = tk.Text(self.setup_frame, height=2,
                                  font=("Consolas", 9))
        self.groq_entry.pack(fill="x", pady=2)
        ttk.Label(self.setup_frame, text="Option B: Google key",
                  font=("Segoe UI", 10, "bold")).pack(pady=(8, 2))
        ttk.Button(self.setup_frame, text="Get Google key (free)",
                   command=lambda: webbrowser.open(KEY_URL)).pack()
        ttk.Label(self.setup_frame, text="Step 2: paste key(s) below",
                  font=("Segoe UI", 10, "bold")).pack(pady=(12, 2))
        ttk.Label(self.setup_frame,
                  text="One key is enough. For heavy days, paste extra keys\n"
                       "(one per line) — each adds more free daily use.",
                  font=("Segoe UI", 8)).pack()
        self.key_entry = tk.Text(self.setup_frame, height=4,
                                 font=("Consolas", 9))
        self.key_entry.pack(fill="x")
        self.setup_msg = tk.StringVar()
        ttk.Label(self.setup_frame, textvariable=self.setup_msg,
                  font=("Segoe UI", 9)).pack(pady=2)
        ttk.Button(self.setup_frame, text="Save & Start",
                   command=self._save_key).pack(pady=8)

    def _save_key(self):
        groq = self.groq_entry.get("1.0", "end")
        pasted = self.key_entry.get("1.0", "end")
        if len(groq.strip()) < 20 and len(pasted.strip()) < 20:
            self.setup_msg.set("Paste at least one full key first.")
            return
        try:
            ng = agent.save_api_key(groq, "groq") if len(groq.strip()) >= 20 else 0
            nm = agent.save_api_key(pasted) if len(pasted.strip()) >= 20 else 0
        except Exception as e:
            self.setup_msg.set("Could not save: " + str(e)[:150])
            return
        self.setup_frame.destroy()
        self._build_main()
        self._say("Setup done — %d Groq + %d Google key(s)." % (ng, nm))

    # -- main chat UI --
    def _build_main(self):
        # Open document header
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")
        self.doc_var = tk.StringVar(value="Open document: (checking...)")
        ttk.Label(top, textvariable=self.doc_var,
                  font=("Segoe UI", 9, "bold")).pack(
            side="left", fill="x", expand=True)
        ttk.Button(top, text="Refresh", command=self.refresh_doc).pack(
            side="right")
        ttk.Button(top, text="Snap", command=self.snap_to_word).pack(
            side="right", padx=(0, 2))
        ttk.Button(top, text="Templates", command=self._templates_popup).pack(
            side="right", padx=(0, 2))
        ttk.Button(top, text="Keys", command=self._keys_popup).pack(
            side="right", padx=(0, 2))

        # Plain-English input
        mid = ttk.Frame(self.root, padding=(8, 0, 8, 0))
        mid.pack(fill="x")
        ttk.Label(mid, text="Tell Word what to do (plain English):").pack(
            anchor="w")
        ttk.Label(mid, text='e.g. "write a leave request letter" or '
                            '"review this document" or '
                            '"turn on track changes"',
                  font=("Segoe UI", 8)).pack(anchor="w")
        self.cmd = tk.Text(mid, height=3, font=("Segoe UI", 11))
        self.cmd.pack(fill="x")
        self.cmd.bind("<Control-Return>", lambda e: self.send())
        btnrow = ttk.Frame(mid)
        btnrow.pack(fill="x", pady=4)
        self.send_btn = ttk.Button(btnrow, text="Send (Ctrl+Enter)",
                                   command=self.send)
        self.send_btn.pack(side="left")
        self.undo_btn = ttk.Button(btnrow, text="⟲ Undo",
                                   command=self.quick_undo)
        self.undo_btn.pack(side="left", padx=4)
        self.topmost_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(btnrow, text="Stay on top",
                        variable=self.topmost_var,
                        command=self._toggle_top).pack(side="right")

        # Preview + confirm
        prev = ttk.LabelFrame(self.root,
                              text="Check before it runs — Word changes only "
                                   "when you tap 'Yes, do it'",
                              padding=8)
        prev.pack(fill="both", expand=False, padx=8, pady=4)
        self.preview = tk.Text(prev, height=8, font=("Segoe UI", 10),
                               state="disabled", wrap="word")
        self.preview.pack(fill="both", expand=True)
        self.approve_btn = ttk.Button(prev, text="Yes, do it",
                                      command=self.approve, state="disabled")
        self.approve_btn.pack(side="left", pady=(4, 0))
        self.reject_btn = ttk.Button(prev, text="No, cancel",
                                     command=self.reject, state="disabled")
        self.reject_btn.pack(side="left", padx=4, pady=(4, 0))
        self.retry_btn = ttk.Button(prev, text="Try again",
                                    command=self.retry, state="disabled")
        self.retry_btn.pack(side="right", pady=(4, 0))

        # Activity
        logf = ttk.LabelFrame(self.root, text="What happened", padding=8)
        logf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.log = tk.Text(logf, font=("Segoe UI", 9), state="disabled",
                           wrap="word")
        self.log.pack(fill="both", expand=True)
        self.refresh_doc()

    # -- helpers --
    def _toggle_top(self):
        self.root.attributes("-topmost", self.topmost_var.get())

    def _note(self, msg):
        self.root.after(0, self._say, msg)

    def _keys_popup(self):
        pop = tk.Toplevel(self.root)
        pop.title("Free keys")
        pop.geometry("380x420")
        cfg = agent.load_config()
        ttk.Label(pop, text="Groq keys: %d (recommended, huge free limit)"
                  % len(cfg.get("GROQ_KEYS", [])), padding=8).pack(anchor="w")
        ttk.Button(pop, text="Get Groq key (free)",
                   command=lambda: webbrowser.open(KEY_URL_GROQ)).pack()
        groq_box = tk.Text(pop, height=3, font=("Consolas", 9))
        groq_box.pack(fill="x", padx=8, pady=2)
        ttk.Label(pop, text="Google keys: %d (20 free uses/day each)"
                  % len(cfg.get("GEMINI_KEYS", [])), padding=8).pack(anchor="w")
        ttk.Button(pop, text="Get Google key (free)",
                   command=lambda: webbrowser.open(KEY_URL)).pack()
        ttk.Label(pop, text="Paste extra key(s), one per line:",
                  padding=(8, 0, 8, 0)).pack(anchor="w")
        box = tk.Text(pop, height=3, font=("Consolas", 9))
        box.pack(fill="x", padx=8)
        msg = tk.StringVar()
        ttk.Label(pop, textvariable=msg, padding=(8, 0, 8, 0)).pack()
        def save():
            try:
                g = groq_box.get("1.0", "end").strip()
                m = box.get("1.0", "end").strip()
                tg = agent.save_api_key(g, "groq") if len(g) >= 20 else \
                    len(cfg.get("GROQ_KEYS", []))
                tm = agent.save_api_key(m) if len(m) >= 20 else \
                    len(cfg.get("GEMINI_KEYS", []))
                msg.set("Saved — %d Groq + %d Google." % (tg, tm))
                self._say("Keys updated: %d Groq + %d Google." % (tg, tm))
            except Exception as e:
                msg.set("Could not save: " + str(e)[:120])
        ttk.Button(pop, text="Save keys", command=save).pack(pady=6)

    def quick_undo(self):
        def work():
            res = word_agent.undo_last(1)
            if res.get("ok"):
                self.root.after(0, self._say, "⟲ Undid last change in Word.")
                self.root.after(0, self._show_preview, "Undid the last change in Word.")
            else:
                err = res.get("error") or res.get("msg") or "Failed"
                self.root.after(0, self._say, "Undo: " + str(err))
        self._say("Undoing last change in Word...")
        self._run_bg(work)

    def snap_to_word(self):
        try:
            hwnd = win32gui.FindWindow("OpusApp", None)
            if not hwnd or not win32gui.IsWindowVisible(hwnd):
                self._say("Word is not open. Open a document first to snap beside it.")
                return False
            rect = win32gui.GetWindowRect(hwnd)
            w_left, w_top, w_right, w_bottom = rect
            s_width = self.root.winfo_screenwidth()
            s_height = self.root.winfo_screenheight()
            w = 400
            h = min(max(w_bottom - w_top, 560), s_height - 60)
            x = min(w_right, s_width - w - 10)
            y = max(0, w_top)
            self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))
            self._say("Snapped beside Word.")
            return True
        except Exception as e:
            self._say("Could not snap to Word: " + str(e)[:100])
            return False

    def _templates_popup(self):
        pop = tk.Toplevel(self.root)
        pop.title("Templates")
        pop.geometry("420x460")
        pop.attributes("-topmost", True)
        ttk.Label(pop, text="One-Click Document Templates",
                  font=("Segoe UI", 11, "bold"), padding=(10, 10, 10, 4)).pack(anchor="w")
        ttk.Label(pop, text="Click a template to load it into the assistant:",
                  font=("Segoe UI", 9), padding=(10, 0, 10, 8)).pack(anchor="w")

        templates = [
            ("📄 Non-Disclosure Agreement (NDA)",
             "Write a comprehensive Non-Disclosure Agreement (NDA) between Disclosing Party and Receiving Party with standard confidentiality terms, exclusions, and a 2-year term."),
            ("📑 Business Project Proposal",
             "Draft a professional Project Proposal with Executive Summary, Scope of Work, Deliverables, Timeline, and Investment sections."),
            ("✉️ Formal Resignation Letter",
             "Write a polite, formal resignation letter stating a 2-week notice period, expressing gratitude for opportunities, and offering transition support."),
            ("📅 Meeting Minutes & Action Items",
             "Create a professional Meeting Minutes template with Date, Attendees, Agenda, Discussion Summary, and an Action Items table."),
            ("⚖️ Document Review & Margin Comments",
             "Review this document thoroughly and add margin comments highlighting any vague phrasing, risks, or grammar improvements."),
            ("📊 Comparison Table",
             "Create a competitor comparison table comparing 3 options across Features, Pricing, Pros, and Cons."),
            ("🏷️ Confidential Watermark & Header",
             "Add a CONFIDENTIAL watermark and set header Confidential — Internal Use Only."),
            ("📑 Table of Contents",
             "Insert a clickable Table of Contents at the beginning of the document."),
        ]

        frame = ttk.Frame(pop, padding=10)
        frame.pack(fill="both", expand=True)

        for title, prompt in templates:
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(fill="x", pady=2)
            def make_handler(p):
                return lambda: self._apply_template(p, pop)
            ttk.Button(btn_frame, text=title, command=make_handler(prompt)).pack(fill="x")

    def _apply_template(self, prompt, pop=None):
        if pop:
            pop.destroy()
        self.cmd.delete("1.0", "end")
        self.cmd.insert("end", prompt)
        self.cmd.focus_set()
        self._say("Loaded template: " + prompt[:50] + "...")

    def _start_global_hotkey(self):
        def worker():
            user32 = ctypes.windll.user32
            MOD_ALT = 0x0001
            MOD_CONTROL = 0x0002
            VK_W = 0x57  # 'W'
            HOTKEY_ID = 101
            if not user32.RegisterHotKey(None, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_W):
                return
            try:
                msg = wintypes.MSG()
                while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                    if msg.message == 0x0312:  # WM_HOTKEY
                        self.root.after(0, self._toggle_summon)
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))
            finally:
                user32.UnregisterHotKey(None, HOTKEY_ID)
        threading.Thread(target=worker, daemon=True).start()

    def _toggle_summon(self):
        try:
            if self.root.state() == "iconic" or not self.root.winfo_viewable():
                self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.cmd.focus_set()
            self._say("Summoned via Ctrl+Alt+W.")
        except Exception:
            pass

    def _say(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _show_preview(self, text):
        self.preview.config(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("end", text)
        self.preview.config(state="disabled")

    def _run_bg(self, fn):
        threading.Thread(target=self._guard(fn), daemon=True).start()

    def _guard(self, fn):
        def wrapper():
            try:
                fn()
            except Exception as e:  # never crash the UI
                self.root.after(0, self._say, "Something went wrong: "
                                + str(e)[:200])
                self.root.after(0, self._unlock_send)
        return wrapper

    def _unlock_send(self):
        self.send_btn.config(state="normal")

    # -- actions --
    def refresh_doc(self):
        def work():
            snap = word_agent.snapshot()
            if snap.get("ok"):
                txt = "Open document: %s" % snap["doc"].get("name")
            elif snap.get("need") == "close-dialog":
                txt = "Word is busy — close any pop-up in Word"
            else:
                txt = "No Word document open — open one first"
            self.root.after(0, self.doc_var.set, txt)
        self._run_bg(work)

    def send(self):
        cmd = self.cmd.get("1.0", "end").strip()
        if not cmd:
            return
        self.send_btn.config(state="disabled")
        self.pending = None
        self.busy_action = None
        self.approve_btn.config(state="disabled")
        self.reject_btn.config(state="disabled")
        self.retry_btn.config(state="disabled")
        self._show_preview("Thinking...")
        self._say("> " + cmd)
        self._run_bg(lambda: self._do_plan(cmd))

    def _do_plan(self, cmd):
        out = agent.plan_action(cmd, self._note)
        self.root.after(0, self._plan_done, cmd, out)

    def _plan_done(self, cmd, out):
        self.send_btn.config(state="normal")
        if out.get("need") == "close-dialog":
            self.busy_action = ("plan", cmd)
            self._show_preview("Word has a pop-up open (like Find or Save).\n"
                               "Close it in Word, then tap Try again.")
            self.retry_btn.config(state="normal")
            self._say("Word is busy — close its pop-up, then Try again.")
            return
        if not out.get("ok"):
            self._show_preview("Sorry, that didn't work:\n"
                               + str(out.get("error")))
            self._say("Not understood: " + str(out.get("error"))[:200])
            return
        self.pending = out
        if out.get("answer_only"):
            self._show_preview("Answer (Word stays unchanged):\n\n"
                               + out.get("answer", "")[:1200])
            self.approve_btn.config(text="OK", state="normal")
            self.reject_btn.config(state="disabled")
            self._say("Answer ready — tap OK.")
            return
        lines = ["I will: %s" % TOOL_NAMES.get(out["tool"], out["tool"])]
        for key, label in ARG_LABELS.get(out["tool"], ()):
            lines.append("%s: %s" % (label, out["args"].get(key, "")))
        lines.append("In: %s" % (out.get("doc") or {}).get("name", "?"))
        self._show_preview("\n".join(lines))
        self.approve_btn.config(text="Yes, do it", state="normal")
        self.reject_btn.config(state="normal")
        self._say("Check it above — 'Yes, do it' or 'No, cancel'.")

    def approve(self):
        if not self.pending:
            return
        plan = self.pending
        self.pending = None
        self.approve_btn.config(state="disabled")
        self.reject_btn.config(state="disabled")
        self.send_btn.config(state="disabled")
        self._say("Doing it in %s..."
                  % (plan.get("doc") or {}).get("name", "Word"))
        self._run_bg(lambda: self._do_execute(plan))

    def _do_execute(self, plan):
        out = agent.execute_action(plan, self._note)
        self.root.after(0, self._exec_done, plan, out)

    def _exec_done(self, plan, out):
        self.send_btn.config(state="normal")
        if out.get("need") == "close-dialog":
            self.busy_action = ("execute", plan)
            self._show_preview("Word has a pop-up open.\nClose it in Word, "
                               "then tap Try again.")
            self.retry_btn.config(state="normal")
            self._say("Word is busy — close its pop-up, then Try again.")
            return
        if out.get("answer_only"):
            self._show_preview("Answer (Word unchanged):\n\n"
                               + str(out.get("answer", ""))[:1500])
            self._say("Answered.")
            self.approve_btn.config(text="Yes, do it")
            return
        if out.get("ok"):
            if out.get("report"):
                self._show_preview("Done!\n\n" + str(out["report"])[:800])
            else:
                self._show_preview("Done! Look at Word — the change is there.")
            self._say("Done: %s" % TOOL_NAMES.get(plan["tool"],
                                                 plan["tool"]))
        else:
            self._show_preview("Sorry, that failed:\n" + str(out.get("error")))
            self._say("Failed: " + str(out.get("error"))[:200])
        self.refresh_doc()

    def reject(self):
        self.pending = None
        self.approve_btn.config(text="Yes, do it", state="disabled")
        self.reject_btn.config(state="disabled")
        self._show_preview("Cancelled — Word was not touched.")
        self._say("Cancelled.")

    def retry(self):
        if not self.busy_action:
            return
        kind, payload = self.busy_action
        self.busy_action = None
        self.retry_btn.config(state="disabled")
        if kind == "plan":
            self._show_preview("Trying again...")
            self._run_bg(lambda: self._do_plan(payload))
        else:
            self._show_preview("Trying again...")
            self._run_bg(lambda: self._do_execute(payload))


def main():
    root = tk.Tk()
    Sidebar(root)
    root.mainloop()


if __name__ == "__main__":
    main()
