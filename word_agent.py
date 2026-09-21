"""Phase 2: 4 Word COM tools with busy-catch + ActiveDocument pin.
Windows + Word 2013+ via pywin32. No LLM here — pure tools.
Every tool returns {"ok": bool, ...}. Busy Word (modal dialog open)
returns {"ok": False, "need": "close-dialog", ...} instead of crashing.
"""
import pythoncom
import pywintypes
import time
import win32com.client

BUSY_HRESULT = 0x80010001  # Call was rejected by callee (modal dialog open)


def _bind():
    """Bind to running Word, pin ActiveDocument. Returns (app, doc) or raises."""
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("Word.Application")
    except Exception as e:
        raise RuntimeError(
            "No running Word found. Open Word with a document first."
        ) from e
    if app.Documents.Count == 0:
        raise RuntimeError("Word is open but has no documents.")
    doc = app.ActiveDocument
    return app, doc


def _doc_id(doc):
    try:
        return {"name": doc.Name, "path": doc.FullName}
    except Exception:
        return {"name": "?", "path": "?"}


def _is_busy(err):
    try:
        # pywintypes.com_error: (hresult, text, exc, arg)
        hr = err.args[0] if err.args else 0
        # hr may be negative signed; compare low 32 bits
        return (hr & 0xFFFFFFFF) == BUSY_HRESULT
    except Exception:
        return False


def _busy_result(action):
    return {
        "ok": False,
        "need": "close-dialog",
        "action": action,
        "msg": "Word is busy — close any open dialog (Find & Replace, Save As, "
               "right-click menu) then Retry.",
    }


def add_footer(text):
    """Set primary footer text on every section. Returns dict."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        for sec in doc.Sections:
            footer = sec.Footers(1)  # 1 = wdHeaderFooterPrimary
            footer.Range.Text = text
        return {"ok": True, "action": "add_footer", "doc": info}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("add_footer")
        return {"ok": False, "action": "add_footer", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "add_footer", "error": str(e)[:300]}


def replace_text(old, new):
    """Replace all occurrences of old with new (paragraph level).
    Note v1: resets intra-paragraph character formatting on touched
    paragraphs (Word Find-Replace via COM is unreliable from Python).
    Returns count of touched paragraphs."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        touched = 0
        for p in doc.Paragraphs:
            try:
                t = p.Range.Text or ""
            except pywintypes.com_error as e:
                if _is_busy(e):
                    return _busy_result("replace_text")
                raise
            if old in t:
                p.Range.Text = t.replace(old, new)
                touched += 1
        return {"ok": True, "action": "replace_text", "doc": info,
                "replaced_paras": touched, "found": touched > 0}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("replace_text")
        return {"ok": False, "action": "replace_text", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "replace_text", "error": str(e)[:300]}


def delete_para_containing(text):
    """Delete paragraphs containing text (case-insensitive substring)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        needle = text.lower()
        removed = 0
        # iterate backwards so deletions don't shift indexes
        paras = list(doc.Paragraphs)
        for p in reversed(paras):
            try:
                if needle in (p.Range.Text or "").lower():
                    p.Range.Delete()
                    removed += 1
            except pywintypes.com_error as e:
                if _is_busy(e):
                    return _busy_result("delete_para_containing")
                raise
        return {"ok": True, "action": "delete_para_containing",
                "doc": info, "removed": removed}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("delete_para_containing")
        return {"ok": False, "action": "delete_para_containing",
                "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "delete_para_containing",
                "error": str(e)[:300]}


def insert_at_end(text):
    """Append a new paragraph at end of document."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        end = doc.Content
        end.Collapse(0)  # 0 = wdCollapseEnd
        end.Text = text + "\r"
        return {"ok": True, "action": "insert_at_end", "doc": info}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_at_end")
        return {"ok": False, "action": "insert_at_end", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_at_end", "error": str(e)[:300]}


def snapshot(max_chars=2000):
    """Read-only doc preview for the LLM prompt. Never fails the loop."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        txt = doc.Content.Text or ""
        return {"ok": True, "doc": info,
                "chars": len(txt), "preview": txt[:max_chars]}
    except pywintypes.com_error as e:
        if _is_busy(e):
            d = _busy_result("snapshot")
            return d
        return {"ok": False, "action": "snapshot", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "snapshot", "error": str(e)[:300]}


def get_context(max_chars=6000):
    """Read selected text if the user highlighted something, else full doc.
    Returns {scope: 'selection'|'document', text: ...} for LLM input."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        try:
            sel_text = (app.Selection.Range.Text or "").strip()
        except Exception:
            sel_text = ""
        # Word selection often includes trailing \r or is just insertion point
        if len(sel_text) > 1:
            return {"ok": True, "doc": info, "scope": "selection",
                    "text": sel_text[:max_chars]}
        txt = doc.Content.Text or ""
        return {"ok": True, "doc": info, "scope": "document",
                "text": txt[:max_chars], "chars": len(txt)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("get_context")
        return {"ok": False, "action": "get_context", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "get_context", "error": str(e)[:300]}


def set_full_text(text):
    """Replace the ENTIRE document body (used by refine/draft-overwrite).
    Keeps sections/footers. Returns dict."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        rng = doc.Content
        rng.Text = text if text.endswith("\r") else text + "\r"
        return {"ok": True, "action": "set_full_text", "doc": info,
                "chars": len(text)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("set_full_text")
        return {"ok": False, "action": "set_full_text", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "set_full_text", "error": str(e)[:300]}


def insert_at_cursor(text):
    """Insert text where the cursor is (falls back to end of doc)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        try:
            sel = app.Selection
            if sel is not None and doc.Name == app.ActiveDocument.Name:
                sel.Range.Text = text
                return {"ok": True, "action": "insert_at_cursor",
                        "doc": info, "at": "cursor"}
        except pywintypes.com_error as e:
            if _is_busy(e):
                return _busy_result("insert_at_cursor")
            raise
        return insert_at_end(text)
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_at_cursor")
        return {"ok": False, "action": "insert_at_cursor",
                "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_at_cursor",
                "error": str(e)[:300]}


def format_text(style, page=None):
    """Underline / bold / italic — in order of preference:
    the given page number, the highlighted text, or the whole document."""
    style = (style or "").strip().lower()
    if style not in ("underline", "bold", "italic"):
        return {"ok": False, "action": "format_text",
                "error": "Unknown style: %r" % (style,)}
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        scope = "document"
        ranges = None
        if page:
            paras = _paras_on_page(doc, int(page))
            if paras is None:
                return _busy_result("format_text")
            if not paras:
                return {"ok": False, "action": "format_text",
                        "error": "Page %s looks empty or missing." % page}
            ranges = [p.Range for p in paras]
            scope = "page %s" % page
        if ranges is None:
            try:
                sel_text = (app.Selection.Range.Text or "")
            except Exception:
                sel_text = ""
            if len(sel_text.strip()) > 1:
                ranges = [app.Selection.Range]
                scope = "selection"
            else:
                ranges = [doc.Content]
        for rng in ranges:
            if style == "underline":
                rng.Font.Underline = 1  # wdUnderlineSingle
            elif style == "bold":
                rng.Font.Bold = True
            else:
                rng.Font.Italic = True
        return {"ok": True, "action": "format_text", "doc": info,
                "style": style, "scope": scope}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("format_text")
        return {"ok": False, "action": "format_text", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "format_text", "error": str(e)[:300]}


# -- page helpers: Word pages are layout-based; these work on the open doc --
WD_GOTO_PAGE = 1      # wdGoToPage
WD_GOTO_ABS = 1       # wdGoToAbsolute
WD_INFO_PAGE = 3      # wdActiveEndPageNumber


def _para_page(p):
    try:
        return int(p.Range.Information(WD_INFO_PAGE))
    except Exception:
        return None


def _paras_on_page(doc, n):
    try:
        return [p for p in doc.Paragraphs if _para_page(p) == n]
    except pywintypes.com_error as e:
        if _is_busy(e):
            return None
        raise


def _page_range(doc, n):
    """Range covering page n (to start of n+1, or doc end)."""
    start = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABS, Count=n)
    try:
        nxt = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABS, Count=n + 1)
        end_pos = nxt.Start if nxt.Start > start.Start else doc.Content.End
    except Exception:
        end_pos = doc.Content.End
    return doc.Range(start.Start, end_pos)


def _target_range(app, doc):
    """User's highlight if any, else whole document. Returns (rng, scope)."""
    try:
        sel = app.Selection.Range.Text or ""
    except Exception:
        sel = ""
    if len(sel.strip()) > 1:
        return app.Selection.Range, "selection"
    return doc.Content, "document"


COLORS = {"black": 0, "white": 16777215, "red": 255, "blue": 16711680,
          "green": 32768, "darkblue": 8388608, "darkred": 128,
          "grey": 8421504, "gray": 8421504, "orange": 26367,
          "purple": 8388736}


def find_text(text):
    """Count occurrences, list pages, and jump the cursor to the first match
    so the user SEES it. Read-only except moving the cursor."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        needle = (text or "").strip()
        if not needle:
            return {"ok": False, "action": "find_text",
                    "error": "Empty search text."}
        hits, pages = 0, set()
        first = None
        for p in doc.Paragraphs:
            t = p.Range.Text or ""
            low = t.lower()
            idx = low.find(needle.lower())
            while idx != -1:
                hits += 1
                pg = _para_page(p)
                if pg:
                    pages.add(pg)
                if first is None:
                    r = p.Range.Duplicate
                    r.Start = p.Range.Start + idx
                    r.End = r.Start + len(needle)
                    first = r
                idx = low.find(needle.lower(), idx + 1)
        if hits == 0:
            return {"ok": True, "action": "find_text", "doc": info,
                    "report": "'%s' was not found in %s." % (needle,
                                                             info["name"])}
        try:
            first.Select()  # jump so the user sees it
        except Exception:
            pass
        pg_txt = (" on page(s) " + ", ".join(map(str, sorted(pages)))
                  if pages else "")
        return {"ok": True, "action": "find_text", "doc": info,
                "report": "Found '%s' %d time(s)%s. "
                          "Cursor jumped to the first match."
                          % (needle, hits, pg_txt)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("find_text")
        return {"ok": False, "action": "find_text", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "find_text", "error": str(e)[:300]}


def select_text(text):
    """Highlight the first match of text in the Word window."""
    r = find_text(text)
    if r.get("ok") and "not found" not in r.get("report", ""):
        r["action"] = "select_text"
    return r


def go_to_page(page):
    """Move the cursor to the top of page N (visible jump)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        rng = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABS, Count=int(page))
        try:
            rng.Select()
        except Exception:
            pass
        return {"ok": True, "action": "go_to_page", "doc": info,
                "report": "Cursor is now at the top of page %s." % page}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("go_to_page")
        return {"ok": False, "action": "go_to_page", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "go_to_page", "error": str(e)[:300]}


def delete_page(page):
    """Delete a whole page by number."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        rng = _page_range(doc, int(page))
        if len((rng.Text or "").strip()) == 0:
            return {"ok": False, "action": "delete_page",
                    "error": "Page %s looks empty or missing." % page}
        rng.Delete()
        return {"ok": True, "action": "delete_page", "doc": info,
                "report": "Page %s deleted." % page}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("delete_page")
        return {"ok": False, "action": "delete_page", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "delete_page", "error": str(e)[:300]}


def apply_heading(text, level=1):
    """Turn the first paragraph containing text into a Heading 1-3."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        level = max(1, min(3, int(level)))
        needle = (text or "").strip().lower()
        for p in doc.Paragraphs:
            if needle and needle in (p.Range.Text or "").lower():
                p.Style = doc.Styles("Heading %d" % level)
                try:
                    app.Selection.GoTo(What=WD_GOTO_PAGE,
                                       Which=WD_GOTO_ABS,
                                       Count=_para_page(p) or 1)
                except Exception:
                    pass
                return {"ok": True, "action": "apply_heading", "doc": info,
                        "report": "Made Heading %d: '%s'." % (level,
                        (p.Range.Text or "").strip()[:80])}
        return {"ok": False, "action": "apply_heading",
                "error": "No paragraph contains '%s'." % text}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("apply_heading")
        return {"ok": False, "action": "apply_heading", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "apply_heading", "error": str(e)[:300]}


def insert_table(rows, cols, place="end"):
    """Insert an empty rows×cols table at cursor or end of document."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        rows, cols = max(1, min(20, int(rows))), max(1, min(10, int(cols)))
        if (place or "end") == "cursor":
            try:
                rng = app.Selection.Range
            except Exception:
                rng = doc.Content
                rng.Collapse(0)
        else:
            rng = doc.Content
            rng.Collapse(0)
            rng.Text = "\r"
            rng.Collapse(0)
        rng.Tables.Add(rng, rows, cols)
        return {"ok": True, "action": "insert_table", "doc": info,
                "report": "Inserted a %dx%d table." % (rows, cols)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_table")
        return {"ok": False, "action": "insert_table", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_table", "error": str(e)[:300]}


def make_list(kind="bullets"):
    """Bulleted or numbered list from highlight (or whole document)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        rng, scope = _target_range(app, doc)
        if (kind or "bullets").startswith("number"):
            rng.ListFormat.ApplyNumberDefault()
            what = "numbered list"
        else:
            rng.ListFormat.ApplyBulletDefault()
            what = "bulleted list"
        return {"ok": True, "action": "make_list", "doc": info,
                "report": "Applied %s to %s." % (what, scope)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("make_list")
        return {"ok": False, "action": "make_list", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "make_list", "error": str(e)[:300]}


ALIGNS = {"left": 0, "center": 1, "right": 2, "justify": 3}


def align_paras(how="left"):
    """Align highlight (or whole document): left|center|right|justify."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        how = (how or "left").strip().lower()
        if how not in ALIGNS:
            return {"ok": False, "action": "align_paras",
                    "error": "Align how? left, center, right or justify."}
        rng, scope = _target_range(app, doc)
        rng.ParagraphFormat.Alignment = ALIGNS[how]
        return {"ok": True, "action": "align_paras", "doc": info,
                "report": "Aligned %s: %s." % (scope, how)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("align_paras")
        return {"ok": False, "action": "align_paras", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "align_paras", "error": str(e)[:300]}


def font_size(size):
    """Set font size of highlight (or whole document)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        size = max(8, min(72, int(size)))
        rng, scope = _target_range(app, doc)
        rng.Font.Size = size
        return {"ok": True, "action": "font_size", "doc": info,
                "report": "Font size %dpt applied to %s." % (size, scope)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("font_size")
        return {"ok": False, "action": "font_size", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "font_size", "error": str(e)[:300]}


def font_color(color):
    """Set font color of highlight (or whole document)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        c = (color or "").strip().lower()
        if c not in COLORS:
            return {"ok": False, "action": "font_color",
                    "error": "Colors: " + ", ".join(sorted(COLORS))}
        rng, scope = _target_range(app, doc)
        rng.Font.Color = COLORS[c]
        return {"ok": True, "action": "font_color", "doc": info,
                "report": "Font color %s applied to %s." % (c, scope)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("font_color")
        return {"ok": False, "action": "font_color", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "font_color", "error": str(e)[:300]}


SPACING = {"single": 0, "1.5": 1, "double": 2}


def line_spacing(spacing="single"):
    """Line spacing of highlight (or whole document): single|1.5|double."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        s = (spacing or "single").strip().lower()
        if s not in SPACING:
            return {"ok": False, "action": "line_spacing",
                    "error": "Spacing: single, 1.5 or double."}
        rng, scope = _target_range(app, doc)
        rng.ParagraphFormat.LineSpacingRule = SPACING[s]
        return {"ok": True, "action": "line_spacing", "doc": info,
                "report": "%s spacing applied to %s." % (s, scope)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("line_spacing")
        return {"ok": False, "action": "line_spacing", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "line_spacing", "error": str(e)[:300]}


def set_orientation(orientation="portrait"):
    """Whole document: portrait or landscape."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        o = (orientation or "portrait").strip().lower()
        if o.startswith("land"):
            w, h = 0, 1  # wdOrientLandscape=1
            for sec in doc.Sections:
                sec.PageSetup.Orientation = 1
            what = "landscape"
        else:
            for sec in doc.Sections:
                sec.PageSetup.Orientation = 0
            what = "portrait"
        return {"ok": True, "action": "set_orientation", "doc": info,
                "report": "Page orientation: %s." % what}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("set_orientation")
        return {"ok": False, "action": "set_orientation",
                "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "set_orientation",
                "error": str(e)[:300]}


def add_header(text):
    """Set header text on every section."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        for sec in doc.Sections:
            sec.Headers(1).Range.Text = text
        return {"ok": True, "action": "add_header", "doc": info,
                "report": "Header set."}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("add_header")
        return {"ok": False, "action": "add_header", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "add_header", "error": str(e)[:300]}


def add_page_numbers():
    """Add centered 'Page X' numbers in the footer."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        for sec in doc.Sections:
            footer = sec.Footers(1)
            footer.Range.Text = ""
            footer.PageNumbers.Add(PageNumberAlignment=1)  # centered
        return {"ok": True, "action": "add_page_numbers", "doc": info,
                "report": "Page numbers added in the footer."}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("add_page_numbers")
        return {"ok": False, "action": "add_page_numbers",
                "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "add_page_numbers",
                "error": str(e)[:300]}


def word_count():
    """Count words, characters, paragraphs — shown, nothing changed."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        txt = doc.Content.Text or ""
        words = len(txt.split())
        paras = len([p for p in doc.Paragraphs
                     if (p.Range.Text or "").strip()])
        return {"ok": True, "action": "word_count", "doc": info,
                "report": "%s: %d words, %d characters, %d paragraphs."
                          % (info["name"], words, len(txt), paras)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("word_count")
        return {"ok": False, "action": "word_count", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "word_count", "error": str(e)[:300]}


def export_pdf(output_path=None, open_folder=True):
    """Save a PDF next to the document (or Desktop if unsaved) and reveal in Explorer."""
    try:
        import os
        import re
        import subprocess
        app, doc = _bind()
        info = _doc_id(doc)
        path = info.get("path", "") or ""
        if output_path:
            pdf = output_path
        elif path and not path.startswith("Document") and os.path.exists(path):
            pdf = os.path.splitext(path)[0] + ".pdf"
        else:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            doc_name = (info.get("name") or "Document").replace(".docx", "").replace(".doc", "")
            clean_name = re.sub(r'[\\/*?:"<>|]', "", doc_name).strip() or "Word_Document"
            pdf = os.path.join(desktop, "%s.pdf" % clean_name)
            if os.path.exists(pdf):
                pdf = os.path.join(desktop, "%s_%d.pdf" % (clean_name, int(time.time())))

        doc.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF = 17

        if open_folder and os.path.exists(pdf):
            try:
                subprocess.Popen(['explorer', '/select,', os.path.normpath(pdf)])
            except Exception:
                pass

        return {"ok": True, "action": "export_pdf", "doc": info,
                "pdf_path": pdf,
                "report": "PDF saved: %s" % pdf}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("export_pdf")
        return {"ok": False, "action": "export_pdf", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "export_pdf", "error": str(e)[:300]}


def get_document_analytics():
    """Compute detailed reading telemetry, word stats, and Flesch-Kincaid readability."""
    try:
        import re
        app, doc = _bind()
        info = _doc_id(doc)

        text = doc.Content.Text or ""
        text = text.replace("\r", "\n").replace("\x07", "").replace("\x0c", "").strip()

        try:
            words_count = doc.ComputeStatistics(0)  # wdStatisticWords
            chars_count = doc.ComputeStatistics(3)  # wdStatisticCharacters
            paras_count = doc.ComputeStatistics(4)  # wdStatisticParagraphs
            pages_count = doc.ComputeStatistics(2)  # wdStatisticPages
        except Exception:
            words_count = len(re.findall(r'\b\w+\b', text))
            chars_count = len(text)
            paras_count = len([p for p in text.split("\n") if p.strip()])
            pages_count = 1

        words = re.findall(r'\b[a-zA-Z0-9_\'-]+\b', text)
        sentences = [s.strip() for s in re.split(r'[.!?]+', text) if s.strip()]

        num_words = max(1, len(words) or words_count)
        num_sentences = max(1, len(sentences))

        def _syllables(w):
            w = w.lower().strip()
            if len(w) <= 3:
                return 1
            w = re.sub(r'(?:[^laeiouy]|ed|es|e)$', '', w)
            w = re.sub(r'^y', '', w)
            matches = re.findall(r'[aeiouy]{1,2}', w)
            return max(1, len(matches))

        total_syllables = sum(_syllables(w) for w in words) if words else num_words

        # Flesch Reading Ease: 206.835 - 1.015*(words/sentences) - 84.6*(syllables/words)
        flesch_score = 206.835 - (1.015 * (num_words / num_sentences)) - (84.6 * (total_syllables / num_words))
        flesch_score = max(0.0, min(100.0, round(flesch_score, 1)))

        # Flesch-Kincaid Grade Level: 0.39*(words/sentences) + 11.8*(syllables/words) - 15.59
        fk_grade = (0.39 * (num_words / num_sentences)) + (11.8 * (total_syllables / num_words)) - 15.59
        fk_grade = max(1.0, round(fk_grade, 1))

        if flesch_score >= 90:
            level = "Very Easy (5th Grade)"
        elif flesch_score >= 80:
            level = "Easy (6th Grade)"
        elif flesch_score >= 70:
            level = "Fairly Easy (7th Grade)"
        elif flesch_score >= 60:
            level = "Standard (8th-9th Grade - Plain English)"
        elif flesch_score >= 50:
            level = "Fairly Difficult (High School)"
        elif flesch_score >= 30:
            level = "Difficult (College)"
        else:
            level = "Very Difficult (Academic/Technical)"

        read_time_min = max(1, round(num_words / 200))
        speak_time_min = max(1, round(num_words / 130))

        report = (
            "📊 Document Telemetry & Health:\n"
            "• Pages: %d | Paragraphs: %d\n"
            "• Words: %s | Characters: %s\n"
            "• Reading Time: ~%d min (silent reading)\n"
            "• Speaking Time: ~%d min (presentation pace)\n"
            "• Readability: %.1f / 100 (%s)\n"
            "• US Grade Level: %.1f"
            % (pages_count, paras_count, f"{num_words:,}", f"{chars_count:,}",
               read_time_min, speak_time_min, flesch_score, level, fk_grade)
        )

        return {
            "ok": True,
            "action": "get_document_analytics",
            "doc": info,
            "report": report,
            "stats": {
                "pages": pages_count,
                "paragraphs": paras_count,
                "words": num_words,
                "characters": chars_count,
                "reading_time_min": read_time_min,
                "speaking_time_min": speak_time_min,
                "flesch_score": flesch_score,
                "flesch_level": level,
                "grade_level": fk_grade,
            }
        }
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("get_document_analytics")
        return {"ok": False, "action": "get_document_analytics", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "get_document_analytics", "error": str(e)[:300]}


def undo_last(steps=1):
    """Undo the last change(s) in Word."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        try:
            steps = max(1, min(10, int(steps)))
        except Exception:
            steps = 1
        for _ in range(steps):
            if not doc.Undo():
                break
        return {"ok": True, "action": "undo_last", "doc": info,
                "report": "Undid the last change(s)."}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("undo_last")
        return {"ok": False, "action": "undo_last", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "undo_last", "error": str(e)[:300]}


def toggle_track_changes(enabled=True):
    """Enable or disable Word Track Changes (redline revision mode)."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        if isinstance(enabled, str):
            enabled = enabled.lower() not in ("false", "off", "0", "disable")
        doc.TrackRevisions = bool(enabled)
        status = "enabled (redlines visible)" if doc.TrackRevisions else "disabled"
        return {"ok": True, "action": "toggle_track_changes", "doc": info,
                "report": "Track Changes is now %s." % status}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("toggle_track_changes")
        return {"ok": False, "action": "toggle_track_changes", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "toggle_track_changes", "error": str(e)[:300]}


def add_comment(text, target=None):
    """Add a margin comment to highlighted text, specific target text, or cursor."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        text = (text or "").strip()
        if not text:
            return {"ok": False, "action": "add_comment", "error": "Comment text cannot be empty."}
        rng = None
        target_str = (target or "").strip()
        if target_str:
            needle = target_str.lower()
            for p in doc.Paragraphs:
                t = (p.Range.Text or "").lower()
                idx = t.find(needle)
                if idx != -1:
                    r = p.Range.Duplicate
                    r.Start = p.Range.Start + idx
                    r.End = r.Start + len(target_str)
                    rng = r
                    break
        if rng is None:
            sel = app.Selection.Range.Text or ""
            if len(sel.strip()) > 1:
                rng = app.Selection.Range
            else:
                rng = app.Selection.Range
        doc.Comments.Add(Range=rng, Text=text)
        preview = text[:80] + ("..." if len(text) > 80 else "")
        return {"ok": True, "action": "add_comment", "doc": info,
                "report": "Margin comment added: '%s'" % preview}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("add_comment")
        return {"ok": False, "action": "add_comment", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "add_comment", "error": str(e)[:300]}


def insert_toc(place="cursor"):
    """Insert a native clickable Table of Contents based on Headings 1-3."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        p = (place or "cursor").strip().lower()
        if p in ("start", "top"):
            rng = doc.Range(0, 0)
        elif p == "cursor":
            try:
                rng = app.Selection.Range
            except Exception:
                rng = doc.Range(0, 0)
        else:
            rng = doc.Content
            rng.Collapse(0)  # 0 = wdCollapseEnd
        doc.TablesOfContents.Add(Range=rng, UseHeadingStyles=True,
                                 UpperHeadingLevel=1, LowerHeadingLevel=3)
        return {"ok": True, "action": "insert_toc", "doc": info,
                "report": "Table of Contents inserted."}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_toc")
        return {"ok": False, "action": "insert_toc", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_toc", "error": str(e)[:300]}


def insert_data_table(headers, rows, place="end", style="Grid Table 4 - Accent 1"):
    """Insert a table populated with headers and row data, styled professionally."""
    import json
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        if isinstance(headers, str):
            try:
                headers = json.loads(headers)
            except Exception:
                headers = [h.strip() for h in headers.split(",") if h.strip()]
        if isinstance(rows, str):
            try:
                rows = json.loads(rows)
            except Exception:
                rows = [[c.strip() for c in line.split(",")]
                        for line in rows.splitlines() if line.strip()]
        if not headers or not rows:
            return {"ok": False, "action": "insert_data_table",
                    "error": "Both headers and rows are required."}
        num_cols = len(headers)
        num_rows = len(rows) + 1  # 1 header row + data rows
        if (place or "end") == "cursor":
            try:
                rng = app.Selection.Range
            except Exception:
                rng = doc.Content
                rng.Collapse(0)
        else:
            rng = doc.Content
            rng.Collapse(0)
            rng.Text = "\r"
            rng.Collapse(0)
        tbl = doc.Tables.Add(rng, num_rows, num_cols)
        try:
            tbl.Style = style
        except Exception:
            try:
                tbl.Style = "Table Grid"
            except Exception:
                pass
        # Populate headers
        for c, h in enumerate(headers, 1):
            cell = tbl.Cell(1, c)
            cell.Range.Text = str(h)
            cell.Range.Bold = True
        # Populate rows
        for r_idx, row in enumerate(rows, 2):
            for c_idx, val in enumerate(row[:num_cols], 1):
                tbl.Cell(r_idx, c_idx).Range.Text = str(val)
        return {"ok": True, "action": "insert_data_table", "doc": info,
                "report": "Inserted a populated %dx%d table." % (num_rows, num_cols)}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_data_table")
        return {"ok": False, "action": "insert_data_table", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_data_table", "error": str(e)[:300]}


def add_watermark(text="CONFIDENTIAL"):
    """Add a diagonal semi-transparent watermark across all sections."""
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        text = (text or "CONFIDENTIAL").strip().upper()
        try:
            app.ActiveWindow.View.Type = 3  # wdPrintView = 3
        except Exception:
            pass
        for sec in doc.Sections:
            header = sec.Headers(1)  # 1 = wdHeaderFooterPrimary
            shape = header.Shapes.AddTextEffect(0, text, "Calibri", 54, 0, 0, 0, 0)
            shape.Name = "AssistantWatermark"
            shape.Rotation = 315
            shape.Fill.Visible = -1  # msoTrue
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = 12632256  # light gray
            shape.Line.Visible = 0  # msoFalse
            shape.RelativeHorizontalPosition = 0
            shape.RelativeVerticalPosition = 0
            shape.Left = -999995  # wdShapeCenter
            shape.Top = -999995  # wdShapeCenter
        return {"ok": True, "action": "add_watermark", "doc": info,
                "report": "Watermark '%s' added across all sections." % text}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("add_watermark")
        return {"ok": False, "action": "add_watermark", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "add_watermark", "error": str(e)[:300]}


def insert_chart(kind="bar", title="Chart", labels="[]", values="[]", place="end"):
    """Insert a professional matplotlib chart (bar, line, or pie) into Word."""
    import json
    import os
    import tempfile
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        if isinstance(labels, str):
            try:
                labels = json.loads(labels)
            except Exception:
                labels = [x.strip() for x in labels.split(",") if x.strip()]
        if isinstance(values, str):
            try:
                values = json.loads(values)
            except Exception:
                values = [float(x.strip()) for x in values.split(",") if x.strip()]
        values = [float(v) for v in values]
        if not labels or not values or len(labels) != len(values):
            return {"ok": False, "action": "insert_chart",
                    "error": "Labels and numerical values must be provided and have matching length."}

        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots(figsize=(6, 3.5), dpi=150)
        kind = (kind or "bar").strip().lower()
        if kind == "pie":
            ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=140)
        elif kind == "line":
            ax.plot(labels, values, marker="o", color="#2b579a", linewidth=2.5, markersize=6)
            ax.grid(True, linestyle="--", alpha=0.5)
        else:  # bar
            colors = ["#2b579a", "#418ab3", "#5ba4cf", "#7fc0eb", "#a6dcef"]
            bar_colors = [colors[i % len(colors)] for i in range(len(labels))]
            ax.bar(labels, values, color=bar_colors, edgecolor="#1f3f70", width=0.55)
            ax.grid(axis="y", linestyle="--", alpha=0.4)

        if title:
            ax.set_title(str(title), fontsize=12, fontweight="bold", pad=12)
        plt.tight_layout()

        chart_path = os.path.join(tempfile.gettempdir(), "word_chart_%d.png" % int(time.time()))
        plt.savefig(chart_path)
        plt.close(fig)

        if (place or "end") == "cursor":
            try:
                rng = app.Selection.Range
            except Exception:
                rng = doc.Content
                rng.Collapse(0)
        else:
            rng = doc.Content
            rng.Collapse(0)
            rng.Text = "\r"
            rng.Collapse(0)

        rng.InlineShapes.AddPicture(FileName=chart_path, LinkToFile=False, SaveWithDocument=True)
        return {"ok": True, "action": "insert_chart", "doc": info,
                "report": "Inserted %s chart: '%s'." % (kind, title or "Chart")}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_chart")
        return {"ok": False, "action": "insert_chart", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_chart", "error": str(e)[:300]}


def insert_image(prompt, place="end"):
    """Generate and insert a free AI illustration or image into Word via Pollinations.ai."""
    import os
    import tempfile
    import requests
    try:
        app, doc = _bind()
        info = _doc_id(doc)
        prompt = (prompt or "").strip()
        if not prompt:
            return {"ok": False, "action": "insert_image", "error": "Image prompt cannot be empty."}

        encoded = requests.utils.quote(prompt)
        url = "https://image.pollinations.ai/prompt/%s?width=800&height=500&nologo=true" % encoded
        r = requests.get(url, timeout=30)
        if r.status_code != 200 or len(r.content) < 1000:
            return {"ok": False, "action": "insert_image",
                    "error": "Failed to generate image from public provider (status %s)." % r.status_code}

        img_path = os.path.join(tempfile.gettempdir(), "word_img_%d.jpg" % int(time.time()))
        with open(img_path, "wb") as f:
            f.write(r.content)

        if (place or "end") == "cursor":
            try:
                rng = app.Selection.Range
            except Exception:
                rng = doc.Content
                rng.Collapse(0)
        else:
            rng = doc.Content
            rng.Collapse(0)
            rng.Text = "\r"
            rng.Collapse(0)

        rng.InlineShapes.AddPicture(FileName=img_path, LinkToFile=False, SaveWithDocument=True)
        return {"ok": True, "action": "insert_image", "doc": info,
                "report": "Inserted image: '%s'." % prompt[:80]}
    except pywintypes.com_error as e:
        if _is_busy(e):
            return _busy_result("insert_image")
        return {"ok": False, "action": "insert_image", "error": str(e)[:300]}
    except Exception as e:
        return {"ok": False, "action": "insert_image", "error": str(e)[:300]}


TOOLS = {
    "add_footer": add_footer,
    "replace_text": replace_text,
    "delete_para_containing": delete_para_containing,
    "insert_at_end": insert_at_end,
    "set_full_text": set_full_text,
    "insert_at_cursor": insert_at_cursor,
    "format_text": format_text,
    "find_text": find_text,
    "select_text": select_text,
    "go_to_page": go_to_page,
    "delete_page": delete_page,
    "apply_heading": apply_heading,
    "insert_table": insert_table,
    "make_list": make_list,
    "align_paras": align_paras,
    "font_size": font_size,
    "font_color": font_color,
    "line_spacing": line_spacing,
    "set_orientation": set_orientation,
    "add_header": add_header,
    "add_page_numbers": add_page_numbers,
    "word_count": word_count,
    "export_pdf": export_pdf,
    "undo_last": undo_last,
    "toggle_track_changes": toggle_track_changes,
    "add_comment": add_comment,
    "insert_toc": insert_toc,
    "insert_data_table": insert_data_table,
    "add_watermark": add_watermark,
    "insert_chart": insert_chart,
    "insert_image": insert_image,
    "get_document_analytics": get_document_analytics,
}

