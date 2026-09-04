"""
Notes Overlay — a small always-on-top notes window that is excluded from
screen capture/recording on Windows 10 (2004+) / 11.

Standalone tool. Not part of any other project. Intended for keeping
personal presenter notes visible only to you while you present or share
your screen. Do not use this for exams, interviews, or any other
setting where using hidden notes would be considered cheating or
dishonest.

Usage:
    pip install -r requirements.txt
    python notes_overlay.py

Hotkeys (global, work even when the window isn't focused, via the
`keyboard` package):
    Ctrl+Alt+N       Show/hide the notes window
    Ctrl+Alt+H       Collapse/expand (roll the window up to its title bar)
    Ctrl+Alt+Up      Increase window opacity
    Ctrl+Alt+Down    Decrease window opacity
    Ctrl+Alt+=       Increase font size
    Ctrl+Alt+-       Decrease font size

In-window:
    Ctrl+F                 Toggle the quick find-in-notes bar
    Ctrl + mouse wheel     Zoom font size in/out
    Ctrl+V                 Paste (pastes an image if the clipboard has
                            one, otherwise pastes text as usual)
    Drag the title bar     Move the window
    Drag the "◢" corner    Resize the window
    Toolbar buttons        A- / A+ font size, insert image from file,
                            paste image from clipboard, import a
                            .txt/.docx/.pdf file, chat with AI, find in
                            notes, clear all

Notes (text + image references) are saved next to this script, in
"notes.json"; inserted images are copied into an "images" subfolder.

The "📥 Import" button reads a .txt/.md, .docx (Word) or .pdf file and
appends its text at the cursor. .txt and .docx use only the standard
library; .pdf needs the "pypdf" package (pip install pypdf).

The "🤖 Chat AI" button opens a separate always-on-top (and
capture-excluded) window that talks to the DeepSeek chat API. Provide a
key via the DEEPSEEK_API_KEY environment variable or the field in that
window; when entered there it is stored in "config.json" next to this
script. Only the standard library is used for the API call.

The chat window can also send images: use its "📎 Ảnh" button to pick
image files, or press Ctrl+V to attach an image from the clipboard
(clipboard paste needs Pillow). As soon as a message includes an image
the conversation switches to DeepSeek's vision model
("deepseek-v4-flash-vision-exp"); PNG/JPEG/GIF/WebP up to 32 MiB each are
accepted. Ctrl+Alt+H (or the "▁" button) rolls the chat window up to just
its top row, the same way it works for the notes window.
"""

import base64
import ctypes
import io
import json
import os
import sys
import threading
import urllib.error
import urllib.request
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox
from xml.etree import ElementTree as ET

try:
    import keyboard  # optional: enables global hotkeys
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False

try:
    from PIL import Image, ImageTk, ImageGrab  # optional: enables images
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NOTES_FILE = os.path.join(BASE_DIR, "notes.json")
LEGACY_NOTES_FILE = os.path.join(BASE_DIR, "notes.txt")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

# DeepSeek chat API (OpenAI-compatible)
DEEPSEEK_API_URL = "https://api.deepseek.com/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"                       # text only
DEEPSEEK_VISION_MODEL = "deepseek-v4-flash-vision-exp"  # accepts images
DEEPSEEK_SYSTEM_PROMPT = "Bạn là trợ lý AI hữu ích. Trả lời ngắn gọn, rõ ràng."

# DeepSeek vision limits: 32 MiB per image, ~48 MiB per request body.
MAX_IMAGE_BYTES = 32 * 1024 * 1024
MAX_ATTACH_BYTES_TOTAL = 30 * 1024 * 1024  # raw bytes; base64 inflates ~33%

# Windows constants for SetWindowDisplayAffinity
WDA_NONE = 0x00000000
WDA_EXCLUDEFROMCAPTURE = 0x00000011  # Windows 10 version 2004+

DEFAULT_FONT_SIZE = 11
MIN_FONT_SIZE = 7
MAX_FONT_SIZE = 32
MAX_IMAGE_WIDTH = 380  # inserted images are scaled down to this width


def set_capture_exclusion(hwnd, exclude=True):
    """Exclude (or re-include) this window from screen capture/recording.
    Requires Windows 10 build 19041+ / Windows 11. No-op elsewhere."""
    if sys.platform != "win32":
        return False
    affinity = WDA_EXCLUDEFROMCAPTURE if exclude else WDA_NONE
    try:
        return bool(ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, affinity))
    except Exception as e:
        print(f"[notes_overlay] Could not set capture exclusion: {e}")
        return False


# --------------------------------------------------- document import ----

_TEXT_EXTS = {".txt", ".md", ".markdown", ".log", ".csv", ".json", ".ini"}
_DOCX_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _read_text_file(path):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, "rb") as f:
        return f.read().decode("utf-8", "replace")


def _read_docx(path):
    """Extract plain text from a .docx without any third-party library.
    A .docx is a zip; the body text lives in word/document.xml."""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError):
        raise RuntimeError("File .docx không hợp lệ hoặc bị hỏng.")
    root = ET.fromstring(xml)
    paragraphs = []
    for p in root.iter(_DOCX_NS + "p"):
        parts = []
        for node in p.iter():
            if node.tag == _DOCX_NS + "t" and node.text:
                parts.append(node.text)
            elif node.tag == _DOCX_NS + "tab":
                parts.append("\t")
            elif node.tag in (_DOCX_NS + "br", _DOCX_NS + "cr"):
                parts.append("\n")
        paragraphs.append("".join(parts))
    return "\n".join(paragraphs)


def _read_pdf(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            raise RuntimeError(
                "Cần cài thư viện đọc PDF:\n    pip install pypdf")
    reader = PdfReader(path)
    return "\n\n".join((page.extract_text() or "") for page in reader.pages)


def extract_text_from_file(path):
    """Return the plain text of a .txt/.md/.docx/.pdf file. Raises
    RuntimeError with a user-facing message when it cannot."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return _read_docx(path)
    if ext == ".pdf":
        return _read_pdf(path)
    if ext == ".doc":
        raise RuntimeError("Định dạng .doc cũ không đọc được. "
                           "Hãy mở trong Word và lưu lại thành .docx.")
    if ext in _TEXT_EXTS or ext == "":
        return _read_text_file(path)
    # Unknown extension: give plain-text a try rather than refuse outright.
    return _read_text_file(path)


class NotesOverlay:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Notes")
        self.root.geometry("420x560+40+40")
        self.root.minsize(240, 180)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.92)
        self.root.configure(bg="#1e1e1e")
        self.root.overrideredirect(True)  # borderless, custom chrome below

        self.font_size = DEFAULT_FONT_SIZE
        self._image_refs = {}     # image name -> ImageTk.PhotoImage (keep alive!)
        self._image_sources = {}  # image name -> absolute file path on disk
        self._image_counter = 0
        self.chat = None          # ChatWindow instance, created on demand
        self._search_matches = [] # [(start_index, end_index), ...]
        self._search_idx = 0
        self._collapsed = False
        self._pre_collapse_geo = None
        self._chat_has_focus = False  # so the global collapse hotkey targets the right window

        self._build_ui()
        self._make_draggable(self.titlebar)
        self._make_resizable(self.grip)

        self._load_notes()

        self.root.update_idletasks()
        self._apply_capture_exclusion()

        if HAS_KEYBOARD:
            try:
                keyboard.add_hotkey("ctrl+alt+n", self.toggle_visibility)
                keyboard.add_hotkey("ctrl+alt+h", self._hotkey_collapse)
                keyboard.add_hotkey("ctrl+alt+up", lambda: self.change_opacity(0.05))
                keyboard.add_hotkey("ctrl+alt+down", lambda: self.change_opacity(-0.05))
                keyboard.add_hotkey("ctrl+alt+=", lambda: self.change_font_size(1))
                keyboard.add_hotkey("ctrl+alt+-", lambda: self.change_font_size(-1))
            except Exception as e:
                print(f"[notes_overlay] Global hotkeys unavailable: {e}")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------------------------------------------------- UI ----

    def _build_ui(self):
        self.titlebar = tk.Frame(self.root, bg="#2d2d2d", height=28)
        self.titlebar.pack(fill="x")
        tk.Label(self.titlebar, text="Notes", bg="#2d2d2d", fg="#aaaaaa",
                 font=("Segoe UI", 9)).pack(side="left", padx=8)
        tk.Button(self.titlebar, text="✕", bg="#2d2d2d", fg="#aaaaaa", bd=0,
                  activebackground="#3d3d3d", command=self.on_close).pack(side="right", padx=4)
        self.collapse_btn = tk.Button(self.titlebar, text="▁", bg="#2d2d2d", fg="#aaaaaa", bd=0,
                                      activebackground="#3d3d3d", command=self.toggle_collapse)
        self.collapse_btn.pack(side="right", padx=0)

        self.toolbar = tk.Frame(self.root, bg="#242424", height=26)
        self.toolbar.pack(fill="x")

        def tbtn(text, cmd):
            b = tk.Button(self.toolbar, text=text, bg="#242424", fg="#cccccc", bd=0,
                          activebackground="#3d3d3d", font=("Segoe UI", 9), command=cmd)
            b.pack(side="left", padx=2, pady=2)
            return b

        tbtn("A-", lambda: self.change_font_size(-1))
        tbtn("A+", lambda: self.change_font_size(1))
        tbtn("🖼 Chèn ảnh", self.insert_image_dialog)
        tbtn("📋 Dán ảnh", self.paste_image_from_clipboard)
        tbtn("📥 Import", self.import_document)
        tbtn("🤖 Chat AI", self.open_chat)
        tbtn("🔍 Tìm", self.toggle_search)
        tbtn("🗑 Xoá hết", self.clear_notes)

        # Quick find-in-notes bar (hidden until toggled with the button or Ctrl+F)
        self.search_bar = tk.Frame(self.root, bg="#242424")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._run_search())
        self.search_entry = tk.Entry(self.search_bar, textvariable=self.search_var,
                                     bg="#1e1e1e", fg="#e0e0e0", insertbackground="#ffffff",
                                     bd=0, font=("Segoe UI", 9))
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(8, 4), ipady=3)
        self.search_count = tk.Label(self.search_bar, text="", bg="#242424", fg="#888888",
                                     font=("Segoe UI", 8), width=7)
        self.search_count.pack(side="left")
        for label, cmd in (("▲", lambda: self._search_step(-1)),
                           ("▼", lambda: self._search_step(1)),
                           ("✕", self._hide_search)):
            tk.Button(self.search_bar, text=label, bg="#242424", fg="#cccccc", bd=0,
                      activebackground="#3d3d3d", font=("Segoe UI", 9),
                      command=cmd).pack(side="left", padx=1)
        self.search_entry.bind("<Return>", lambda e: self._search_step(1))
        self.search_entry.bind("<Shift-Return>", lambda e: self._search_step(-1))
        self.search_entry.bind("<Escape>", lambda e: self._hide_search())

        self.body = tk.Frame(self.root, bg="#1e1e1e")
        self.body.pack(fill="both", expand=True)

        self.scrollbar = tk.Scrollbar(self.body)
        self.scrollbar.pack(side="right", fill="y")

        self.text = tk.Text(self.body, bg="#1e1e1e", fg="#e0e0e0", insertbackground="#ffffff",
                             font=("Segoe UI", self.font_size), wrap="word", bd=0, padx=10, pady=10,
                             undo=True, yscrollcommand=self.scrollbar.set)
        self.text.pack(side="left", fill="both", expand=True)
        self.scrollbar.config(command=self.text.yview)

        self.text.bind("<KeyRelease>", self._on_change)
        self.text.bind("<Control-MouseWheel>", self._on_ctrl_wheel)         # Windows
        self.text.bind("<Control-Button-4>", lambda e: self.change_font_size(1))   # Linux
        self.text.bind("<Control-Button-5>", lambda e: self.change_font_size(-1))  # Linux
        self.text.bind("<Control-v>", self._on_ctrl_v)
        self.text.bind("<Control-f>", lambda e: self.toggle_search())
        self.root.bind("<Control-f>", lambda e: self.toggle_search())

        self.text.tag_configure("search_hit", background="#5a4b00")
        self.text.tag_configure("search_current", background="#c58900", foreground="#000000")

        # Resize grip, bottom-right corner (window has no OS border to drag)
        self.grip = tk.Label(self.root, text="◢", bg="#1e1e1e", fg="#555555",
                              cursor="bottom_right_corner", font=("Segoe UI", 10))
        self.grip.place(relx=1.0, rely=1.0, anchor="se")

    def _on_ctrl_wheel(self, event):
        self.change_font_size(1 if event.delta > 0 else -1)
        return "break"

    def _on_ctrl_v(self, event):
        # If the clipboard holds an image, insert it; otherwise let the
        # normal text-paste behaviour run.
        if HAS_PIL:
            try:
                img = ImageGrab.grabclipboard()
                if isinstance(img, Image.Image):
                    self._insert_pil_image(img)
                    return "break"
            except Exception:
                pass
        return None

    # ------------------------------------------------ quick find ----

    def toggle_search(self):
        if self.search_bar.winfo_ismapped():
            self._hide_search()
        else:
            self.search_bar.pack(fill="x", after=self.toolbar)
            self.search_entry.focus_set()
            self.search_entry.select_range(0, "end")
            self._run_search()
        return "break"

    def _hide_search(self, event=None):
        self.search_bar.pack_forget()
        self.text.tag_remove("search_hit", "1.0", "end")
        self.text.tag_remove("search_current", "1.0", "end")
        self._search_matches = []
        self.text.focus_set()
        return "break"

    def _run_search(self):
        self.text.tag_remove("search_hit", "1.0", "end")
        self.text.tag_remove("search_current", "1.0", "end")
        self._search_matches = []
        term = self.search_var.get()
        if term:
            start = "1.0"
            while True:
                pos = self.text.search(term, start, stopindex="end", nocase=True)
                if not pos:
                    break
                end = f"{pos}+{len(term)}c"
                self.text.tag_add("search_hit", pos, end)
                self._search_matches.append((pos, end))
                start = end
        if self._search_matches:
            insert = self.text.index("insert")
            self._search_idx = next(
                (i for i, (s, _e) in enumerate(self._search_matches)
                 if self.text.compare(s, ">=", insert)), 0)
            self._highlight_current()
        else:
            self._update_search_count()

    def _search_step(self, delta):
        if not self._search_matches:
            return "break"
        self._search_idx = (self._search_idx + delta) % len(self._search_matches)
        self._highlight_current()
        return "break"

    def _highlight_current(self):
        self.text.tag_remove("search_current", "1.0", "end")
        s, e = self._search_matches[self._search_idx]
        self.text.tag_add("search_current", s, e)
        self.text.see(s)
        self._update_search_count()

    def _update_search_count(self):
        n = len(self._search_matches)
        if n:
            self.search_count.configure(text=f"{self._search_idx + 1}/{n}")
        elif self.search_var.get():
            self.search_count.configure(text="0/0")
        else:
            self.search_count.configure(text="")

    # ------------------------------------------ dragging / resizing ----

    def _make_draggable(self, widget):
        widget.bind("<ButtonPress-1>", self._start_move)
        widget.bind("<B1-Motion>", self._do_move)

    def _start_move(self, event):
        self._drag_x = event.x
        self._drag_y = event.y

    def _do_move(self, event):
        x = self.root.winfo_pointerx() - self._drag_x
        y = self.root.winfo_pointery() - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    def _make_resizable(self, grip):
        grip.bind("<ButtonPress-1>", self._start_resize)
        grip.bind("<B1-Motion>", self._do_resize)

    def _start_resize(self, event):
        self._resize_start = (event.x_root, event.y_root,
                               self.root.winfo_width(), self.root.winfo_height())

    def _do_resize(self, event):
        x0, y0, w0, h0 = self._resize_start
        new_w = max(240, w0 + (event.x_root - x0))
        new_h = max(180, h0 + (event.y_root - y0))
        self.root.geometry(f"{new_w}x{new_h}")

    # ------------------------------------------------------ font size --

    def change_font_size(self, delta):
        self.font_size = max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, self.font_size + delta))
        self.text.configure(font=("Segoe UI", self.font_size))
        self._on_change()

    # ---------------------------------------------------------- images --

    def insert_image_dialog(self):
        if not HAS_PIL:
            messagebox.showwarning("Thiếu thư viện", "Cần cài Pillow để chèn ảnh:\npip install Pillow")
            return
        path = filedialog.askopenfilename(
            title="Chọn ảnh",
            filetypes=[("Ảnh", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("Tất cả file", "*.*")])
        if path:
            try:
                img = Image.open(path)
                self._insert_pil_image(img)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không mở được ảnh:\n{e}")

    def paste_image_from_clipboard(self):
        if not HAS_PIL:
            messagebox.showwarning("Thiếu thư viện", "Cần cài Pillow để dán ảnh:\npip install Pillow")
            return
        try:
            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                self._insert_pil_image(img)
            else:
                messagebox.showinfo("Không có ảnh", "Clipboard hiện không chứa ảnh.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không dán được ảnh:\n{e}")

    def _insert_pil_image(self, pil_img, max_width=MAX_IMAGE_WIDTH):
        if pil_img.mode not in ("RGB", "RGBA"):
            pil_img = pil_img.convert("RGBA")
        if pil_img.width > max_width:
            ratio = max_width / pil_img.width
            pil_img = pil_img.resize((max_width, max(1, int(pil_img.height * ratio))), Image.LANCZOS)

        self._image_counter += 1
        name = f"img_{self._image_counter}"
        photo = ImageTk.PhotoImage(pil_img)
        self._image_refs[name] = photo

        os.makedirs(IMAGES_DIR, exist_ok=True)
        save_path = os.path.join(IMAGES_DIR, f"{name}.png")
        pil_img.save(save_path, "PNG")
        self._image_sources[name] = save_path

        self.text.image_create(tk.INSERT, image=photo, name=name)
        self.text.insert(tk.INSERT, "\n")
        self._on_change()

    # ------------------------------------------------------- AI chat --

    def open_chat(self):
        if self.chat is not None and self.chat.alive():
            self.chat.lift()
            return
        self.chat = ChatWindow(self.root, self.exclude_from_capture, owner=self)

    # -------------------------------------------------- import file --

    def import_document(self):
        path = filedialog.askopenfilename(
            title="Import file văn bản",
            filetypes=[
                ("Văn bản (txt, Word, PDF)", "*.txt *.md *.docx *.pdf"),
                ("Word", "*.docx"),
                ("PDF", "*.pdf"),
                ("Text", "*.txt *.md *.log *.csv"),
                ("Tất cả file", "*.*"),
            ])
        if not path:
            return
        try:
            text = extract_text_from_file(path).strip()
        except Exception as e:
            messagebox.showerror("Không import được", str(e))
            return
        if not text:
            messagebox.showinfo("File trống",
                                "Không tìm thấy văn bản trong file này.")
            return

        name = os.path.basename(path)
        block = f"\n\n===== {name} =====\n{text}\n"
        if self.text.get("1.0", "end").strip() == "":
            block = block.lstrip("\n")
        self.text.insert("insert", block)
        self.text.see("insert")
        self._on_change()

    # ------------------------------------------------------ images --

    def clear_notes(self):
        if messagebox.askyesno("Xoá hết", "Xoá toàn bộ ghi chú (kể cả ảnh đã chèn)?"):
            self.text.delete("1.0", "end")
            self._image_refs.clear()
            self._image_sources.clear()
            self._on_change()

    # ------------------------------------------------------ save / load --

    def _on_change(self, event=None):
        if hasattr(self, "_save_after_id"):
            self.root.after_cancel(self._save_after_id)
        self._save_after_id = self.root.after(500, self._save_notes)

    def _save_notes(self):
        try:
            segments = []
            buf = []
            for key, value, _index in self.text.dump("1.0", "end", text=True, image=True):
                if key == "text":
                    buf.append(value)
                elif key == "image":
                    if buf:
                        segments.append({"type": "text", "value": "".join(buf)})
                        buf = []
                    src = self._image_sources.get(value)
                    if src:
                        segments.append({"type": "image", "file": os.path.relpath(src, BASE_DIR)})
            if buf:
                segments.append({"type": "text", "value": "".join(buf)})

            data = {"font_size": self.font_size, "segments": segments}
            with open(NOTES_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[notes_overlay] Could not save notes: {e}")

    def _load_notes(self):
        if os.path.exists(NOTES_FILE):
            try:
                with open(NOTES_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.font_size = data.get("font_size", DEFAULT_FONT_SIZE)
                self.text.configure(font=("Segoe UI", self.font_size))
                for seg in data.get("segments", []):
                    if seg["type"] == "text":
                        self.text.insert("end", seg["value"])
                    elif seg["type"] == "image" and HAS_PIL:
                        img_path = os.path.join(BASE_DIR, seg["file"])
                        if os.path.exists(img_path):
                            try:
                                pil_img = Image.open(img_path)
                                self._image_counter += 1
                                name = f"img_{self._image_counter}"
                                photo = ImageTk.PhotoImage(pil_img)
                                self._image_refs[name] = photo
                                self._image_sources[name] = img_path
                                self.text.image_create("end", image=photo, name=name)
                            except Exception as e:
                                print(f"[notes_overlay] Could not load image {img_path}: {e}")
            except Exception as e:
                print(f"[notes_overlay] Could not load notes.json: {e}")
        elif os.path.exists(LEGACY_NOTES_FILE):
            # migrate from the older plain-text-only format
            try:
                with open(LEGACY_NOTES_FILE, "r", encoding="utf-8") as f:
                    self.text.insert("1.0", f.read())
            except Exception as e:
                print(f"[notes_overlay] Could not load legacy notes.txt: {e}")

    # -------------------------------------------------------------- misc --

    def toggle_visibility(self):
        if self.root.state() == "withdrawn":
            self.root.deiconify()
        else:
            self.root.withdraw()

    def _hotkey_collapse(self):
        # Global Ctrl+Alt+H: collapse the chat window if it's the one in
        # use, otherwise collapse the notes window.
        if self._chat_has_focus and self.chat is not None and self.chat.alive():
            self.chat.toggle_collapse()
        else:
            self.toggle_collapse()

    def toggle_collapse(self):
        """Roll the window up to just its title bar (and back). Handy for
        parking it in a corner while keeping it grabbable."""
        x, y = self.root.winfo_x(), self.root.winfo_y()
        w = self.root.winfo_width()
        if self._collapsed:
            self.toolbar.pack(fill="x", after=self.titlebar)
            self.body.pack(fill="both", expand=True)
            self.grip.place(relx=1.0, rely=1.0, anchor="se")
            self.root.minsize(240, 180)
            geo = self._pre_collapse_geo or f"{w}x560+{x}+{y}"
            # keep the current on-screen position, restore only the size
            size = geo.split("+")[0]
            self.root.geometry(f"{size}+{x}+{y}")
            self.collapse_btn.configure(text="▁")
            self._collapsed = False
        else:
            self._pre_collapse_geo = self.root.geometry()
            if self.search_bar.winfo_ismapped():
                self._hide_search()
            self.body.pack_forget()
            self.toolbar.pack_forget()
            self.grip.place_forget()
            self.root.update_idletasks()
            h = max(self.titlebar.winfo_reqheight(), 26)
            self.root.minsize(240, h)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
            self.collapse_btn.configure(text="🗖")
            self._collapsed = True

    def change_opacity(self, delta):
        current = self.root.attributes("-alpha")
        new_val = min(1.0, max(0.2, current + delta))
        self.root.attributes("-alpha", new_val)

    def _apply_capture_exclusion(self):
        hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id()) if sys.platform == "win32" else None
        if hwnd:
            excluded = set_capture_exclusion(hwnd, True)
            if not excluded:
                print("[notes_overlay] Warning: capture exclusion not applied "
                      "(needs Windows 10 2004+ / Windows 11).")

    @staticmethod
    def exclude_from_capture(window):
        """Exclude any Tk window (e.g. the chat window) from screen capture."""
        if sys.platform != "win32":
            return
        try:
            window.update_idletasks()
            wid = window.winfo_id()
            hwnd = ctypes.windll.user32.GetParent(wid) or wid
            set_capture_exclusion(hwnd, True)
        except Exception as e:
            print(f"[notes_overlay] Could not exclude chat window from capture: {e}")

    def on_close(self):
        self._save_notes()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


class ChatWindow:
    """A small always-on-top window for chatting with DeepSeek's chat API.

    The API key is read from the DEEPSEEK_API_KEY environment variable, or
    entered in the window and stored in config.json next to this script.
    Requests run on a background thread so the UI never freezes.
    """

    def __init__(self, master, exclude_fn, owner=None):
        self.owner = owner  # NotesOverlay, for routing the global collapse hotkey
        self.win = tk.Toplevel(master)
        self.win.title("Chat AI — DeepSeek")
        self.win.geometry("460x600+480+40")
        self.win.minsize(320, 320)
        self.win.attributes("-topmost", True)
        self.win.configure(bg="#1e1e1e")

        self.messages = [{"role": "system", "content": DEEPSEEK_SYSTEM_PROMPT}]
        self._busy = False
        self.pending_images = []   # [{"name": str, "data_uri": str, "bytes": int}]
        self.uses_vision = False   # sticky once an image has been sent
        self._collapsed = False
        self._pre_collapse_geo = None

        self._build_ui()
        self._load_key()

        self.win.bind("<FocusIn>", self._on_focus_in)
        self.win.bind("<FocusOut>", self._on_focus_out)

        self.win.update_idletasks()
        exclude_fn(self.win)

    # ---------------------------------------------------------- UI ----

    def _build_ui(self):
        self.top = top = tk.Frame(self.win, bg="#242424")
        top.pack(fill="x")
        tk.Label(top, text="API key", bg="#242424", fg="#cccccc",
                 font=("Segoe UI", 9)).pack(side="left", padx=(8, 4), pady=6)
        self.key_var = tk.StringVar()
        self.key_entry = tk.Entry(top, textvariable=self.key_var, show="•",
                                  bg="#1e1e1e", fg="#e0e0e0", insertbackground="#ffffff",
                                  bd=0, font=("Segoe UI", 9))
        self.key_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(top, text="Lưu", bg="#242424", fg="#cccccc", bd=0,
                  activebackground="#3d3d3d", font=("Segoe UI", 9),
                  command=self._save_key).pack(side="left", padx=6)
        self.collapse_btn = tk.Button(top, text="▁", bg="#242424", fg="#cccccc", bd=0,
                                      activebackground="#3d3d3d", font=("Segoe UI", 9),
                                      command=self.toggle_collapse)
        self.collapse_btn.pack(side="left", padx=(0, 6))

        self.body = body = tk.Frame(self.win, bg="#1e1e1e")
        body.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(body)
        scrollbar.pack(side="right", fill="y")
        self.log = tk.Text(body, bg="#1e1e1e", fg="#e0e0e0", wrap="word", bd=0,
                           padx=10, pady=10, state="disabled", font=("Segoe UI", 10),
                           yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log.yview)
        self.log.tag_configure("you", foreground="#7db6ff", font=("Segoe UI", 10, "bold"))
        self.log.tag_configure("ai", foreground="#8ae28a", font=("Segoe UI", 10, "bold"))
        self.log.tag_configure("sys", foreground="#e0a35a", font=("Segoe UI", 9, "italic"))

        self.bottom = bottom = tk.Frame(self.win, bg="#242424")
        bottom.pack(fill="x")

        attach_row = tk.Frame(bottom, bg="#242424")
        attach_row.pack(fill="x")
        tk.Button(attach_row, text="📎 Ảnh", bg="#242424", fg="#cccccc", bd=0,
                  activebackground="#3d3d3d", font=("Segoe UI", 9),
                  command=self.attach_image).pack(side="left", padx=4, pady=(4, 0))
        self.attach_label = tk.Label(attach_row, text="", bg="#242424", fg="#e0a35a",
                                     font=("Segoe UI", 8), cursor="hand2")
        self.attach_label.pack(side="left", padx=6)
        self.attach_label.bind("<Button-1>", lambda e: self.clear_attachments())

        input_row = tk.Frame(bottom, bg="#242424")
        input_row.pack(fill="x")
        self.input = tk.Text(input_row, height=3, bg="#1e1e1e", fg="#e0e0e0",
                             insertbackground="#ffffff", wrap="word", bd=0,
                             padx=8, pady=6, font=("Segoe UI", 10))
        self.input.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        self.send_btn = tk.Button(input_row, text="Gửi", bg="#2d2d2d", fg="#e0e0e0", bd=0,
                                  activebackground="#3d3d3d", font=("Segoe UI", 10),
                                  command=self.send)
        self.send_btn.pack(side="right", padx=4, pady=4, fill="y")

        # Enter sends, Shift+Enter inserts a newline, Ctrl+V pastes an image
        self.input.bind("<Return>", self._on_return)
        self.input.bind("<Shift-Return>", lambda e: None)
        self.input.bind("<Control-v>", self._paste_image)

    def _on_return(self, event):
        self.send()
        return "break"

    # --------------------------------------------------- attachments --

    IMAGE_MAGIC = (
        (b"\x89PNG\r\n\x1a\n", "image/png"),
        (b"\xff\xd8\xff", "image/jpeg"),
        (b"GIF87a", "image/gif"),
        (b"GIF89a", "image/gif"),
    )

    @classmethod
    def _sniff_mime(cls, raw):
        for magic, mime in cls.IMAGE_MAGIC:
            if raw.startswith(magic):
                return mime
        if raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
            return "image/webp"
        return None

    def _add_attachment(self, raw, name):
        if len(raw) > MAX_IMAGE_BYTES:
            self._append("sys", "Hệ thống", f"Ảnh \"{name}\" lớn hơn 32 MiB, bỏ qua.")
            return
        mime = self._sniff_mime(raw)
        if mime is None:
            self._append("sys", "Hệ thống",
                         f"\"{name}\" không phải PNG/JPEG/GIF/WebP, bỏ qua.")
            return
        total = sum(im["bytes"] for im in self.pending_images) + len(raw)
        if total > MAX_ATTACH_BYTES_TOTAL:
            self._append("sys", "Hệ thống",
                         "Tổng dung lượng ảnh đính kèm quá lớn, bỏ qua ảnh này.")
            return
        b64 = base64.b64encode(raw).decode("ascii")
        self.pending_images.append({
            "name": name,
            "data_uri": f"data:{mime};base64,{b64}",
            "bytes": len(raw),
        })
        self._refresh_attach_label()

    def attach_image(self):
        paths = filedialog.askopenfilenames(
            title="Chọn ảnh gửi cho AI",
            filetypes=[("Ảnh", "*.png *.jpg *.jpeg *.gif *.webp"), ("Tất cả file", "*.*")])
        for path in paths:
            try:
                with open(path, "rb") as f:
                    raw = f.read()
            except Exception as e:
                self._append("sys", "Hệ thống", f"Không đọc được ảnh: {e}")
                continue
            self._add_attachment(raw, os.path.basename(path))

    def _paste_image(self, event=None):
        """Ctrl+V: attach an image from the clipboard, else fall through to
        the normal text paste."""
        if not HAS_PIL:
            return None
        try:
            img = ImageGrab.grabclipboard()
        except Exception:
            return None
        if not isinstance(img, Image.Image):
            return None
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        self._add_attachment(buf.getvalue(), "clipboard.png")
        return "break"

    def clear_attachments(self):
        if self.pending_images:
            self.pending_images = []
            self._refresh_attach_label()

    def _refresh_attach_label(self):
        n = len(self.pending_images)
        self.attach_label.configure(
            text=f"📎 {n} ảnh — bấm để bỏ" if n else "")

    # ------------------------------------------------------- helpers --

    def alive(self):
        try:
            return bool(self.win.winfo_exists())
        except tk.TclError:
            return False

    def lift(self):
        self.win.deiconify()
        self.win.lift()
        self.win.focus_force()

    def _on_focus_in(self, event):
        if self.owner is not None:
            self.owner._chat_has_focus = True

    def _on_focus_out(self, event):
        if self.owner is not None:
            self.owner._chat_has_focus = False

    def toggle_collapse(self):
        """Roll the chat window up to just the API-key row (and back),
        keeping its position. Ctrl+Alt+H or the "▁" button."""
        x, y = self.win.winfo_x(), self.win.winfo_y()
        w = self.win.winfo_width()
        if self._collapsed:
            self.body.pack(fill="both", expand=True)
            self.bottom.pack(fill="x")
            self.win.minsize(320, 320)
            size = (self._pre_collapse_geo or f"{w}x600").split("+")[0]
            self.win.geometry(f"{size}+{x}+{y}")
            self.collapse_btn.configure(text="▁")
            self._collapsed = False
        else:
            self._pre_collapse_geo = self.win.geometry()
            self.body.pack_forget()
            self.bottom.pack_forget()
            self.win.update_idletasks()
            h = max(self.top.winfo_reqheight(), 26)
            self.win.minsize(320, h)
            self.win.geometry(f"{w}x{h}+{x}+{y}")
            self.collapse_btn.configure(text="🗖")
            self._collapsed = True

    def _append(self, tag, label, text):
        self.log.configure(state="normal")
        if self.log.index("end-1c") != "1.0":
            self.log.insert("end", "\n")
        self.log.insert("end", f"{label}\n", tag)
        self.log.insert("end", text + "\n")
        self.log.configure(state="disabled")
        self.log.see("end")

    # --------------------------------------------------- API key I/O --

    def _load_key(self):
        key = os.environ.get("DEEPSEEK_API_KEY", "")
        if not key and os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    key = json.load(f).get("deepseek_api_key", "")
            except Exception as e:
                print(f"[notes_overlay] Could not read config.json: {e}")
        if key:
            self.key_var.set(key)
        else:
            self._append("sys", "Hệ thống",
                         "Chưa có API key. Dán key DeepSeek vào ô trên rồi bấm \"Lưu\".")

    def _save_key(self):
        key = self.key_var.get().strip()
        try:
            data = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            data["deepseek_api_key"] = key
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._append("sys", "Hệ thống", "Đã lưu API key vào config.json.")
        except Exception as e:
            self._append("sys", "Hệ thống", f"Không lưu được API key: {e}")

    # ----------------------------------------------------- messaging --

    def send(self):
        if self._busy:
            return
        prompt = self.input.get("1.0", "end").strip()
        if not prompt and not self.pending_images:
            return
        key = self.key_var.get().strip()
        if not key:
            self._append("sys", "Hệ thống", "Chưa có API key. Nhập key rồi bấm \"Lưu\".")
            return

        self.input.delete("1.0", "end")

        if self.pending_images:
            content = []
            if prompt:
                content.append({"type": "text", "text": prompt})
            for im in self.pending_images:
                content.append({"type": "image_url", "image_url": {"url": im["data_uri"]}})
            self.messages.append({"role": "user", "content": content})
            names = ", ".join(im["name"] for im in self.pending_images)
            shown = (prompt + "\n" if prompt else "") + f"[đã gửi {len(self.pending_images)} ảnh: {names}]"
            self._append("you", "Bạn", shown)
            self.pending_images = []
            self._refresh_attach_label()
            self.uses_vision = True
        else:
            self.messages.append({"role": "user", "content": prompt})
            self._append("you", "Bạn", prompt)

        self._busy = True
        self.send_btn.configure(state="disabled", text="…")
        threading.Thread(target=self._request, args=(key,), daemon=True).start()

    def _request(self, key):
        model = DEEPSEEK_VISION_MODEL if self.uses_vision else DEEPSEEK_MODEL
        payload = json.dumps({
            "model": model,
            "messages": self.messages,
            "stream": False,
        }).encode("utf-8")
        req = urllib.request.Request(
            DEEPSEEK_API_URL, data=payload, method="POST",
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {key}"})
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            reply = data["choices"][0]["message"]["content"].strip()
            self.messages.append({"role": "assistant", "content": reply})
            self._done("ai", "AI", reply)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            self.messages.pop()  # drop the user turn that failed
            self._done("sys", "Lỗi", f"HTTP {e.code}: {body}")
        except Exception as e:
            self.messages.pop()
            self._done("sys", "Lỗi", str(e))

    def _done(self, tag, label, text):
        def apply():
            if not self.alive():
                return
            self._busy = False
            self.send_btn.configure(state="normal", text="Gửi")
            self._append(tag, label, text)
        try:
            self.win.after(0, apply)
        except tk.TclError:
            pass


if __name__ == "__main__":
    if sys.platform != "win32":
        print("[notes_overlay] Note: capture exclusion only works on Windows 10 (2004+) / 11. "
              "The notes window will still work, but it will be visible in screen shares.")
    if not HAS_KEYBOARD:
        print("[notes_overlay] Tip: run 'pip install keyboard' to enable global hotkeys.")
    if not HAS_PIL:
        print("[notes_overlay] Tip: run 'pip install Pillow' to enable inserting/pasting images.")
    app = NotesOverlay()
    app.run()
