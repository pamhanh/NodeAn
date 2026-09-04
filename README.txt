Notes Overlay
=============

A tiny, standalone notes window for keeping personal notes visible only to
you while you present or share your screen. It is not part of any other
project. The notes side works fully offline; the optional "Chat AI" window
calls the DeepSeek API, which needs your own API key and is billed by
DeepSeek per use.

How it works
------------
On Windows 10 (version 2004 / build 19041 or later) and Windows 11, this
uses the OS-level SetWindowDisplayAffinity API to exclude the window from
screen capture and recording (Zoom, Teams, Google Meet, OBS, etc.) while
it stays fully visible to you on your own screen. This is the same
mechanism legitimate teleprompter and "presenter notes" apps use.

Setup
-----
1. Install Python 3.8+ if you don't already have it.
2. Double-click run.bat (it installs the required libraries automatically
   and starts the app), or manually:
       pip install -r requirements.txt
       python notes_overlay.py

Features
--------
- Auto-saving text notes (saved to notes.json next to the script).
- Adjustable font size:
    * Toolbar "A-" / "A+" buttons
    * Ctrl + mouse wheel over the text
    * Global hotkeys Ctrl+Alt+= / Ctrl+Alt+-
- Insert images:
    * "🖼 Chèn ảnh" button — pick an image file
    * "📋 Dán ảnh" button, or just Ctrl+V — paste an image from the
      clipboard (e.g. a screenshot you just copied)
    * Images are auto-scaled and saved into an "images" subfolder, and
      reloaded the next time you open the app
- Chat with AI ("🤖 Chat AI" button):
    * Opens a separate always-on-top window that is also hidden from
      screen capture, for asking a DeepSeek model questions while you work
    * Needs a DeepSeek API key. Either set the DEEPSEEK_API_KEY
      environment variable, or paste the key into the field at the top of
      the chat window and press "Lưu" (it is saved to config.json next to
      the script)
    * Enter sends a message; Shift+Enter adds a newline. Requests run in
      the background so the notes window never freezes
    * Send images: the "📎 Ảnh" button picks image files, or press Ctrl+V
      to attach an image from the clipboard (clipboard paste needs
      Pillow). Click the "📎 N ảnh" label to drop the attachments.
      PNG/JPEG/GIF/WebP up to 32 MiB each are accepted
    * When a message includes an image the chat automatically switches to
      DeepSeek's vision model (deepseek-v4-flash-vision-exp), which is
      billed separately; text-only chats use deepseek-chat
    * Uses only the Python standard library — no extra install needed
    * Ctrl+Alt+H (or the "▁" button) rolls the chat window up to just its
      top row, and back — same as the notes window
- Import a document ("📥 Import" button):
    * Reads a .txt / .md, .docx (Word) or .pdf file and appends its text
      into the notes at the cursor, under a "===== filename =====" header
    * .txt and .docx work with just the standard library; .pdf needs the
      "pypdf" package (included in requirements.txt). Old .doc files must
      be re-saved as .docx first
    * Only text is imported, not images or formatting
- Quick find in notes ("🔍 Tìm" button or Ctrl+F):
    * A small search bar appears above the text; type to highlight every
      match as you go
    * Enter jumps to the next match, Shift+Enter to the previous one; the
      "1/5" counter shows the position. Use the ▲ / ▼ buttons too
    * Esc or the "✕" button closes the bar and clears the highlights
- Resizable window: drag the "◢" handle in the bottom-right corner
- Movable window: drag the title bar
- "🗑 Xoá hết" button clears all notes and images (asks for confirmation)
- Adjustable transparency: Ctrl+Alt+Up / Ctrl+Alt+Down
- Show/hide: Ctrl+Alt+N (works even when the window isn't focused)
- Collapse/expand: Ctrl+Alt+H, or the "▁" button in the title bar --
  rolls the window up to just its title bar so it takes almost no space,
  press again to restore it to its previous size

All global hotkeys need the optional "keyboard" package (included in
requirements.txt); image support needs "Pillow" and PDF import needs
"pypdf" (both also included).

A note on appropriate use
--------------------------
This tool is meant for your own presenter notes, cue cards, or reminders
during a talk, demo, or livestream -- similar to a teleprompter or
PowerPoint's Presenter View. Please don't use it in a context where having
hidden notes would count as dishonest or against the rules, such as a
proctored exam or a job interview.
