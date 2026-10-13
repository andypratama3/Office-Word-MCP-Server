#!/usr/bin/env python3
"""smart_write.py — satu pintu tulis Word: otomatis pilih mode terbaik.

- Dokumen TERBUKA di Word -> mode LIVE (AppleScript, tanpa tutup-buka).
- Dokumen TERTUTUP        -> mode FILE (python-docx via MCP tools).
- Operasi yang tidak bisa live (tabel, replace) saat dokumen terbuka:
  save dulu -> tulis file -> reload (aman, tidak ada yang hilang).
- --watch: bukakan dokumen di Word dulu agar bisa ditonton.

Pakai:
  smart_write.py append FILE "teks" [--style "Heading 2"] [--delay 1.5] [--watch]
  smart_write.py heading FILE "teks" [--level 1] [--delay 1.5] [--watch]
  smart_write.py pagebreak FILE [--watch]
  smart_write.py table FILE ROWS COLS '[["a","b"],["1","2"]]' [--watch]
  smart_write.py replace FILE "lama" "baru" [--watch]
  smart_write.py read FILE
  smart_write.py info FILE
"""
import asyncio
import json
import subprocess
import sys
import time

OFFICE_DIR = "/Users/andypratama3/Development/tool_new/office"
sys.path.insert(0, OFFICE_DIR)


def osa(script: str) -> str:
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=30)
    if r.returncode != 0:
        raise RuntimeError("AppleScript: " + r.stderr.strip())
    return r.stdout.strip()


def esc(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"')


def open_docs() -> list:
    try:
        out = osa('tell application "Microsoft Word" to get name of every document')
    except RuntimeError:
        return []
    if not out or out == "missing value":
        return []
    return [d.strip() for d in out.split(",")]


def is_open(word_file: str) -> bool:
    return word_file.rsplit("/", 1)[-1] in open_docs()


def front(word_file: str):
    osa(f'tell application "Microsoft Word" to open POSIX file "{word_file}"')


def word_save():
    osa('tell application "Microsoft Word" to save active document')


def word_reload(word_file: str):
    doc = word_file.rsplit("/", 1)[-1]
    osa(f'''tell application "Microsoft Word"
    try
        close document "{esc(doc)}" saving no
    end try
    open POSIX file "{word_file}"
    activate
end tell''')


def live_append(lines: list, style: str, delay: float):
    for line in lines:
        osa(f'''tell application "Microsoft Word"
    tell active document
        set n to count characters of text object of active document
        set endPos to collapse range (create range start 0 end n) direction collapse end
        set content of endPos to ("{esc(line)}" & (ASCII character 13))
        set style of last paragraph to "{esc(style)}"
    end tell
end tell''')
        time.sleep(delay)
    word_save()


def live_pagebreak():
    osa('''tell application "Microsoft Word"
    tell active document
        set n to count characters of text object of active document
        set endPos to collapse range (create range start 0 end n) direction collapse end
        set content of endPos to (ASCII character 12)
    end tell
end tell''')
    word_save()


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    cmd, word_file = sys.argv[1], sys.argv[2]
    rest = sys.argv[3:]
    watch = "--watch" in rest
    rest = [a for a in rest if a != "--watch"]

    def opt(name, default):
        if name in rest:
            return rest[rest.index(name) + 1]
        return default

    if watch and not is_open(word_file):
        front(word_file)
        time.sleep(2)

    opened = is_open(word_file)
    if opened:
        front(word_file)  # pastikan jendela target yang di depan

    from word_document_server.tools import document_tools, content_tools

    async def file_op(coro, need_reload=True):
        if opened:
            word_save()  # flush ketikan user dulu, aman
        res = await coro
        print(res)
        if opened and need_reload:
            word_reload(word_file)
            print(f"[reload] {word_file}")
        return res

    async def run():
        if cmd == "append":
            style = opt("--style", "Normal")
            delay = float(opt("--delay", "1.5"))
            lines = [a for a in rest if not a.startswith("--")]
            # buang nilai opsi dari lines
            clean = []
            skip = False
            for i, a in enumerate(rest):
                if skip:
                    skip = False
                    continue
                if a in ("--style", "--delay"):
                    skip = True
                    continue
                clean.append(a)
            if opened:
                live_append(clean, style, delay)
                print(f"[live] {len(clean)} paragraf -> {word_file}")
            else:
                for ln in clean:
                    await file_op(content_tools.add_paragraph(word_file, ln), need_reload=False)
        elif cmd == "heading":
            level = int(opt("--level", "1"))
            delay = float(opt("--delay", "1.0"))
            text = rest[0]
            if opened:
                live_append([text], f"Heading {level}", delay)
                print(f"[live] heading -> {word_file}")
            else:
                await file_op(content_tools.add_heading(word_file, text, level=level),
                              need_reload=False)
        elif cmd == "pagebreak":
            if opened:
                live_pagebreak()
                print(f"[live] page break -> {word_file}")
            else:
                await file_op(content_tools.add_page_break(word_file), need_reload=False)
        elif cmd == "table":
            rows, cols, data = int(rest[0]), int(rest[1]), json.loads(rest[2])
            await file_op(content_tools.add_table(word_file, rows, cols, data))
            print(f"[file+reload] tabel {rows}x{cols} -> {word_file}")
        elif cmd == "replace":
            find, repl = rest[0], rest[1]
            await file_op(content_tools.search_and_replace(word_file, find, repl))
            print(f"[file+reload] replace '{find}' -> {word_file}")
        elif cmd == "read":
            print(await document_tools.get_document_text(word_file))
        elif cmd == "info":
            print(await document_tools.get_document_info(word_file))
        else:
            print("Perintah tak dikenal:", cmd)
            sys.exit(1)

    asyncio.run(run())


if __name__ == "__main__":
    main()
