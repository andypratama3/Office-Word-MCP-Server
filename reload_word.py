#!/usr/bin/env python3
"""Reload dokumen Word yang sedang terbuka agar menampilkan versi terbaru dari disk.

Cara kerja: tutup dokumen TANPA menyimpan (agar versi memori yang basi dibuang),
lalu buka ulang file dari disk. Panggil skrip ini setiap selesai menulis via MCP.
"""
import subprocess
import sys

WORD_FILE = sys.argv[1] if len(sys.argv) > 1 else "/Users/andypratama3/Downloads/Doc1.docx"
DOC_NAME = WORD_FILE.rsplit("/", 1)[-1]

SCRIPT = f'''
tell application "Microsoft Word"
    try
        close document "{DOC_NAME}" saving no
    end try
    open POSIX file "{WORD_FILE}"
    activate
end tell
'''

r = subprocess.run(["osascript", "-e", SCRIPT], capture_output=True, text=True, timeout=30)
if r.returncode != 0:
    print("RELOAD GAGAL:", r.stderr.strip())
    sys.exit(1)
print(f"Reloaded: {DOC_NAME}")
