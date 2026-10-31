---
name: office-word-mcp
description: Create, read, and edit Microsoft Word (.docx) files via the Office-Word MCP server and smart_write helper. Use when the user mentions Word documents, DOCX read/write, live Word editing, reloading Word, tables, headings, or document automation.
license: MIT
compatibility: opencode
metadata:
  audience: developers
  format: docx
---

# SKILL: Office-Word MCP — Baca/Tulis Word Manual & Live

Skill ini mengajarkan agent cara membuat, membaca, dan mengedit dokumen Word (`.docx`)
lewat **Office-Word-MCP-Server** (protocol MCP) dan helper **`smart_write.py`**.

Repo: hasil fork dari `gongrzhe/office-word-mcp-server` + helper live-edit macOS.

---

## 0. Cara kerja (pahami dulu)

| Jalur | Mekanisme | Cocok untuk |
|-------|-----------|-------------|
| **MCP tools** (`word-document-server`) | `python-docx`: tulis langsung ke file di disk | Semua operasi, termasuk tabel & format |
| **Live mode** (`live_write.py` via AppleScript) | Sisip teks ke dokumen yang **sedang terbuka** di Word for Mac | Append paragraf/heading di ujung dokumen, terlihat mengetik |
| **smart_write.py** | Otomatis pilih jalur + jaga sinkronisasi | **Pakai ini sebagai default** |

> Batasan fundamental: `python-docx` tidak berkomunikasi dengan aplikasi Word.
> Word memuat dokumen ke memori — tulisan via file baru terlihat setelah **reload**.
> Live mode menulis ke memori Word langsung — terlihat seketika, tanpa tutup-buka.

---

## 1. Perintah utama — `smart_write.py`

```bash
export PYTHONPATH=/path/ke/repo-office
VENV=/path/ke/repo-office/.venv-mcp/bin/python

# Tulis live (kalau dokumen terbuka) atau via file (kalau tertutup)
$VENV smart_write.py append FILE.docx "teks paragraf" [--style "Heading 2"] [--delay 1.5] [--watch]
$VENV smart_write.py heading FILE.docx "Judul Bab" [--level 1] [--delay 1] [--watch]
$VENV smart_write.py pagebreak FILE.docx [--watch]

# Operasi file (otomatis save → tulis → reload bila dokumen terbuka)
$VENV smart_write.py table FILE.docx ROWS COLS '[["a","b"],["1","2"]]' [--watch]
$VENV smart_write.py replace FILE.docx "lama" "baru" [--watch]

# Baca (tidak mengubah apa pun)
$VENV smart_write.py read FILE.docx
$VENV smart_write.py info FILE.docx
```

Aturan sinkronisasi (jangan dilawan):

1. Habis **live-write** → `save` dulu sebelum operasi file.
2. Habis **operasi file** saat dokumen terbuka → **reload** dulu sebelum live-write.
3. `smart_write.py` sudah melakukan 1–2 otomatis. Jangan campur pemanggilan
   langsung ke MCP tools + AppleScript manual tanpa mengikuti urutan ini.

Kemampuan live mode: **append di ujung dokumen saja**, style per paragraf
(`Normal`, `Heading 1`–`Heading 9`), page break. Tabel, replace tengah,
format per-kata, dan hapus → lewat jalur file.

---

## 2. MCP server langsung (tanpa smart_write)

Berguna untuk operasi batch/skrip, atau saat Word tidak terlibat.

```bash
PYTHONPATH=/path/ke/repo-office $VENV -c "
import asyncio
from word_document_server.tools import document_tools, content_tools, format_tools
async def main():
    print(await document_tools.create_document('lap.docx', title='Judul', author='Nama'))
    print(await content_tools.add_heading('lap.docx', 'Pendahuluan', level=1))
    print(await content_tools.add_paragraph('lap.docx', 'Isi...'))
    print(await content_tools.add_table('lap.docx', 2, 2, [['A','B'],['1','2']]))
    print(await format_tools.highlight_table_header('lap.docx', 0))
    print(await document_tools.get_document_text('lap.docx'))
asyncio.run(main())
"
```

Semua tool MCP bersifat `async` — bungkus dengan `asyncio.run`.

---

## 3. Mendaftarkan MCP ke OpenCode

`install.sh` melakukannya otomatis. Manual:

```json
{
  "mcp": {
    "word-document-server": {
      "type": "local",
      "command": "/path/ke/repo-office/.venv-mcp/bin/python",
      "args": ["/path/ke/repo-office/word_mcp_server.py"],
      "env": {
        "PYTHONPATH": "/path/ke/repo-office",
        "MCP_TRANSPORT": "stdio"
      },
      "enabled": true
    }
  }
}
```

Restart OpenCode setelah mengubah config.

---

## 4. Anti-patterns

| ❌ Jangan | ✅ Lakukan |
|-----------|-----------|
| Harap tulisan file langsung terlihat di Word yang terbuka | `reload_word.py` / tutup-buka ulang / pakai live mode |
| Mengetik manual di Word saat live-write berjalan | Diam dulu — reload memakai `saving no`, ketikan belum-save hilang |
| Ganti isi tengah dokumen via live mode | `replace` (jalur file) |
| Klaim tabel/format aman tanpa cek | `info` → `table_count`, `read` → verifikasi teks |
| Commit `.venv-mcp/` ke git | Sudah di `.gitignore` — tiap device bikin venv sendiri via `install.sh` |
