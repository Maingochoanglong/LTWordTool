# LTWordTool

[🇻🇳 Tiếng Việt](README.md) | 🇬🇧 English

![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

**Desktop tool for processing Word (.docx) files containing MathType equations** — format-preserving find & replace, plus automatic fixing of auto-expanding parentheses in MathType equations.

Built for office workers, teachers, and editors working with Vietnamese-language documents that contain math formulas.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [License](#license)

---

## Features

- **Format-preserving find & replace** — search & replace one or more content pairs in a `.docx` file, keeping formatting (bold/italic/color/font...) intact, with automatic format propagation in both directions up to an optional stop string.
- **MathType parentheses fixing** — automatically converts auto-expanding parentheses/brackets around simple text or numbers in MathType equations into hard (fixed) ones, correctly handling multiple levels of nesting.
- **Automatic MathType preview refresh** (requires Windows + Word + MathType) — after fixing parentheses, the tool silently opens Word so MathType redraws the correct preview image, without manually double-clicking each equation.
- Simple GUI, no command line needed — pick a file, check the desired actions, click run.

## Requirements

| Component    | Required | Notes                                                                    |
|---------------|:--------:|----------------------------------------------------------------------------|
| Python 3.10+  | Yes      | Only needed if running from source                                        |
| PySide6       | Yes      | GUI                                                                         |
| lxml          | Yes      | Reading/writing XML inside `.docx`                                        |
| olefile       | Yes      | Reading OLE/CFB structure                                                  |
| pywin32       | No       | Only needed for the preview-refresh step (Windows + Word + MathType installed) |
| Windows       | No       | Required specifically for the preview-refresh step; other operations run on any OS |

## Installation

### Option 1 — Download the prebuilt binary (recommended for most users)

Download the latest `.exe` from [Releases](https://github.com/Maingochoanglong/LTWordTool/releases/latest) — no Python needed.

### Option 2 — Run from source

```bash
git clone https://github.com/Maingochoanglong/LTWordTool.git
cd LTWordTool
pip install -r requirements.txt
python gui.py
```

On Windows, install `pywin32` as well if you want the preview-refresh feature:

```bash
pip install pywin32
```

## Usage

1. Open the app, choose your **SOURCE FILE** (the `.docx` to process).
2. Check one or both actions:
   - **Replace content** — add (find file, replacement file) pairs to the list.
   - **Fix MathType parentheses** — no extra configuration needed.
3. Choose where to save the **OUTPUT FILE** (or leave blank to auto-name it).
4. Click **Run**, and follow progress in the log panel.

If both actions are checked, the order is always fixed: replace content first, fix MathType parentheses second.

## Project Structure

| File                                    | Role                                                              |
|-------------------------------------------|---------------------------------------------------------------------|
| `gui.py`                                | Qt6 (PySide6) GUI, wires the actions together                       |
| `panel.py`                              | Shared UI framework (background thread, buttons, log, saved settings) |
| `pipeline.py`                           | Chains the processing steps sequentially                            |
| `replace_docx.py`                       | Format-preserving content replacement in `.docx`                    |
| `fix_mathtype_parens.py`                | Orchestration for MathType parentheses fixing                       |
| `mtef_parser.py` / `mtef_transform.py`  | Reads & transforms the binary MTEF structure                        |
| `cfb_builder.py`                        | Repacks the OLE/CFB container                                       |
| `mathtype_refresh.py`                   | Automatic preview refresh via Word COM                              |

## License

Source code is released under **[AGPL-3.0](LICENSE)** — free to use, modify, and run.

Want to use it differently (commercially, embedded in a product, redistributed outside AGPL-3.0 terms)? Contact: **<your-email>**

## Bugs / Contributing

Found a bug, have feedback, or a question about using it?

1. **Check first** in [existing Issues](https://github.com/Maingochoanglong/LTWordTool/issues?q=is%3Aissue) — it may already be reported or answered.
2. Not there? **Open a [new Issue](https://github.com/Maingochoanglong/LTWordTool/issues/new)**, describing clearly: what you were doing, what you expected, what actually happened (attach a sample `.docx` if convenient, with any sensitive content removed).

Please use Issues rather than direct messages — it helps other users with the same problem find an answer, and doesn't rely on a single point of contact.

Pull requests improving the project are always welcome.
