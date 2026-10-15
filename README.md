# TapeProcessFormGenerator

Fills the IBAD-PF-009 Tape Process Form from a Monday.com or RR124 Excel
export and saves all forms as one Word file (1–3 pages per tape, depending on
its length), so they can still be edited in Word. The Word template,
`IBAD-PF-009_Tape Process Form_Rev.0.docx`, is included in this repository.

Each value replaces its blank in the form. A value missing from the export
leaves the blank as printed, for hand entry.

## Running it

- **As a program (.exe):** run `TapeProcessFormGenerator.exe` with
  `IBAD-PF-009_Tape Process Form_Rev.0.docx` in the same folder. No Python
  needed.
- **From source:** with Python 3 installed, double-click
  `TapeProcessFormGenerator.pyw` (Word template in the same folder). Missing
  libraries are installed automatically on first start.

The program finds each blank by the form's wording around it. If the form is
revised so that wording changes, it stops with a message naming the blank it
couldn't find, rather than fill the wrong place.

## Building the .exe

1. Install Python 3 from [python.org](https://www.python.org/downloads/) (only
   the PC that builds the .exe needs it).
2. Double-click `build_exe.bat`. The first run downloads PyInstaller and the
   program's libraries into `build\venv` and takes a few minutes.
3. The result is `dist\TapeProcessFormGenerator.exe`, with `logo.ico` as its
   icon. The Word template is copied into `dist\` next to it.

Hand out the .exe together with the Word template.
