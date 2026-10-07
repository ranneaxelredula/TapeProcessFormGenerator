# TapeProcessFormGenerator

Fills the IBAD-PF-009 Tape Process Form PDF from a Monday.com or RR124 Excel
export, with 1–3 pages per tape depending on its length. The form template,
`IBAD-PF-009-Rev00_Tape_Process_Form.pdf`, is included in this repository.

## Running it

- **As a program (.exe):** run `TapeProcessFormGenerator.exe` with
  `IBAD-PF-009-Rev00_Tape_Process_Form.pdf` in the same folder. No Python needed.
- **From source:** with Python 3 installed, double-click
  `TapeProcessFormGenerator.pyw` (template PDF in the same folder). Missing
  libraries are installed automatically on first start.

## Building the .exe

1. Install Python 3 from [python.org](https://www.python.org/downloads/) (only
   the PC that builds the .exe needs it).
2. Double-click `build_exe.bat`. The first run downloads PyInstaller and the
   program's libraries into `build\venv` and takes a few minutes.
3. The result is `dist\TapeProcessFormGenerator.exe`, with `logo.ico` as its
   icon. The template PDF is copied into `dist\` next to it.

Hand out the .exe together with the template PDF.
