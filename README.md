# Conversor Milheiro → Unidade

## Why does this app exist?

This app exists solely to fulfill my father-in-law's recurring need. Every month, he asks me to convert a few columns in his spreadsheets from **"milheiro" (per-thousand pricing)** to **unit pricing**. And well, being the good son-in-law that I am, I decided to do a little more than just that. 😄

## What it does

A web-based tool that converts wholesale per-thousand prices to individual unit prices in Excel/CSV spreadsheets, while **preserving all original formatting** — borders, fonts, colors, column widths, merged cells, headers, everything.

### Features

- ✅ Auto-detects the actual header row (skips title/subtitle lines)
- ✅ Converts text-formatted values (`"R$ 1.234,56"`, `"850,00"`) to numbers
- ✅ Preserves all Excel formatting (borders, fonts, colors, column widths, merged cells)
- ✅ Shows sample values for each column so you know what you're converting
- ✅ Configurable divisor (default 1000)
- ✅ Supports `.xlsx`, `.xls`, `.csv`, `.ods`
- ✅ Can be packaged as a standalone `.exe` — no Python needed for end users

## Project Structure

```
conversor-milheiro/
├── server.py              # Flask backend API
├── build_exe.py           # Script to generate the .exe
├── static/
│   └── index.html         # Frontend (HTML/CSS/JS)
├── GERAR_EXE.bat          # ⭐ Double-click → generates the .exe automatically
├── EXECUTAR.bat            # Runs directly without generating .exe
├── .gitignore
└── README.md
```

## Getting Started

### Option 1: Generate a standalone `.exe` (recommended for distribution)

1. Make sure **Python 3.10+** is installed (with "Add to PATH" checked)
2. Double-click `GERAR_EXE.bat`
3. Wait 2–5 minutes
4. `ConversorMilheiro.exe` will be created in the project folder
5. Distribute the `.exe` — it works without Python installed

### Option 2: Run directly (development)

1. Double-click `EXECUTAR.bat`
2. The browser will open automatically at `http://localhost:5000`

### Option 3: Command line

```bash
pip install flask pandas openpyxl xlrd
python server.py
```

## How to Use

1. **Upload** — Drag & drop or select your spreadsheet
2. **Select** — Check the columns with per-thousand prices (only numeric/convertible columns are selectable)
3. **Convert** — Click "Converter Selecionadas"
4. **Download** — Download the converted spreadsheet with all formatting intact

## Tech Stack

- **Backend:** Python 3, Flask, Pandas, OpenPyXL
- **Frontend:** HTML5, CSS3 (dark theme), Vanilla JavaScript
- **Packaging:** PyInstaller

## License

MIT