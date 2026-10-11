# Land record document generator

A local Tkinter desktop application that fills the original Word document in `templates/`. Client data stays on your computer. The document engine uses python-docx and targeted XML edits; it updates all three repeated land schedules, including the nested schedule, and copies every other DOCX package part unchanged.

Read [TEMPLATE_INSPECTION.md](TEMPLATE_INSPECTION.md) for the actual fields, table structures, DevLys encoding, and fields that remain fixed. The engine was implemented and verified before connecting the desktop form.

## Mac setup

Use a Python 3.11 or newer installation with Tk support. The [official macOS installer](https://docs.python.org/3/using/mac.html) from python.org is a straightforward option. Tkinter is part of Python's standard library, but some distributors package it separately; it is not installed with pip. Verify it using [Python's Tkinter check](https://docs.python.org/3/library/tkinter.html):

```bash
cd "/Users/Varun/This Mac/Drive D/Replacer"
python3 -m tkinter
```

Close the demonstration window, then create a virtual environment using that same Python:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

If `python3` resolves to Apple's Command Line Tools installation or a Homebrew Python lacking `_tkinter`, choose the full path to your Python installation with Tk when creating the virtual environment. For Homebrew, install its `python-tk` formula matching your Python minor version. Recreate the virtual environment if you change interpreters. Run the commands in this workspace; `python -m landdocs` also opens the form.

**White window on macOS:** Apple's Command Line Tools Python 3.9 links to the obsolete Tk 8.5 supplied by macOS. A hidden-window integration check can pass even when visible controls fail to paint. This application now refuses Tk older than 8.6 on macOS before creating a window. See [Python's guidance on Mac Tcl/Tk](https://www.python.org/download/mac/tcltk/). Check your environment without opening a window:

```bash
.venv/bin/python -c 'import sys, tkinter; print(sys.version); print("Tk", tkinter.TkVersion)'
```

To replace an incompatible environment, preserve it and build a new one using an interpreter with modern Tk (substitute the actual Python path):

```bash
mv .venv .venv-old
/path/to/modern/python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

For this workspace, the environment was rebuilt using the locally available bundled Python 3.12.14 with Tk 9.0. The previous environment is retained in `.venv-apple-tk85/`. This `.venv` depends on that locally installed runtime; use the steps above with your own modern Python if the runtime is removed or the project is moved to another Mac.

Install legally obtained **DevLys 010, DevLys 240 and DevLys 160** fonts with macOS Font Book. The template does not embed them. Restart Word after installing fonts. Without these fonts the original Hindi may appear as Latin glyph codes. The desktop form uses DevLys 010 only. Font substitution can change line wrapping and pagination.

## Desktop workflow

1. Choose **English** or **हिन्दी** in the interface-language dropdown. Labels, buttons, record dialogs, validation, totals, status and window titles change immediately without changing client values or unsaved state. This is a session preference, independent of document encoding. macOS-owned file-picker and message-box button captions follow the system language; technical OS errors retain their original details.
2. Enter client name, father's name, **age**, **caste**, address, land village, Tehsil, District, **one Bahi number for the document**, and date. Age is whole years, 1–150; the date uses `YYYY-MM-DD`. Age and caste are required before saving or generating. Address and village are independent values.
3. Type already encoded DevLys keyboard codes. The form displays human-text fields in **DevLys 010** and provides one shared inline preview box. Focus a field and the box shows how its text renders in DevLys 010. There is no verified Unicode↔DevLys mapping in this project, so no conversion is performed and Unicode Hindi is not silently inserted into DevLys runs.
4. Add land records containing only **Khasra number, area in hectares, and Lagan**. Bahi is entered once above and reused in all prose references and the merged Bahi schedule column. Select a row and choose Edit, or double-click it. Select one or more rows and choose Remove. Identifiers retain leading zeros, slashes, punctuation and letter case. Area and Lagan must be non-negative decimal strings with a dot, up to 15 integer and 8 fractional digits; scientific notation, commas, NaN and infinity are rejected. At least one record is required to save or generate; the limit is 500 records.
5. Save JSON for future editing. Open JSON restores the complete client, document encoding and records. Unsaved changes are checked before New, Open, Load sample and Close. Save can replace an existing JSON after the native save dialog's confirmation. Cmd/Ctrl-S saves; Cmd/Ctrl-O opens.
6. Choose Generate new DOCX. The document is written automatically to `output/` as `<clientname>_data.docx`; unsafe filename characters are replaced and an incrementing suffix is used if that file already exists. Existing DOCX files are never overwritten. Generating does not save JSON automatically. The form reports that all three schedules were updated.

`sample_data/client.json` retains the template's original encoded client values and adds **illustrative age 40**, which is not a fact extracted from the template. The original caste glyph codes are retained. Choose Load sample to try it, then save your own copy.

Client JSON now uses `schema_version: 2`, with `age`, `caste` and `bahi` at the client level; records have only `khasra`, `area` and `lagan`. Every identifier and numeric input is stored as a string. Totals are recomputed with `Decimal`, not persisted or trusted from JSON. Version 1 files can still be opened when all records have the same Bahi. Age and caste open blank and must be filled before saving/generating; the importer invents neither value. Older files with conflicting Bahi numbers are rejected with an explanation.

## Command line

The same engine can be run without the desktop:

```bash
python -m landdocs --data sample_data/client.json --output output/my-client.docx
```

If `--output` is omitted, the generated file uses `output/<clientname>_data.docx`. Choose an explicit output path when needed. The updated generated example is in `output/sample-client-v2.docx`.

## Editing behavior

The original is read afresh on each generation and checked against its inspected SHA-256. Edited, substituted or newly added templates require a new inspection/profile; the application deliberately refuses them. It uses the one original in `templates/`, not the copies in `Docs/`.

Names are updated at all 13 repeated identity mentions. Age fills both existing affidavit blanks; caste updates all ten references, including the separate caste line. Address updates the nine residence occurrences; village updates land references. Client Tehsil and District references are separate from bank, advocate, execution-city and registrar references. Six document-date slots, two prose area summaries and four prose Bahi references are updated. Historical dates and record years remain fixed. New strings are never rescanned for replacements, so a name containing an old template token is safe.

Every matching table is found recursively, preserving its header, grid, widths, borders and source formatting. Rows are cloned from the existing data row and the total row. The two four-column tables keep their columns; the five-column table uses the document-level Bahi with its original vertical merge retained. The same Bahi replaces all four prose references. Area and Lagan totals use exact decimal addition, and identifiers never pass through integer/float conversion. The number of rows and longer text can change wrapping or page count; no font shrinking is performed.

The five-column table's original `SUM(ABOVE)` fields are materialized as static Decimal totals, preserving their font properties, to prevent Word or a renderer from recalculating/rounding them. Mixed-font total labels keep their Hindi suffix runs; new dates use isolated Times New Roman runs for literal punctuation. Age preserves the source DevLys run.

Fixed legal assertions, patwari number, bank branch, advocate details, valuation amounts and document metadata are retained. Review these and the rendered output for the intended client before sharing. The application does not perform land valuation or revise the legal declarations.

## Verification

No extra test framework is needed:

```bash
python -m unittest discover -s tests -v
python -m scripts.verify_gui
```

The first command checks exact Decimal/JSON round trips, input rejection, leading zeros and slashes, all repeated and deeply nested schedules, row growth/shrinkage, single merged Bahi, old JSON migration/conflict rejection, age/caste/identity/address/location/date scope, split-run replacements, explicit document fonts, original-template immutability, create-only DOCX output, and byte-for-byte preservation of every other ZIP member. It also compares table headers/grids, section geometry and untouched paragraph XML.

The GUI check needs an active desktop and working Tk. It creates a hidden window, uses temporary files and exercises add/edit/remove, totals, the DevLys preview, Hindi/English switching with data/selection preservation, translated record dialogs, JSON save/load, actual DOCX generation, validation and unsaved-change handling. File/message dialogs are mocked so it needs no manual clicks. This verifies widget behavior, not visible painting: also launch `python app.py` and confirm the labels, inputs, preview and table are visible.

For visual QA, open the original and generated DOCX in Word with the original fonts installed. Check every page, all three schedules, date placement, line wrapping and the merged declaration schedule. Automated XML checks cannot establish the intended spelling of manually encoded DevLys or guarantee Word pagination. `VERIFICATION.md` records the checks performed during implementation.

## Source layout

- `landdocs/model.py`: validation, JSON persistence, Decimal totals, atomic file writes.
- `landdocs/engine.py`: inspected-template profile, run-aware field edits, recursive land schedules, DOCX package preservation.
- `landdocs/gui.py`: Tkinter form and record editor.
- `landdocs/i18n.py`: Hindi/English interface and validation translations.
- `landdocs/__main__.py` / `app.py`: GUI and CLI entry points.
- `sample_data/client.json`: original encoded template values with illustrative age.
- `tests/test_engine.py` / `scripts/verify_gui.py`: automated checks.

Runtime dependency: `python-docx` (which installs lxml). Tkinter, JSON, Decimal and the verification framework use Python's standard library. Word is not required to generate documents.

## Build a standalone application

The project includes a PyInstaller configuration in `packaging/landdocs.spec`. Build on the operating system where the application will run; a Windows `.exe` must be built on Windows, and a macOS `.app` must be built on macOS. PyInstaller does not produce one portable executable for both systems.

On macOS:

```bash
python3 -m venv .build-venv
.build-venv/bin/python -m pip install -r requirements.txt -r requirements-build.txt
.build-venv/bin/pyinstaller --noconfirm --clean packaging/landdocs.spec
open dist/LandDocs.app
```

On Windows PowerShell:

```powershell
py -m venv .build-venv
.build-venv\Scripts\python -m pip install -r requirements.txt -r requirements-build.txt
.build-venv\Scripts\pyinstaller --noconfirm --clean packaging\landdocs.spec
Start-Process dist\LandDocs\LandDocs.exe
```

The build includes the inspected template and sample JSON. The generated `output` folder is placed beside the executable, so the user can choose a save location normally. Install the legally licensed DevLys fonts on the target computer; fonts are not bundled.

### Build the Windows EXE from a Mac

PyInstaller does not cross-compile a native Windows executable from macOS. The repository includes `.github/workflows/build-windows.yml`, which runs the build on a Windows runner. From your Mac:

1. Create a GitHub repository and push this project, including the `.github` folder.
2. Open the repository's **Actions** tab and select **Build Windows application**.
3. Choose **Run workflow** and wait for the job to finish.
4. Open the completed run, download the **LandDocs-Windows** artifact, and unzip it.
5. Copy `LandDocs.exe` to the Windows PC and double-click it. Install the DevLys fonts on that PC first.

The Windows package includes the template and sample data. It creates an `output` folder beside the executable. Windows SmartScreen may show an unsigned-app warning because this build is not code-signed.
