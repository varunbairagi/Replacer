# Verification performed

Checked on this Mac on 11 October 2026.

## Current age, caste, Bahi and language update

- **19 automated model/engine/runtime checks pass** under Python 3.12.14 with python-docx 1.2.0. New checks cover both age slots, all ten caste references, the same document-level Bahi in four prose slots and the merged table, age/caste/font validation, schema 2 string persistence, schema 1 migration and rejection of conflicting older Bahi numbers. Missing age/caste from older JSON are never invented and block save/generation until completed.
- **Tk 9.0 integration passes** with real Tk widgets and real DOCX generation: add/edit/remove now uses three record fields; Bahi is a client field. It checks both font dropdown choices, Hindi/English labels/titles/status/totals, an actual Hindi record dialog and its validation, and preservation of data, selected rows and unsaved state when switching languages. Native dialogs are mocked. Visible desktop painting is still not established by this hidden-window check.
- Generated **`output/sample-client-v2.docx`** from the updated sample, with illustrative age 40. Both age blanks are filled and the original caste glyph codes retained. Original template fingerprint and every other ZIP part remain unchanged.
- Rendered the updated sample with the same explicit DevLys/font configuration. It remains **12 pages**. Inspected page 1 with both age insertions and confirmed pages 2–12 are pixel-identical to the previously reviewed final sample. An initial attempt to set ages in Times New Roman changed line metrics and pagination; the final legacy implementation keeps those digits in the original DevLys run, with correct visible digits and the original pagination.
- `sample_data/client_mangal.json` is a separate fictional Hindi Unicode example, not an asserted conversion of the template's encoded client. Mangal was not found in this Mac's installed font inventory. The engine's XML checks confirm new Unicode values are explicitly tagged Mangal and never placed in DevLys runs; exact Mangal rendering requires installing that font. No Unicode-to-DevLys converter or guessed mapping was added.
- Application-owned labels, buttons, messages and validation are bilingual. macOS-owned file-picker/message-box button captions follow macOS; OS-level technical error details retain their native text.

## Earlier implementation verification

- **14 automated engine/model checks passed** with the workspace Python 3.9 environment and python-docx 1.2.0. See `tests/test_engine.py` and the README command to rerun them.
- **Hidden-window desktop integration passed** with both the initial workspace Python 3.9/Tk 8.5 and bundled Python 3.12/Tk 9.0. It exercised add, edit, multi-row remove, Decimal totals, text-mode controls, JSON save/load, actual DOCX generation, invalid form data, load errors, and unsaved-change cancellation/reset. Dialogs were mocked; the Tk widgets and document generation were real. This did **not** verify visible widget painting: the user subsequently reported a completely white window under Apple Tk 8.5. Treat that original runtime as unsupported, regardless of the hidden-window result.
- All application/test Python source files compiled successfully in memory.
- The CLI generated `output/sample-client.docx` from `sample_data/client.json` and reported **3 land tables updated**. Exact expected totals: **1.229 hectares** and **8.81 Lagan**.
- The original template's SHA-256 remains `825a595e9a5004f662ea2c1d19944a1780b005756e63ab9b480ae597dd9f4820`. The automated check compares every non-document ZIP member byte for byte, as well as section geometry, table grids/headers and untouched paragraphs.

## Mac white-window repair

The original `.venv` used Apple Command Line Tools Python 3.9.6/Tk 8.5 on macOS 27.0.1. A native screenshot subsequently confirmed that its content area was completely white. The environment is now rebuilt with the available Python 3.12.14/Tk 9.0 runtime; the old environment is retained in `.venv-apple-tk85/`.

All **16 automated checks** (14 engine/model checks plus two runtime compatibility checks) pass in the rebuilt environment. The hidden-window desktop integration check also passes under Tk 9.0. Launching with the old interpreter now fails with an actionable message before creating a window. Launching `.venv/bin/python app.py` starts the rebuilt runtime. The standalone modern interpreter is not exposed as a selectable application in the available native automation tool, so visible rendering of its window requires user confirmation; process liveness alone is not visual verification.

## Visual verification

Rendered the original and generated sample with the documents skill's bundled LibreOffice/Poppler renderer, then reviewed every output page. Both render to **12 pages**. Existing DevLys fonts in the Mac user's font directory had to be supplied explicitly to the isolated renderer using `FONTCONFIG_FILE` and `SAL_FONTPATH`; without them, it substituted fonts, showed Latin glyph codes, and rendered 15 pages. The DOCX and application do not depend on that QA configuration.

Visual review caught and corrected three issues:

1. Mixed Times New Roman/DevLys total labels must change only the count, retaining the Hindi suffix's original run.
2. New dates need isolated Times New Roman runs to display literal punctuation.
3. The five-column source schedule's Word `SUM(ABOVE)` fields refresh to a rounded area display. Only those two numeric total fields are materialized as static Decimal results with their existing formatting. The final render displays **1.229**, not 1.23. A final image comparison confirmed this correction changed only that total on page 3.

The sample retains the source's long questionnaire pagination, short continuation pages, and narrow nested area column (which wraps some source decimal values across lines). The original page geometry and table widths are retained. The final sample shows the existing Hindi text, all three schedules, exact totals and six populated date slots; no new clipping or overlap was observed compared with the source.

This is a LibreOffice rendering check, not a Microsoft Word rendering check. Word pagination can differ. Longer client values or more land rows will reflow the template and need visual review. Legacy input is explicitly pre-encoded; no conversion mapping or semantic spelling verification is claimed. Fixed legal statements and valuations remain as described in `TEMPLATE_INSPECTION.md`.
