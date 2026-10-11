# Inspection of the original template

Inspected directly from `templates/Hindi Ah Boi  4 Old Vrndavan Mandkheda 22.09.2026.docx` before implementing the engine.

SHA-256: `825a595e9a5004f662ea2c1d19944a1780b005756e63ab9b480ae597dd9f4820`.
The application verifies this fingerprint and refuses an uninspected template. It uses no files in `Docs/`.

One portrait section: page size 11909 × 16834 twips (A4), top margin 1440, bottom margin 1152, left/right margins 1440 twips. With the installed DevLys fonts explicitly available to the bundled LibreOffice renderer, both original and sample output render to 12 pages. Source pagination includes long legal tables splitting across pages and short continuation pages; these are retained.

## Editable fields

These are existing values, not Word merge fields or placeholders. The strings below are raw glyph codes, **not an asserted Unicode translation**.

| Field | Existing raw value | Scope |
| --- | --- | --- |
| Client name | ``o`Unkou`` | All 13 client mentions, including the one-cell identity table and report table |
| Father's name | `txUukFk` | All 13 corresponding mentions |
| Age | Spaces between `vk;q` and `tkfr` | Both affidavit age slots, inserted as literal digits |
| Caste | `czkEg.k` | All ten occurrences, including the separate caste line |
| Address | `eaM[ksMk` after `fuoklh`, sometimes preceded by `xzke` | 9 residence references; distinct from land village |
| Village | `eaM[ksMk` | Land-location references and schedule village cell |
| Tehsil | `';keiqj` or `’;keiqj` | Client/land-location references; registrar jurisdiction text remains intact |
| District | `lhgksj` | Client/land-location references and schedule cell; bank, advocate, execution city and registrar references remain intact |
| Date | Six blank/partial date slots | Affidavit verification, report date, search cutoff, certificate date, bank-letter date, declaration date |
| Bahi | `LD-445692` | Four prose occurrences plus five-column schedule; one document-level value used consistently |
| Total area | `1.229` | Two prose summaries plus every land schedule |

Historical dates `25@08@2026` and `11-10-2022`, record years, loan amounts, advocate details, bank branch, patwari number 13, property valuations, and legal declarations are outside this form's editable scope. Existing assertions and valuations require human review for each client; the program does not recalculate valuations or revise legal prose.

## Tables

Table numbers below are zero-based XML document order and include nested tables.

| Table | Structure | Treatment |
| --- | --- | --- |
| 0 | 6 × 4: serial, Khasra, area, Lagan | Header + 4 data rows + total; rebuilt from cloned rows |
| 1 | 1 × 1 identity paragraph | Targeted field replacements |
| 2 | 6 × 5: serial, Khasra, area, Lagan, Bahi | Bahi originally vertically merged across all body and total rows |
| 3 | 24 × 2 legal questionnaire | Preserve questions and prose; update targeted identity/geography mentions |
| 4 | 1 × 1 (`ugh`) | Preserve |
| 5 | Declaration schedule: 11-column grid; 3 rows, physical cell counts 11, 11, 9 | Preserve merged columns and layout; update location cells |
| 6 | 6 × 4, nested in table 5 row 2 cell 5 | Same land data and total as table 0 |

The engine walks all descendant `w:tbl` elements. Four-column schedules retain their original columns; the single document-level Bahi appears in the five-column schedule with the original vertical merge retained across body and total rows. Row count follows the input, serial numbers are generated, and the total label retains the original `fdrk` suffix. Table grids, widths, borders, headers, cell/paragraph/run properties and surrounding paragraphs are retained. Pagination can change with more rows or longer values.

The five-column schedule's area and Lagan totals contain complex `SUM(ABOVE)` Word fields with cached results. Rendering can refresh these and round the area to two decimal places. The engine materializes only these two fields as static Decimal results, retaining their result-run properties, so the displayed totals cannot be recalculated or rounded by Word. The other two schedules have plain-text totals.

## Hindi font encoding

In `word/document.xml`, explicit `w:rFonts/@w:ascii` references include DevLys 010 (1126), DevLys 240 (28), DevLys 160 (1), Times New Roman (135), Book Antiqua (2), and Courier New (1). These are property references, including paragraph properties, not counts of visible words. The font table also lists Mangal, Calibri and other fonts. Default style fonts are Calibri with Mangal for complex script. No font files are embedded.

The Hindi prose uses Latin/Western code points under DevLys. XML itself is UTF-8, but the visible Hindi depends on the DevLys font's glyph mapping. A font name alone does not establish a verified Unicode↔DevLys mapping. None is supplied or guessed.

`legacy_devlys` accepts explicitly pre-encoded strings and preserves field font properties. It rejects Devanagari/other unrelated Unicode scripts. It cannot prove that entered glyph codes spell the intended Hindi. `unicode` is an explicit user selection: only inserted human text and identifiers use the Mangal font. Existing DevLys text keeps its font. Numeric values and serials remain in existing template runs. New dates use isolated Times New Roman runs so separators render literally rather than as DevLys glyphs. Age preserves its original DevLys run in legacy mode (ASCII digits were verified in the render), and uses an isolated Mangal run in Unicode mode. The mixed-font total label is retained: only its count changes, leaving the Hindi suffix in its original font. Correct rendering still requires original DevLys fonts, including in Unicode mode.

## Preservation contract

Only `word/document.xml` is changed in the generated ZIP. Every other package member, including styles, fonts, settings, theme, numbering, footnotes, endnotes, relationships and custom XML, is copied byte for byte. No header/footer parts are present. Word text-box content includes advocate contact details and remains untouched. The source is never saved or overwritten. Original document metadata is retained, so review it before sharing.
