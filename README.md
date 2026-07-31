# Repository-friendly export of WB files

This package contains generated, text-readable snapshots of 7 source files.

## Recommended source-of-truth model

1. Keep the original `.xlsx`, `.xlsm` and `.pdf` files under `data/source/`.
2. Commit the generated JSON/JSONL under `data/json/`.
3. Treat `manifest.json` and SHA-256 hashes as the link between originals and generated snapshots.
4. Regenerate snapshots with `scripts/export_office_to_repo_json.py` whenever a source file changes.
5. Never replace the original `.xlsm` with JSON: VBA, buttons, drawings, styles and Excel behavior are not fully round-trippable.

## Why JSONL instead of one huge JSON file

Large WB reports contain hundreds of thousands of cells. One monolithic JSON file is difficult to diff, review and stream. Therefore each sheet is exported as sharded `records.part-XXXXX.jsonl` files, one logical row per line. Small `preview.json` files make quick review easy in GitHub.

## Workbook export layout

- `workbook.json` - workbook-level metadata, source checksum, sheet index, defined names and calculation properties.
- `styles.json` - style indexes and number formats used to decode dates.
- `sheets/<sheet>/sheet.meta.json` - dimensions, merged ranges, header mapping, tables, hyperlinks, row/column properties and shard index.
- `sheets/<sheet>/records.part-*.jsonl` - values using readable header names and original row number `_row`.
- `sheets/<sheet>/formulas.part-*.jsonl` - formulas, cached results and shared-formula attributes.
- `sheets/<sheet>/preview.json` - preamble plus first 50 records.
- `vba/source/*` - readable VBA modules extracted from the XLSM through LibreOffice, when available.

## PDF export layout

- `document.json` - metadata, section index and page index.
- `pages/page-XXX.json` - page text, positioned text blocks and links.

## Important limitations

- Formula values are cached values stored in the source workbook; this export does not recalculate Excel formulas.
- JSON is an audit/read layer, not a replacement for the working Excel model.
- VBA text extracted through LibreOffice is intended for review and diffs; run macros only from the original XLSM in Microsoft Excel.
- PDF tables are preserved as text/layout blocks, not guaranteed relational tables.

## Suggested repository layout

```text
data/
  source/                 # original XLSX/XLSM/PDF, authoritative
  json/                   # contents of this package's data/json
scripts/
  export_office_to_repo_json.py
manifest.json
```

## Validation

Every source and generated shard has a SHA-256 checksum. `validation_report.json` records parse and count checks performed after export.
