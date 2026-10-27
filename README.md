# NBD Configuration Generator

Use Node.js 18+ (Windows PowerShell users can run `npm.cmd` to avoid npm.ps1 execution-policy errors)..

```bash
npm.cmd install
node generate.js NBD_Zevra_Miplyffa_2_HCPs_v2.1.xlsx
```

The script finds `Section` and `Content` headers on every Excel tab and writes one JSON document and one validation report per sheet in `output/`. `identities.json` overrides person IDs and names where the checklist does not supply them. Missing `person_id` is flagged (Shayota's cell is empty in the workbook).

**Important:** The supplied checklist is prose, not a fully structured data specification. The generator faithfully fills basic NBD strings but cannot infer scores, date metadata, or standardized connection fields. These stay empty, and warnings identify missing structured data. No MongoDB insert is performed, and no `_id` is reused. Review scientific content, HTML, links, and score/event requirements manually before import.

For production-ready, exact formatting (headings, lists, dates, source links) add dedicated structured Excel columns for these fields. `Where to Connect` maps to `when_to_connect`.

## Dependency note

This version uses ExcelJS instead of SheetJS (`xlsx`). As of the reported audit, ExcelJS may introduce moderate-severity transitive `uuid` advisories. Run `npm.cmd audit` and review the installed dependency tree; do not use `npm audit fix --force` without testing because it may downgrade ExcelJS. Consider running the generator only on trusted internal workbooks and keep it offline from production MongoDB.
