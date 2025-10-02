# TB&S Automation — Phase 1

Phase-1 TB&S automation: normalize dispatch data and generate invoices from a control sheet—documentation, runbook, and support included.

## Live site (GitHub Pages)
Live site: https://lvlsetnext-1.github.io/tbs-automation/
Short link: https://bit.ly/48Lttrt

## Operational workbooks (production)
- **Shared Dispatch (Export):** https://docs.google.com/spreadsheets/d/1ke0ILFUkq8uwbMapm6NAPBoUXYSgsf_t85imQlQedAM/edit?usp=sharing
- **Delivery Schedule:** https://docs.google.com/spreadsheets/d/1ZyLIDlFM7cV6txxHNp7PT8KGFxv689093qfHHDYn814/edit?usp=sharing
- **Invoice Master:** https://docs.google.com/spreadsheets/d/1vNjSvLf7KpJ2lWTiJfBBpRrAduevIugi5_0NlNo046g/edit?usp=sharing

## Docs
See `docs/LINKS.md`.

## How it works (happy path)
1) Dispatcher updates **Working/Dispatch_Entry**; **Export** auto-updates (audit: **Y** “Last Updated”, **Z1** timestamp).  
2) Tracker/Biller (TB&S menu): **Clone Template → Sync From Dispatch → (Admin: Apply Validations if needed) → Create Invoices**.  
3) Protections: totals/headers locked; operators edit only unlocked memo/notes.

## Apps Script (control workbook)
`onOpen()`, `buildMenu_()`, `cloneTemplate_()`, `syncExport_()`, `createInvoices_()`, `applyValidations_()`, `fixCommonIssues_()`, `carryOver_()`, `getConfig_()`, `getInvoiceTemplate_()`, `getInvoiceMap_()`, `protectTotalsAndHeaders_()`.

## License
MIT (see `LICENSE`).

