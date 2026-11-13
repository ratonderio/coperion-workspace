# Coperion Workspace

A personal workspace for version controlling and modernizing the engineering
tooling I build and maintain at work. The repository is named "Coperion" because
that was the company name when it was created. The company has since been renamed
again (now K-Tron), but the repo name is unchanged.

The original goal was to put the VBA behind our most heavily used Excel
workbooks under version control. It has grown into a place to experiment with
newer approaches (for example Windchill OData) as these tools are modernized.

## Repository layout

```
coperion-workspace/
├── windchill/   OData metadata for Windchill (XML + odata.json)
└── ww-tools/    Exported VBA workbooks, one folder per tool
    ├── order-review/
    ├── production-engineering-manager/
    ├── production-engineering-schedule/
    ├── production-order-entry/
    ├── scheduler-status/
    ├── scheduling-tool/
    └── sap-tools/
```

### `windchill/`

XML metadata documents (one per Windchill domain, e.g. `DocMgmt.xml`,
`ChangeMgmt.xml`, `ProdMgmt.xml`) plus `odata.json`, used as reference while
working with Windchill's OData API.

### `ww-tools/`

Each folder holds the macro-enabled workbook (`.xlsm`/`.docm`) alongside its
exported VBA: standard modules (`.bas`), class modules (`.cls`), forms
(`.frm`/`.frx`) and sheet/workbook code (`.doccls`). The workbook is the
deployed artifact; the exported code is what makes diffs and history usable.

| Tool | Purpose |
| --- | --- |
| `order-review` | Order review forms (EE and ME). Mostly saves a copy of the completed form to the appropriate network folder. |
| `production-engineering-manager` | "Front end" for manually updating SQL fields used for tracking. Overlaps with some functionality in the production engineering schedule. |
| `production-engineering-schedule` | Each engineer's personal schedule. Includes the approval and certified date dashboard, order comments, scripts that generate email templates, scripts that download associated order files, and add/delete entry and history. |
| `production-order-entry` | Entry of new orders and assignment to engineers, plus change orders and their emails. This is where orders are inserted into the `Prod_Eng` SQL table. |
| `scheduler-status` | What orders each scheduler has, with comments and associated dates. |
| `scheduling-tool` | Used by schedulers to email the engineering manager their intended schedule for new orders and announce that orders need assignment. |
| `sap-tools` | Creating BOMs and new part numbers in SAP, plus searching SQL databases that replicate SAP data for easier retrieval. |

## How the tools fit together

1. Sales sends an order and an **order review** is completed. The review is
   saved to a network drive.
2. The **schedulers** build a schedule for the order and generate an order
   announcement PDF using the **scheduling tool**, which goes to the
   engineering manager.
3. The engineering manager uses **production order entry** to assign the order
   to engineers and insert it into the `Prod_Eng` database for tracking.
4. Engineers manage their own work in the **production engineering schedule**
   (status, approvals, certified/release dates, comments, add/deletes).
5. **Scheduler status** gives schedulers visibility into their orders, and the
   **SAP tools** support the BOM and part number work throughout.

## Working with the VBA

- Edit in the VBA editor inside the workbook, then re-export modules into the
  matching `ww-tools/<tool>/` folder so the repo stays in sync with the workbook.
- Some modules require references (e.g. Microsoft ActiveX Data Objects 6.1,
  Microsoft VBScript Regular Expressions, Scripting Runtime, Outlook, and the
  SAP GUI Scripting library).
- Some exported files use a legacy Windows-1252 encoding rather than UTF-8
  (for example `Emails.bas`), so use tooling that preserves encoding.
- Exports from the VBA editor use CRLF line endings; the repo copies are LF.
  Whitespace-only differences of this kind are not real changes.

## Notes

- Several modules contain database connection details as constants. Moving
  these out of source control is worth doing as these tools are modernized.
