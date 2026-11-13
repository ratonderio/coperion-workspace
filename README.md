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

Each folder holds the exported VBA for one workbook: standard modules (`.bas`),
class modules (`.cls`), forms (`.frm` with its `.frx`) and sheet/workbook code
(`.doccls`). The workbooks themselves (`.xlsm`/`.docm`) are not stored here.
They live on the network drive (see Network paths), and `.gitignore` keeps them
out of the repo. Worksheet modules with no code are not exported either.

Keep each `.frm` and `.frx` together. A form imported without its `.frx` loses
its control layout and images.

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

- Edit the code in the repo, then update the live workbook: download the repo
  as a zip, delete the changed modules in the workbook's VBA project, import
  the matching files from the zip, and save over the file on the network drive.
  Archive the previous version in the archive folder first.
- The VBA editor and Rubberduck add `Attribute` header lines when exporting.
  If code is edited in the workbook instead, re-export it back into
  `ww-tools/<tool>/` so the repo does not fall behind.
- `.doccls` files hold sheet and `ThisWorkbook` code. These cannot be imported
  as new modules; paste their code into the existing sheet module instead.
- Some modules require references (e.g. Microsoft ActiveX Data Objects 6.1,
  Microsoft VBScript Regular Expressions, Scripting Runtime, Outlook, and the
  SAP GUI Scripting library).
- Some exported files use a legacy Windows-1252 encoding rather than UTF-8
  (for example `Emails.bas`), so use tooling that preserves encoding.
- Exports from the VBA editor use CRLF line endings; the repo copies are LF.
  Whitespace-only differences of this kind are not real changes.

## Network paths

The shared files now live on a consolidated file server that is mapped as `I:`
(`I:` = `\\cop.hi.corp\fileserver\CFE\CPS`). Several shares that used to be
separate servers are now folders under it, for example:

| Old UNC | New location |
| --- | --- |
| `\\uswwq-p-fs01.cps.local` | `I:\WWQ-FS01` (`\\cop.hi.corp\fileserver\CFE\CPS\WWQ-FS01`) |

The old UNC paths still resolve to the same place, so the old names found in
the code (`\\uswwq-p-fs01.cps.local\...`, `\\USWWQ-P-FS01\...`) keep working for
now. This looks like redirection set up by IT, not a permanent guarantee. When
a path stops working, check here first. Prefer the UNC form over the `I:`
letter in code, because drive letters are mapped per user.

## Email addresses

Addresses are now `first.last@ktron.com`. The shared engineering clerk mailbox
(`fpm_PAC_EngineeringClerk`) has no ktron.com equivalent yet, so the code still
uses the coperion.com address, which redirects to the right people. It is
defined once as `ENGINEERING_CLERK_EMAIL` in `Schedule.bas` and in the order
entry module, so switching to a ktron.com group later is a one line change per
workbook. Some addresses also come from the `WW_Prod_Eng_Names` SQL table
(`WW_Email`), which is outside this repo.

## Notes

- Several modules contain database connection details as constants. Moving
  these out of source control is worth doing as these tools are modernized.
