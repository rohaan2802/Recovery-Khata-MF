# Recovery-Khata-MF

Excel **VBA automation** for **Recovery Manager** and **Khata** (ledger) master workbooks: year/month/employee folder trees, templated recovery day-sheets, khata opening-balance chains, dropdowns, sheet protection, bulk refresh, and backup utilities.

**Primary workbook:** `Recovery & Khata Master File/Recovery Manager.xlsm`  
**Extracted VBA (for review/diff):** `_vba_extract/modules/`

---

## Overview

Recovery-Khata-MF is an operations workbook used to manage field **recovery registers** and **khata** books per employee and calendar period. From the **Home** sheet, users set a root folder, year, month, and employee; macros then:

1. Create the on-disk folder hierarchy.
2. Generate recovery workbooks with working-day sheets from **Recovery Template**.
3. Generate / refresh khata workbooks from **Khata Template**, including opening-balance cascade.
4. Apply named-range dropdowns, protect sheets while allowing table edits, and optionally backup or refresh all employee files.

Workbook sheets (from `workbook.xml`): **Home**, **Recovery Template**, **Khata Template**, **Settings**.

Named ranges drive UI state (`txtRootFolder`, `selYear`, `selMonth`, `selEmployee`, `AreaList`, `empList`, `MonthList`, `YearList`, ...).

---

## Features

| Module | Role |
|--------|------|
| `modFolders` | `CreateNewYear`, `CreateNewMonth`, `CreateEmployeeFolder` |
| `modWorkbook` | `CreateRecoveryWorkbook` / silent create, `CreateKhataWorkbook`, `RefreshKhataOpeningBalances`, `RebuildKhataOpeningChain`, `UpgradeExistingRecoveryRegisters`, working-day sheet generation, install child VBA |
| `modRecoveryChild` | Day-sheet change handlers, PTO carry-forward, remaining formulas, working-day helpers |
| `modDropdowns` | Hidden lists sheet, named ranges, per-sheet dropdown application |
| `modProtection` | Auto-fit, editable locks, `SecureWorkbook`, trusted location helpers |
| `modRefresh` | `RefreshAllEmployeeWorkbooks` |
| `modBackup` | `WorkDoneCreateBackup` - scan folders, copy/fit workbooks into backup tree |
| `ThisWorkbook` | `Workbook_Open`, sheet-change hooks, khata border / green-triangle cleanup |
| `modMain` / `modSettings` | Entry/shared notes; settings helpers |

Python helpers in the local tree (`fix_workbook.py`, `update_mod_*.py`, `diagnose_*.py`, ...) support VBA extract/repair workflows around the `.xlsm`.

---

## Repository structure

```text
Recovery-Khata-MF/
└── Recovery & Khata Master File/
    ├── Recovery Manager.xlsm
    ├── _vba_extract/
    │   ├── modules/
    │   │   ├── modFolders.bas
    │   │   ├── modWorkbook.bas
    │   │   ├── modRecoveryChild.bas
    │   │   ├── modDropdowns.bas
    │   │   ├── modProtection.bas
    │   │   ├── modRefresh.bas
    │   │   ├── modBackup.bas
    │   │   ├── modMain.bas | modSettings.bas
    │   │   ├── ThisWorkbook.cls
    │   │   └── Sheet1-4.cls
    │   └── unzipped/xl/...          # OOXML dump (workbook.xml, sheets, tables)
    └── *.py                       # maintenance / diagnose scripts (local)
```

---

## Build / run

### Runtime

1. Windows + **Microsoft Excel** with **macros enabled**.
2. Open `Recovery Manager.xlsm`.
3. If prompted, trust the file location (`modProtection.EnsureTrustedLocation` supports this flow).
4. On **Settings** / **Home**, set:
   - Root folder path (`txtRootFolder`)
   - Year / month / employee (and areas as needed from Settings tables)
5. Run UI-bound macros (buttons or Developer → Macros), commonly:
   - Create year → month → employee folders
   - Create Recovery workbook / Create Khata workbook
   - Refresh khata openings / refresh all employee workbooks
   - Work-done backup

### Viewing VBA without Excel

Open `_vba_extract/modules/*.bas` in any editor. Re-import into the `.xlsm` via the VBA IDE if you change sources.

---

## Usage

**Typical recovery period setup**

1. Choose root directory for the company/season.
2. `CreateNewYear` → `CreateNewMonth` → `CreateEmployeeFolder`.
3. `CreateRecoveryWorkbook` - generates working-day sheets, installs recovery child logic, applies protection/dropdowns.
4. Staff enter daily recovery table data; `modRecoveryChild` validates changes and can carry PTO rows forward.

**Khata flow**

1. `CreateKhataWorkbook` from the Khata template.
2. `RefreshKhataOpeningBalances` / `RebuildKhataOpeningChain` to keep month-to-month balances consistent.
3. Use Home refresh status (`RefreshStatus`) after bulk `RefreshAllEmployeeWorkbooks`.

**Backup**

- `WorkDoneCreateBackup` walks the tree, copies workbooks, and auto-fits tables for archival readability.

---

## Extending

- Keep business rules in `modRecoveryChild` / `modWorkbook`; leave `modMain` as a thin pointer module.
- After VBA edits, re-export modules to `_vba_extract/modules` so Git diffs stay reviewable (binary `.xlsm` alone is opaque).
- Parameterize protection passwords and template sheet names in `Settings` tables.
- Add automated smoke tests that open the workbook via `win32com` and call `CreateNewMonth` against a temp folder (see local `test_*.py` patterns).

---

## Security notes

- Workbook protection and editable-region locks are enforced in VBA - treat passwords and trusted locations carefully.
- Do not commit live employee financial data; keep sample/demo folders separate from production roots.

---

## License

Internal / operational automation tool - restrict distribution of production workbooks containing personal or financial data.
