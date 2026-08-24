# Recovery-Khata-MF

Excel **VBA** operations workbook for **Recovery Manager** (daily recovery registers) and **Khata** (ledger) books: year/month/employee folder trees, templated day-sheets, opening-balance chains, dropdowns, sheet protection, bulk refresh, and backups.

**Primary file:** `Recovery & Khata Master File/Recovery Manager.xlsm`  
**Reviewable VBA:** `_vba_extract/modules/`  
[rohaan2802](https://github.com/rohaan2802)

---

## Table of contents

1. [Workbook sheets and named ranges](#workbook-sheets-and-named-ranges)
2. [Modules](#modules)
3. [Typical operating sequence](#typical-operating-sequence)
4. [Python helpers](#python-helpers)
5. [Security](#security)
6. [How to open](#how-to-open)

---

## Workbook sheets and named ranges

From `workbook.xml`: **Home**, **Recovery Template**, **Khata Template**, **Settings**.

Named ranges drive the Home UI, including:

`txtRootFolder`, `selYear`, `selMonth`, `selEmployee`, `AreaList`, `empList`, `MonthList`, `YearList`, plus refresh status (`RefreshStatus`).

Staff never edit VBA to change period — they set these cells/dropdowns and press the bound macros.

---

## Modules

| Module | Responsibilities |
|--------|------------------|
| `modFolders` | `CreateNewYear`, `CreateNewMonth`, `CreateEmployeeFolder` — on-disk tree under the root |
| `modWorkbook` | `CreateRecoveryWorkbook` (and silent create), `CreateKhataWorkbook`, `RefreshKhataOpeningBalances`, `RebuildKhataOpeningChain`, `UpgradeExistingRecoveryRegisters`, working-day sheet generation, install **child** VBA into generated books |
| `modRecoveryChild` | Day-sheet `Change` handlers, **PTO carry-forward**, remaining formulas, working-day helpers |
| `modDropdowns` | Hidden lists sheet, named ranges, per-sheet dropdowns |
| `modProtection` | Auto-fit, editable regions, `SecureWorkbook`, **trusted location** helpers |
| `modRefresh` | `RefreshAllEmployeeWorkbooks` |
| `modBackup` | `WorkDoneCreateBackup` — walk tree, copy, auto-fit for archive |
| `ThisWorkbook` | `Workbook_Open`, sheet-change hooks, khata border / green-triangle cleanup |
| `modMain` / `modSettings` | Thin entry / settings notes |
| `Sheet1`–`Sheet4` `.cls` | Sheet-local event shells |

Generated recovery books get **child** code so daily tables validate even when the master is closed.

---

## Typical operating sequence

**Period setup**

1. Set root, year, month, employee (and areas from Settings).  
2. `CreateNewYear` → `CreateNewMonth` → `CreateEmployeeFolder`.  
3. `CreateRecoveryWorkbook` — working-day sheets from **Recovery Template**, dropdowns, protection, child VBA.  
4. Field staff fill tables; `modRecoveryChild` can carry PTO rows forward.

**Khata**

1. `CreateKhataWorkbook` from **Khata Template**.  
2. `RefreshKhataOpeningBalances` / `RebuildKhataOpeningChain` so month *n* opening matches month *n−1* closing.  
3. Bulk: `RefreshAllEmployeeWorkbooks` and watch Home `RefreshStatus`.

**Backup**

- `WorkDoneCreateBackup` copies the tree into a backup layout and auto-fits tables for print/PDF.

---

## Python helpers

Local scripts (`fix_workbook.py`, `update_mod_*.py`, `diagnose_*.py`, `test_*.py`) extract/repair VBA around the binary `.xlsm`. After VBA edits, **re-export** `.bas` / `.cls` into `_vba_extract/modules` so Git diffs stay readable.

OOXML dump: `_vba_extract/unzipped/xl/` (`workbook.xml`, sheets, tables).

---

## Security

- Sheet protection passwords live in VBA — treat as sensitive.  
- `EnsureTrustedLocation` exists so Excel will run macros; only trust paths you control.  
- **Do not commit live employee financial data.** Demo roots ≠ production roots.

---

## How to open

Windows + Excel, **macros enabled**. Open `Recovery Manager.xlsm`, trust the location if prompted, then use Home/Settings buttons (or Developer → Macros) for the names above.

Without Excel: read `_vba_extract/modules/*.bas`. Re-import via the VBA IDE after edits.

---

## Author

Internal operations automation · [rohaan2802](https://github.com/rohaan2802)
