# Recovery-Khata-MF

Excel **VBA** operations workbook for field **Recovery** day-registers and **Khata** (ledger) books: year / month / employee folders on disk, templated working-day sheets, opening-balance chains, dropdowns, sheet protection, bulk refresh, and backups.

**Primary file:** `Recovery & Khata Master File/Recovery Manager.xlsm`  
**Reviewable VBA:** `Recovery & Khata Master File/_vba_extract/modules/`

**Author:** Mohammad Rohaan · **Roll:** 22I-2327 · **GitHub:** [rohaan2802](https://github.com/rohaan2802)

---

## Table of contents

1. [Problem and context](#problem-and-context)
2. [Workbook sheets and tables](#workbook-sheets-and-tables)
3. [On-disk folder layout](#on-disk-folder-layout)
4. [Module reference](#module-reference)
5. [Recovery registers](#recovery-registers)
6. [Khata registers](#khata-registers)
7. [Dropdowns](#dropdowns)
8. [Protection and trusted locations](#protection-and-trusted-locations)
9. [Backup and bulk refresh](#backup-and-bulk-refresh)
10. [How to open (enable macros)](#how-to-open-enable-macros)
11. [OOXML extract](#ooxml-extract)
12. [Limitations](#limitations)
13. [Author](#author)

---

## Problem and context

The master `.xlsm` is a **manager** workbook. Staff pick a root folder, year, month, and employee on **Home**, then run macros that create Windows folders and generate per-employee **Recovery Register** and **Khata Register** workbooks (also `.xlsm`). Child books copy templates, Excel Tables, data-validation lists, and protection so daily entry can continue when the master is closed.

Language (GitHub): VBA. Default branch: `main`. Message boxes use the title **Recovery Manager**.

There is no Python runtime required to *use* the workbook. `_vba_extract/` exists so Git can diff VBA as text.

---

## Workbook sheets and tables

Four worksheets (Sheet1–Sheet4 class modules are empty shells; names come from VBA string literals):

| Sheet | Role in code |
|-------|----------------|
| **Home** | Named cells `txtRootFolder`, `selYear`, `selMonth`, `selEmployee`. Backup and refresh read the root from `Home!txtRootFolder`. |
| **Recovery Template** | Seed day-sheet copied into new recovery books. `ThisWorkbook_Open` clears Excel “green triangle” errors on **Remaining Amount**. |
| **Khata Template** | Seed ledger. `Workbook_SheetChange` restores **Balance** formulas and table borders. |
| **Settings** | ListObjects **`tblMonths`**, **`tblAreas`**, **`tblDayStatus`**, **`tblRecoveryStatus`**. Month folders are named `"N. MonthName"` from `tblMonths` (number + name). |

OOXML dump also has `xl/tables/table1.xml` … `table9.xml` (nine tables in the unzipped package).

Recovery day-sheet headers filled by VBA:

- `A8` — employee name  
- `C8` — long date (`dddd, dd mmmm yyyy`)  
- `E8` — **Area** dropdown (`AreasList`)  
- `G8` — **Day Status** dropdown (`DayStatusList`)

Summary cells rewritten by `FixSummaryFormulas` (table name `tblRecovery2`, `tblRecovery3`, …):

| Range | Formula purpose |
|-------|-----------------|
| B46 | `SUM([Bill Amount])` |
| B49 | `SUM([Recovered Amount])` |
| B52 | `SUM([Remaining Amount])` |
| B55 | count of non-blank **Reference No.** |
| B58 | count **Bill Status** = `"Returned"` |
| E46 | count `"Paid"` |
| E49 | count `"Partial"` |
| E52 | count `"PTO"` |
| E55 | count `"RETURNED"` |
| E58 | count `"HAND OVER TO SUPPLY MAN"` |

Khata **Balance** column (from `ThisWorkbook`):

```text
=IF(AND([@Debit]="",[@Credit]=""),"",
    SUM($E$8,INDEX([Debit],1):[@Debit])-SUM(INDEX([Credit],1):[@Credit]))
```

`E8` on a Khata sheet is the **opening balance**.

---

## On-disk folder layout

`modFolders` builds:

```text
{txtRootFolder}\{selYear}\{N. MonthName}\{selEmployee}\
```

Example: `D:\Data\2026\1. January\Ali\`.

Macros:

| Sub | Requires | Action |
|-----|----------|--------|
| `CreateNewYear` | root, year | `MkDir` year folder |
| `CreateNewMonth` | root, year, month | month folder via `GetMonthFolderName` |
| `CreateEmployeeFolder` | + employee | innermost folder |

Order is enforced (year before month before employee). Duplicate folders show an information `MsgBox`.

Generated filenames (`modWorkbook`):

```text
Recovery Register - {Month} {Year} - {Employee}.xlsm
Khata Register - {Month} {Year} - {Employee}.xlsm
```

Legacy `.xlsx` paths with the same stem are treated as “already exists”.

---

## Module reference

| File | Lines (extract) | Public surface |
|------|-----------------|----------------|
| `modMain.bas` | 6 | Comment only: entry lives in other modules |
| `modSettings.bas` | 6 | Comment: month naming is `GetMonthFolderName` |
| `modFolders.bas` | 182 | `CreateNewYear`, `CreateNewMonth`, `CreateEmployeeFolder`, `GetMonthFolderName` |
| `modWorkbook.bas` | 1412 | Create/refresh recovery + khata; working-day sheets; child VBA install |
| `modDropdowns.bas` | 332 | HiddenLists, named ranges, per-sheet validation |
| `modProtection.bas` | 411 | Trusted location, autofit, `SecureWorkbook` |
| `modRefresh.bas` | 358 | `RefreshAllEmployeeWorkbooks` |
| `modBackup.bas` | 356 | `WorkDoneCreateBackup` |
| `ThisWorkbook.cls` | 90 | `Workbook_Open`, Khata change handler |
| `Sheet1.cls`–`Sheet4.cls` | 10 each | Empty |

---

## Recovery registers

`CreateRecoveryWorkbook`:

1. Read Home named ranges; abort if any blank.
2. `EnsureTrustedLocation` on root and employee folder.
3. `Workbooks.Add`; copy **Recovery Template**.
4. `CreateHiddenListsSheet`, `CreateNamedRanges`, `ApplyDropdownsToSheet`.
5. `GenerateWorkingDaySheets` — for each calendar day in the month **except Friday** (`Weekday(..., vbSunday) <> vbFriday`), copy the seed, name the sheet `dd-mmmm-yyyy` (fallback `dd-mmm-yyyy`), fill A8/C8, `FixSummaryFormulas`.
6. Delete leftover `Sheet1` and the seed “Recovery Template” sheet.
7. `InstallAutoOpenSheetCode` + `ActivateStartupSheet` (open on today’s working day).
8. Save as `.xlsm`, `SecureWorkbook`, close.

During create, calculation is manual, events off, `AutomationSecurity = 3` (skip `Workbook_Open` in files being written). Remaining Amount is forced numeric (`SetupRemainingAmountColumn`) so green error triangles stay off.

---

## Khata registers

`CreateKhataWorkbook`:

- Copies **Khata Template**.
- `A8` = employee; `C8` = `"Month, Year"`; `E8` = `GetPreviousKhataClosingBalance` (0 if first book).
- Sheet renamed `{Employee} {Month} Khata` (31-char Excel limit; timestamp suffix on clash).
- `InstallKhataCascadeCode` injects child VBA so later months can chain.
- `RefreshKhataOpeningBalances` → `RebuildKhataOpeningChain`: collect that employee’s Khata files under the root, sort by period, set each opening from the previous file’s closing (`prevClose`; first file opening **0**).

---

## Dropdowns

`CreateHiddenListsSheet` adds a **very hidden** sheet `HiddenLists` with columns Areas / Day Status / Recovery Status, copied from Settings tables.

Named ranges:

| Name | Refers to |
|------|-----------|
| `AreasList` | `HiddenLists!A2` dynamic `INDEX/COUNTA` |
| `DayStatusList` | column B |
| `RecoveryStatusList` | column C |

`ApplyDropdownsToSheet`: list validation on Area (E8) and Day Status (G8), and on the table **Bill Status** column (`RecoveryStatusList`). `UpdateWorkbookDropdowns` can refresh an existing file; `applyValidations:=False` is a fast path (lists only). Already-open workbooks are not closed.

---

## Protection and trusted locations

`modProtection.bas` defines `SHEET_PASSWORD` (hard-coded string). Treat it as sensitive; rotate if this repo is shared.

`SecureWorkbook`:

- Unprotect workbook structure.
- For every sheet except `HiddenLists`, `ProtectSheetAllowTableEdit`: contents protected, **insert/delete rows**, sort, filter, and formatting allowed; **delete columns** disallowed; `UserInterfaceOnly:=True`.
- Protect workbook structure with the same password.

`EnsureTrustedLocation` writes HKCU Office Excel **Trusted Locations** for the root (and subfolders), sets `AllowNetworkLocations`, and labels the location **Recovery Manager**. That is why macros can run from employee folders without repeated Trust Center prompts — only use this on machines you control.

`CloseOpenRecoveryRegisters` saves and closes other workbooks whose names contain **Recovery Register**.

Autofit helpers: `CopyColumnWidths` (columns 1–20, rows 6–9), `FitMergedCell`, `FitSheetContent`, `AutoFitWorkbook`.

---

## Backup and bulk refresh

`WorkDoneCreateBackup` copies every workbook under `Home!txtRootFolder` to **`{drive}:\Recovery_Backup`** (hidden folder; e.g. `D:\Ops` → `D:\Recovery_Backup`), preserving relative paths, with optional `FastFitWorkbook` on tables. Status bar: `Work Done & Backup: i / n`. Success: **Work Done & Sucessfull Back Up** (spelling as in VBA). Failures: **Failed Back Up**. `AutomationSecurity = 3` skips `Workbook_Open` while files are opened for fitting.

`RefreshAllEmployeeWorkbooks` walks Recovery Register files (`CollectRecoveryFiles`), loads Areas / Day Status / Recovery Status **once** from Settings, then `FastRefreshOneWorkbook` per file. Status bar: `Refreshing dropdowns: i / n`. Final **Success.** or **Failed.** `BringManagerToFront` keeps the manager window visible.

---

## How to open (enable macros)

1. Windows + **Microsoft Excel** (desktop). Macros in `.xlsm` will not run in Excel for the web.
2. Open `Recovery & Khata Master File/Recovery Manager.xlsm`.
3. If prompted, **Enable Content** / Enable Macros. Prefer trusting the folder (the VBA also registers a Trusted Location).
4. On Home: set root folder, year, month, employee (dropdowns from Settings).
5. Run macros from assigned buttons or **Developer → Macros**:
   - Folders: `CreateNewYear` → `CreateNewMonth` → `CreateEmployeeFolder`
   - Books: `CreateRecoveryWorkbook`, `CreateKhataWorkbook`
   - Maintain: `RefreshKhataOpeningBalances`, `RefreshAllEmployeeWorkbooks`, `WorkDoneCreateBackup`

Without Excel, read `_vba_extract/modules/*.bas`. After editing `.bas`, re-import in the VBA IDE (Alt+F11) and save the `.xlsm`.

---

## OOXML extract

```text
Recovery & Khata Master File/Recovery Manager.xlsm
Recovery & Khata Master File/_vba_extract/modules/   # .bas / .cls above
Recovery & Khata Master File/_vba_extract/unzipped/xl/
  workbook.xml, worksheets/sheet1–4.xml, tables/table1–9.xml, vbaProject.bin
```

---

## Limitations

- Windows + desktop Excel; folder APIs and Trusted Location registry writes are Windows-oriented.
- Sheet password is a source constant — not a security boundary.
- Fridays are omitted from recovery day sheets by design.
- Backup always uses `{drive}:\Recovery_Backup`, not a user-chosen path.
- **Do not commit live employee financial workbooks.** Point `txtRootFolder` at a demo tree.
- `modMain` / `modSettings` are stubs; Sheet class modules have no events.

---

## Author

**Mohammad Rohaan** · Roll **22I-2327** · [github.com/rohaan2802](https://github.com/rohaan2802)
