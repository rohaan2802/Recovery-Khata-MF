Attribute VB_Name = "modProtection"
Option Explicit

Public Const SHEET_PASSWORD As String = "2968"

Public Sub EnsureWorkbookWindowsVisible(ByVal wb As Workbook)

    Dim wn As Window

    On Error Resume Next
    For Each wn In wb.Windows
        wn.Visible = True
    Next wn
    On Error GoTo 0

End Sub

Public Sub EnsureTrustedLocation(ByVal folderPath As String)

    Dim sh As Object
    Dim ver As String
    Dim base As String
    Dim secKey As String
    Dim pathVal As String
    Dim i As Long
    Dim normalized As String

    If Len(Trim$(folderPath)) = 0 Then Exit Sub

    normalized = Trim$(folderPath)
    Do While Right$(normalized, 1) = "\"
        normalized = Left$(normalized, Len(normalized) - 1)
    Loop
    If Len(normalized) = 0 Then Exit Sub
    normalized = normalized & "\"

    ver = Application.Version
    secKey = "HKCU\Software\Microsoft\Office\" & ver & "\Excel\Security\"
    base = secKey & "Trusted Locations\"

    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")

    ' Make sure Trusted Locations feature is usable
    sh.RegWrite base & "AllowNetworkLocations", 1, "REG_DWORD"
    ' 1 = Enable all macros with notification bypass inside trusted locations
    ' Keep Trust Center allowing trusted locations (0 = not disabled)
    sh.RegWrite base & "AllTrustedLocationsDisabled", 0, "REG_DWORD"

    ' Already listed?
    For i = 0 To 100
        Err.Clear
        pathVal = CStr(sh.RegRead(base & "Location" & CStr(i) & "\Path"))
        If Err.Number = 0 Then
            If StrComp(LCase$(pathVal), LCase$(normalized), vbTextCompare) = 0 Or _
               StrComp(LCase$(pathVal), LCase$(Left$(normalized, Len(normalized) - 1)), vbTextCompare) = 0 Then
                sh.RegWrite base & "Location" & CStr(i) & "\AllowSubfolders", 1, "REG_DWORD"
                sh.RegWrite base & "Location" & CStr(i) & "\Description", "Recovery Manager", "REG_SZ"
                On Error GoTo 0
                Exit Sub
            End If
        End If
    Next i

    ' Add new trusted location (includes all subfolders / any drive where Root Folder lives)
    For i = 0 To 100
        Err.Clear
        pathVal = CStr(sh.RegRead(base & "Location" & CStr(i) & "\Path"))
        If Err.Number <> 0 Then
            Err.Clear
            sh.RegWrite base & "Location" & CStr(i) & "\Path", normalized, "REG_SZ"
            sh.RegWrite base & "Location" & CStr(i) & "\AllowSubfolders", 1, "REG_DWORD"
            sh.RegWrite base & "Location" & CStr(i) & "\Description", "Recovery Manager", "REG_SZ"
            sh.RegWrite base & "Location" & CStr(i) & "\Date", Format$(Now, "yyyy-mm-dd"), "REG_SZ"
            Exit For
        End If
    Next i

    On Error GoTo 0

End Sub

Public Sub CloseOpenRecoveryRegisters()

    Dim wb As Workbook
    Dim i As Long

    Application.DisplayAlerts = False

    For i = Application.Workbooks.Count To 1 Step -1
        Set wb = Application.Workbooks(i)
        If Not wb Is ThisWorkbook Then
            If InStr(1, wb.Name, "Recovery Register", vbTextCompare) > 0 Then
                On Error Resume Next
                EnsureWorkbookWindowsVisible wb
                wb.Save
                wb.Close SaveChanges:=False
                On Error GoTo 0
            End If
        End If
    Next i

End Sub

' Keep Recovery Template column widths exactly
Public Sub CopyColumnWidths(ByVal src As Worksheet, ByVal dst As Worksheet)

    Dim i As Long

    On Error Resume Next
    For i = 1 To 20
        dst.Columns(i).ColumnWidth = src.Columns(i).ColumnWidth
    Next i
    ' Header / value row heights from template
    dst.Rows("6:9").RowHeight = src.Rows("6:9").RowHeight
    On Error GoTo 0

End Sub

' Fit one cell (merged or not). Expands width/height if needed — never shrinks below current.
Public Sub FitMergedCell(ByVal cell As Range)

    Dim area As Range
    Dim firstCell As Range
    Dim i As Long
    Dim oldWidths() As Double
    Dim oldTotal As Double
    Dim needed As Double
    Dim wasAlerts As Boolean

    If cell Is Nothing Then Exit Sub

    On Error Resume Next

    If Len(Trim$(CStr(cell.MergeArea.Cells(1, 1).Value))) = 0 Then Exit Sub

    If Not cell.MergeCells Then
        needed = cell.ColumnWidth
        cell.EntireColumn.AutoFit
        If cell.ColumnWidth < needed Then cell.ColumnWidth = needed
        cell.EntireRow.AutoFit
        Exit Sub
    End If

    Set area = cell.MergeArea
    Set firstCell = area.Cells(1, 1)

    ReDim oldWidths(1 To area.Columns.Count)
    oldTotal = 0
    For i = 1 To area.Columns.Count
        oldWidths(i) = area.Columns(i).ColumnWidth
        oldTotal = oldTotal + oldWidths(i)
    Next i

    wasAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    area.UnMerge

    firstCell.EntireColumn.AutoFit
    needed = firstCell.ColumnWidth

    area.Merge
    Application.DisplayAlerts = wasAlerts

    ' Restore template widths first
    For i = 1 To area.Columns.Count
        area.Columns(i).ColumnWidth = oldWidths(i)
    Next i

    ' Expand only if content needs more space
    If area.Columns.Count = 1 Then
        If needed > oldWidths(1) Then
            firstCell.ColumnWidth = needed
        End If
    Else
        If needed > oldTotal Then
            firstCell.ColumnWidth = oldWidths(1) + (needed - oldTotal)
        End If
    End If

    area.Rows.EntireRow.AutoFit
    On Error GoTo 0

End Sub

' Creation + Work Done: fit headings/values/table data; keep template widths as baseline
Public Sub FitSheetContent(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim col As ListColumn
    Dim minW As Double
    Dim needed As Double

    On Error Resume Next

    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then Exit Sub

    ' Headings
    FitMergedCell ws.Range("A6")   ' SALES MAN
    FitMergedCell ws.Range("C6")   ' DATE
    FitMergedCell ws.Range("E6")   ' AREA
    FitMergedCell ws.Range("G6")   ' DAY STATUS

    ' Values (date cell is critical)
    FitMergedCell ws.Range("A8")   ' employee name
    FitMergedCell ws.Range("C8")   ' date value
    FitMergedCell ws.Range("E8")   ' area value
    FitMergedCell ws.Range("G8")   ' day status value

    ' Table columns: expand to fit data, never shrink below current (template) width
    For Each lo In ws.ListObjects
        For Each col In lo.ListColumns
            If Not col.Range Is Nothing Then
                minW = col.Range.Columns(1).ColumnWidth
                col.Range.Columns(1).AutoFit
                needed = col.Range.Columns(1).ColumnWidth
                If needed < minW Then
                    col.Range.Columns(1).ColumnWidth = minW
                End If
            End If
        Next col
        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.Rows.EntireRow.AutoFit
        End If
    Next lo

    On Error GoTo 0

End Sub

Public Sub AutoFitWorksheet(ByVal ws As Worksheet)
    FitSheetContent ws
End Sub

Public Sub AutoFitWorkbook(ByVal wb As Workbook)

    Dim ws As Worksheet

    On Error Resume Next
    wb.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    For Each ws In wb.Worksheets
        If ws.Name <> "HiddenLists" Then
            On Error Resume Next
            ws.Unprotect Password:=SHEET_PASSWORD
            On Error GoTo 0
            FitSheetContent ws
        End If
    Next ws

    ' Re-apply locks + protection after fitting
    SecureWorkbook wb

End Sub

' Lock whole sheet; unlock Area, Day Status, and ALL table body cells.
' IMPORTANT: Excel blocks Insert/Delete Row if ANY cell in that row is locked
' (e.g. Remaining Amount formula). So entire table body must be unlocked.
Public Sub ApplyEditableCellLocks(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim isKhata As Boolean

    On Error Resume Next

    isKhata = IsKhataSheet(ws)
    ws.Cells.Locked = True

    ' Recovery sheet: Area/Day Status editable
    If Not isKhata Then
        ws.Range("E8").MergeArea.Locked = False   ' Area
        ws.Range("G8").MergeArea.Locked = False   ' Day Status
    End If

    For Each lo In ws.ListObjects
        If Not lo.HeaderRowRange Is Nothing Then
            lo.HeaderRowRange.Locked = True
        End If
        If Not lo.DataBodyRange Is Nothing Then
            ' Unlock FULL worksheet rows — Excel greys Insert/Delete Table Rows
            ' if any cell on that row (even outside the table) is locked.
            lo.DataBodyRange.EntireRow.Locked = False
            ' Row immediately below table (resize / insert at bottom)
            lo.DataBodyRange.Rows(lo.DataBodyRange.Rows.Count).Offset(1, 0).EntireRow.Locked = False
        End If

        ' Keep formula columns locked + ignore green indicators
        On Error Resume Next
        If isKhata Then
            If Not lo.ListColumns("Balance").DataBodyRange Is Nothing Then
                lo.ListColumns("Balance").DataBodyRange.Locked = True
                IgnoreRangeErrors lo.ListColumns("Balance").DataBodyRange
            End If
        Else
            If Not lo.ListColumns("Remaining Amount").DataBodyRange Is Nothing Then
                lo.ListColumns("Remaining Amount").DataBodyRange.Locked = True
                IgnoreRangeErrors lo.ListColumns("Remaining Amount").DataBodyRange
            End If
        End If
        On Error GoTo 0

        ' Totals row (if any) stay locked
        If Not lo.TotalsRowRange Is Nothing Then
            lo.TotalsRowRange.Locked = True
        End If
    Next lo

    ' Keep summary formulas visible in formula bar (esp. Recovered Amount B49)
    EnsureSummaryFormulasVisible ws

    On Error GoTo 0

End Sub

Private Function IsKhataSheet(ByVal ws As Worksheet) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    IsKhataSheet = False
    For Each lo In ws.ListObjects
        If InStr(1, lo.Name, "Khata", vbTextCompare) > 0 Then
            IsKhataSheet = True
            Exit Function
        End If
        If Not lo.ListColumns("Balance") Is Nothing Then
            IsKhataSheet = True
            Exit Function
        End If
    Next lo
    On Error GoTo 0
End Function

Private Sub IgnoreRangeErrors(ByVal rng As Range)
    Dim c As Range
    Dim i As Long
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    For Each c In rng.Cells
        For i = 1 To 9
            c.Errors.Item(i).Ignore = True
        Next i
    Next c
    On Error GoTo 0
End Sub

Private Sub EnsureSummaryFormulasVisible(ByVal ws As Worksheet)
    Dim addr As Variant
    Dim c As Range
    On Error Resume Next
    For Each addr In Array("B46", "B49", "B52", "B55", "B58", "E46", "E49", "E52", "E55", "E58")
        Set c = ws.Range(CStr(addr))
        If c.MergeCells Then
            c.MergeArea.FormulaHidden = False
        Else
            c.FormulaHidden = False
        End If
    Next addr
    On Error GoTo 0
End Sub

Public Sub ProtectSheetAllowTableEdit(ByVal ws As Worksheet)

    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    ApplyEditableCellLocks ws

    ' Sheet stays protected (formulas outside table / headers safe).
    ' Insert/Delete rows explicitly allowed for table use.
    ws.Protect _
        Password:=SHEET_PASSWORD, _
        UserInterfaceOnly:=True, _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=True, _
        AllowInsertingRows:=True, _
        AllowDeletingColumns:=False, _
        AllowDeletingRows:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True

    On Error Resume Next
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0

End Sub

Public Sub SecureWorkbook(ByVal wb As Workbook)

    Dim ws As Worksheet

    On Error Resume Next
    wb.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    For Each ws In wb.Worksheets
        If ws.Name <> "HiddenLists" Then
            ProtectSheetAllowTableEdit ws
        End If
    Next ws

    wb.Protect Password:=SHEET_PASSWORD, Structure:=True

End Sub
