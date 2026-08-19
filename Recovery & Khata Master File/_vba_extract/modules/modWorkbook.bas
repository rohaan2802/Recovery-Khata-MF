Attribute VB_Name = "modWorkbook"
Option Explicit

Public Function GetWorkbookName(ByVal monthName As String, _
                                ByVal YearName As String, _
                                ByVal EmployeeName As String) As String

    GetWorkbookName = "Recovery Register - " & _
                      monthName & " " & _
                      YearName & " - " & _
                      EmployeeName & ".xlsm"

End Function

Public Function GetKhataWorkbookName(ByVal monthName As String, _
                                     ByVal YearName As String, _
                                     ByVal EmployeeName As String) As String

    GetKhataWorkbookName = "Khata Register - " & _
                           monthName & " " & _
                           YearName & " - " & _
                           EmployeeName & ".xlsm"

End Function

Public Sub CreateRecoveryWorkbook()

    Dim RootFolder As String
    Dim selectedYear As String
    Dim selectedMonth As String
    Dim selectedEmployee As String
    Dim monthFolderName As String
    Dim employeeFolder As String
    Dim workbookName As String
    Dim WorkbookPath As String
    Dim legacyPath As String
    Dim wbNew As Workbook
    Dim templateSheet As Worksheet
    Dim MonthNo As Long
    Dim YearNo As Long
    Dim oldCalc As XlCalculation
    Dim oldSecurity As Long

    oldCalc = Application.Calculation

    On Error Resume Next
    oldSecurity = Application.AutomationSecurity
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    Application.Cursor = xlWait

    On Error Resume Next
    Application.PrintCommunication = False
    Application.AutomationSecurity = 3
    On Error GoTo ErrorHandler

    RootFolder = Trim(Range("txtRootFolder").Value)
    selectedYear = Trim(Range("selYear").Value)
    selectedMonth = Trim(Range("selMonth").Value)
    selectedEmployee = Trim(Range("selEmployee").Value)

    If RootFolder = "" Then
        MsgBox "Please select Root Folder.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    If selectedYear = "" Then
        MsgBox "Please select Year.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    If selectedMonth = "" Then
        MsgBox "Please select Month.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    If selectedEmployee = "" Then
        MsgBox "Please select Employee.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    EnsureTrustedLocation RootFolder

    monthFolderName = GetMonthFolderName(selectedMonth)

    If monthFolderName = "" Then
        MsgBox "Selected month was not found in Settings.", vbCritical, "Recovery Manager"
        GoTo ExitRoutine
    End If

    employeeFolder = RootFolder & "\" & selectedYear & "\" & monthFolderName & "\" & selectedEmployee

    If Dir(employeeFolder, vbDirectory) = "" Then
        MsgBox "Employee folder does not exist." & vbCrLf & _
               "Please create Employee Folder first.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    EnsureTrustedLocation employeeFolder

    workbookName = GetWorkbookName(selectedMonth, selectedYear, selectedEmployee)
    WorkbookPath = employeeFolder & "\" & workbookName
    legacyPath = employeeFolder & "\Recovery Register - " & selectedMonth & " " & _
                 selectedYear & " - " & selectedEmployee & ".xlsx"

    If Dir(WorkbookPath) <> "" Or Dir(legacyPath) <> "" Then
        MsgBox "Workbook already exists.", vbInformation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    YearNo = CLng(selectedYear)
    MonthNo = Month(DateValue("1 " & selectedMonth & " " & selectedYear))

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    ThisWorkbook.Worksheets("Recovery Template").Copy After:=wbNew.Worksheets(1)

    ' Keep explicit reference â€” CreateHiddenListsSheet adds a sheet at the end,
    ' so Worksheets.Count must NOT be used as the day-template.
    Set templateSheet = wbNew.Worksheets(wbNew.Worksheets.Count)

    CreateHiddenListsSheet wbNew
    CreateNamedRanges wbNew
    ApplyDropdownsToSheet templateSheet

    GenerateWorkingDaySheets wbNew, templateSheet, YearNo, MonthNo, selectedEmployee

    Application.DisplayAlerts = False
    On Error Resume Next
    wbNew.Worksheets("Sheet1").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    InstallAutoOpenSheetCode wbNew, MonthNo, YearNo
    ActivateStartupSheet wbNew, MonthNo, YearNo
    EnsureWorkbookWindowsVisible wbNew

    wbNew.SaveAs FileName:=WorkbookPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled

    SecureWorkbook wbNew
    EnsureWorkbookWindowsVisible wbNew
    wbNew.Save
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

    MsgBox "Workbook created successfully.", vbInformation, "Recovery Manager"

ExitRoutine:
    On Error Resume Next
    Application.Cursor = xlDefault
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    Application.AutomationSecurity = oldSecurity
    ThisWorkbook.Activate
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, "Recovery Manager"
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    On Error GoTo 0
    Resume ExitRoutine

End Sub

Public Sub CreateKhataWorkbook()

    Dim RootFolder As String
    Dim selectedYear As String
    Dim selectedMonth As String
    Dim selectedEmployee As String
    Dim monthFolderName As String
    Dim employeeFolder As String
    Dim workbookName As String
    Dim WorkbookPath As String
    Dim legacyPath As String
    Dim wbNew As Workbook
    Dim templateSheet As Worksheet
    Dim oldCalc As XlCalculation
    Dim oldSecurity As Long

    oldCalc = Application.Calculation

    On Error Resume Next
    oldSecurity = Application.AutomationSecurity
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    Application.Cursor = xlWait

    On Error Resume Next
    Application.PrintCommunication = False
    Application.AutomationSecurity = 3
    On Error GoTo ErrorHandler

    RootFolder = Trim(Range("txtRootFolder").Value)
    selectedYear = Trim(Range("selYear").Value)
    selectedMonth = Trim(Range("selMonth").Value)
    selectedEmployee = Trim(Range("selEmployee").Value)

    If RootFolder = "" Then
        MsgBox "Please select Root Folder.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    If selectedYear = "" Then
        MsgBox "Please select Year.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    If selectedMonth = "" Then
        MsgBox "Please select Month.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    If selectedEmployee = "" Then
        MsgBox "Please select Employee.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    EnsureTrustedLocation RootFolder

    monthFolderName = GetMonthFolderName(selectedMonth)

    If monthFolderName = "" Then
        MsgBox "Selected month was not found in Settings.", vbCritical, "Recovery Manager"
        GoTo ExitRoutine
    End If

    employeeFolder = RootFolder & "\" & selectedYear & "\" & monthFolderName & "\" & selectedEmployee

    If Dir(employeeFolder, vbDirectory) = "" Then
        MsgBox "Employee folder does not exist." & vbCrLf & _
               "Please create Employee Folder first.", vbExclamation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    EnsureTrustedLocation employeeFolder

    workbookName = GetKhataWorkbookName(selectedMonth, selectedYear, selectedEmployee)
    WorkbookPath = employeeFolder & "\" & workbookName
    legacyPath = employeeFolder & "\Khata Register - " & selectedMonth & " " & _
                 selectedYear & " - " & selectedEmployee & ".xlsx"

    If Dir(WorkbookPath) <> "" Or Dir(legacyPath) <> "" Then
        MsgBox "Khata workbook already exists.", vbInformation, "Recovery Manager"
        GoTo ExitRoutine
    End If

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    ThisWorkbook.Worksheets("Khata Template").Copy After:=wbNew.Worksheets(1)

    Set templateSheet = wbNew.Worksheets(wbNew.Worksheets.Count)

    ' Fill header boxes from Home selections
    templateSheet.Range("A8").Value = selectedEmployee
    templateSheet.Range("C8").MergeArea.NumberFormat = "@"
    templateSheet.Range("C8").MergeArea.Value = selectedMonth & ", " & selectedYear

    ' Opening balance: 0 if first Khata; else prior Khata closing (any earlier month/year)
    templateSheet.Range("E8").MergeArea.NumberFormat = "#,##0"
    templateSheet.Range("E8").MergeArea.Value = GetPreviousKhataClosingBalance( _
        RootFolder, selectedEmployee, selectedMonth, CLng(selectedYear))

    On Error Resume Next
    FitMergedCell templateSheet.Range("A8")
    FitMergedCell templateSheet.Range("C8")
    FitMergedCell templateSheet.Range("E8")
    On Error GoTo ErrorHandler

    FixKhataSummaryFormulas templateSheet

    On Error Resume Next
    templateSheet.Name = Left(CleanSheetName(selectedEmployee & " " & selectedMonth & " Khata"), 31)
    If Err.Number <> 0 Then
        Err.Clear
        templateSheet.Name = Left(CleanSheetName(selectedEmployee & " " & selectedMonth & " Khata"), 28) & "_" & Format(Now, "hhmm")
    End If
    On Error GoTo ErrorHandler

    Application.DisplayAlerts = False
    On Error Resume Next
    wbNew.Worksheets("Sheet1").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    EnsureWorkbookWindowsVisible wbNew

    ' Events off: avoid AfterSave cascade during create (was leaving book open / no MsgBox)
    Application.EnableEvents = False
    On Error Resume Next
    InstallKhataCascadeCode wbNew
    Err.Clear
    On Error GoTo ErrorHandler

    wbNew.SaveAs FileName:=WorkbookPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    SecureWorkbook wbNew
    wbNew.Close SaveChanges:=True
    Set wbNew = Nothing

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    MsgBox "Khata workbook created successfully.", vbInformation, "Recovery Manager"

ExitRoutine:
    On Error Resume Next
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    Application.AutomationSecurity = oldSecurity
    ThisWorkbook.Activate
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, "Recovery Manager"
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    On Error GoTo 0
    Resume ExitRoutine

End Sub

Private Function CleanSheetName(ByVal rawName As String) As String
    Dim s As String
    s = Trim(rawName)
    s = Replace(s, "\", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "?", " ")
    s = Replace(s, "*", " ")
    s = Replace(s, "[", " ")
    s = Replace(s, "]", " ")
    s = Replace(s, ":", " ")
    If Len(s) = 0 Then s = "Khata"
    CleanSheetName = s
End Function

' Balance formula on template is correct:
' Opening(E8) + Debits to date - Credits to date. Closing = E8 + Total Debit - Total Credit.
Private Sub FixKhataSummaryFormulas(ByVal ws As Worksheet)
    Dim lo As ListObject
    Dim t As String

    On Error Resume Next
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    t = lo.Name
    If Len(t) = 0 Then Exit Sub

    ws.Range("B31").Formula = "=SUM(" & t & "[Debit])"
    ws.Range("B34").Formula = "=SUM(" & t & "[Credit])"
    ws.Range("B37").Formula = "=$E$8+SUM(" & t & "[Debit])-SUM(" & t & "[Credit])"
    ws.Range("B31,B34,B37,E8").NumberFormat = "#,##0"
    On Error GoTo 0
End Sub

' Home button / manual: refresh openings for selected employee from prior closings
Public Sub RefreshKhataOpeningBalances()
    Dim RootFolder As String
    Dim selectedEmployee As String

    RootFolder = Trim(Range("txtRootFolder").Value)
    selectedEmployee = Trim(Range("selEmployee").Value)

    If RootFolder = "" Then
        MsgBox "Please select Root Folder.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If
    If selectedEmployee = "" Then
        MsgBox "Please select Employee.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    On Error GoTo RefreshFail

    RebuildKhataOpeningChain RootFolder, selectedEmployee

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox "Khata opening balances updated for " & selectedEmployee & ".", vbInformation, "Recovery Manager"
    Exit Sub

RefreshFail:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "Recovery Manager"
End Sub

' Sort all Khata files for employee by period; set each opening from previous closing.
Public Sub RebuildKhataOpeningChain(ByVal RootFolder As String, ByVal EmployeeName As String)
    Dim files As Collection
    Dim i As Long
    Dim filePath As String
    Dim prevClose As Double
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim alreadyOpen As Boolean
    Dim oldEvents As Boolean
    Dim oldSec As Long
    Dim oldCalc As XlCalculation
    Dim parts() As String

    On Error Resume Next
    Set files = CollectEmployeeKhataFiles(RootFolder, EmployeeName)
    If files Is Nothing Then Exit Sub
    If files.Count = 0 Then Exit Sub
    SortKeyPathCollection files

    oldEvents = Application.EnableEvents
    oldSec = Application.AutomationSecurity
    oldCalc = Application.Calculation
    Application.EnableEvents = False
    Application.AutomationSecurity = 3
    Application.Calculation = xlCalculationAutomatic

    prevClose = 0
    For i = 1 To files.Count
        parts = Split(CStr(files(i)), "|")
        If UBound(parts) >= 1 Then
            filePath = parts(1)
            alreadyOpen = False
            Set wb = Nothing
            Set wb = GetOpenWorkbookByPath(filePath)
            If Not wb Is Nothing Then
                alreadyOpen = True
            Else
                Set wb = Workbooks.Open(FileName:=filePath, UpdateLinks:=0)
            End If

            If Not wb Is Nothing Then
                Set ws = wb.Worksheets(1)
                FixKhataSummaryFormulas ws

                If i = 1 Then
                    ' Chronologically first Khata always starts at 0
                    ws.Range("E8").MergeArea.NumberFormat = "#,##0"
                    ws.Range("E8").MergeArea.Value = 0
                Else
                    ws.Range("E8").MergeArea.NumberFormat = "#,##0"
                    ws.Range("E8").MergeArea.Value = prevClose
                End If

                Application.CalculateFull
                If IsNumeric(ws.Range("B37").Value2) Then
                    prevClose = CDbl(ws.Range("B37").Value2)
                Else
                    prevClose = 0
                End If

                ' Upgrade Open/Save sync macros (events already off)
                InstallKhataCascadeCode wb

                wb.Save
                If Not alreadyOpen Then wb.Close SaveChanges:=False
            End If
        End If
    Next i

    Application.EnableEvents = oldEvents
    Application.AutomationSecurity = oldSec
    Application.Calculation = oldCalc
    On Error GoTo 0
End Sub

Private Function GetOpenWorkbookByPath(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb
    Set GetOpenWorkbookByPath = Nothing
End Function

' Returns Collection of "0000245|fullpath" sorted later by SortKeyPathCollection
Private Function CollectEmployeeKhataFiles(ByVal RootFolder As String, _
                                          ByVal EmployeeName As String) As Collection
    Dim FSO As Object
    Dim yearFolder As Object
    Dim monthFolder As Object
    Dim empFolder As Object
    Dim File As Object
    Dim col As Collection
    Dim y As Long
    Dim m As Long
    Dim fileKey As Long

    Set col = New Collection
    Set CollectEmployeeKhataFiles = col

    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO Is Nothing Then Exit Function
    If Not FSO.FolderExists(RootFolder) Then Exit Function

    For Each yearFolder In FSO.GetFolder(RootFolder).SubFolders
        If IsNumeric(yearFolder.Name) Then
            For Each monthFolder In yearFolder.SubFolders
                For Each empFolder In monthFolder.SubFolders
                    If StrComp(empFolder.Name, EmployeeName, vbTextCompare) = 0 Then
                        For Each File In empFolder.Files
                            If Left(File.Name, 2) <> "~$" Then
                                If InStr(1, File.Name, "Khata Register", vbTextCompare) > 0 Then
                                    m = 0
                                    y = 0
                                    If Not ParseKhataFilePeriod(File.Name, m, y) Then
                                        ParseKhataPathPeriod monthFolder.Name, yearFolder.Name, m, y
                                    End If
                                    If m >= 1 And m <= 12 And y > 1900 Then
                                        fileKey = y * 12 + m
                                        col.Add Format(fileKey, "000000") & "|" & File.Path
                                    End If
                                End If
                            End If
                        Next File
                    End If
                Next empFolder
            Next monthFolder
        End If
    Next yearFolder
    On Error GoTo 0
End Function

Private Sub SortKeyPathCollection(ByRef col As Collection)
    Dim arr() As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    Dim n As Long

    n = col.Count
    If n <= 1 Then Exit Sub

    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CStr(col(i))
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    Do While col.Count > 0
        col.Remove 1
    Loop
    For i = 1 To n
        col.Add arr(i)
    Next i
End Sub

' Inject Open/AfterSave so openings follow latest prior-month closing
' (May filled later → July picks it up when July opens, or when May is saved)
Private Sub InstallKhataCascadeCode(ByVal wb As Workbook)
    Dim cm As Object
    Dim code As String

    On Error GoTo InstallFail

    Set cm = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    code = "Private Const SHEET_PW As String = ""2968""" & vbCrLf & vbCrLf
    code = code & "Private Sub Workbook_Open()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Application.EnableEvents = True" & vbCrLf
    code = code & "    SyncThisKhataOpeningFromPrior" & vbCrLf
    code = code & "    ApplyProtectionForEditing" & vbCrLf
    code = code & "    ApplyKhataTableBordersInSheet Worksheets(1)" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub Workbook_AfterSave(ByVal Success As Boolean)" & vbCrLf
    code = code & "    If Not Success Then Exit Sub" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    SyncAllKhataOpeningsForEmployee" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    code = code & "    Dim lo As ListObject, zone As Range, inTable As Boolean" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    If Sh Is Nothing Or Target Is Nothing Then Exit Sub" & vbCrLf
    code = code & "    inTable = False" & vbCrLf
    code = code & "    For Each lo In Sh.ListObjects" & vbCrLf
    code = code & "        If Not lo.Range Is Nothing Then" & vbCrLf
    code = code & "            Set zone = lo.Range.Resize(lo.Range.Rows.Count + 1, lo.Range.Columns.Count)" & vbCrLf
    code = code & "            If Not Intersect(Target, zone) Is Nothing Then inTable = True: Exit For" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "    If inTable Then" & vbCrLf
    code = code & "        If Sh.ProtectContents Then Sh.Unprotect Password:=SHEET_PW" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        If Not Sh.ProtectContents Then ProtectKhataSheet Sh" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    code = code & "    Dim lo As ListObject, dCol As ListColumn, cCol As ListColumn, bCol As ListColumn" & vbCrLf
    code = code & "    Dim rng As Range, cell As Range, f As String, i As Long" & vbCrLf
    code = code & "    On Error GoTo CleanExit" & vbCrLf
    code = code & "    If Sh Is Nothing Or Target Is Nothing Then GoTo CleanExit" & vbCrLf
    code = code & "    Application.EnableEvents = False" & vbCrLf
    code = code & "    f = ""=IF(AND([@Debit]="""" ,[@Credit]=""""),"""",SUM($E$8,INDEX([Debit],1):[@Debit])-SUM(INDEX([Credit],1):[@Credit]))""" & vbCrLf
    code = code & "    For Each lo In Sh.ListObjects" & vbCrLf
    code = code & "        Set dCol = Nothing: Set cCol = Nothing: Set bCol = Nothing" & vbCrLf
    code = code & "        On Error Resume Next" & vbCrLf
    code = code & "        Set dCol = lo.ListColumns(""Debit"")" & vbCrLf
    code = code & "        Set cCol = lo.ListColumns(""Credit"")" & vbCrLf
    code = code & "        Set bCol = lo.ListColumns(""Balance"")" & vbCrLf
    code = code & "        On Error GoTo CleanExit" & vbCrLf
    code = code & "        If Not bCol Is Nothing Then" & vbCrLf
    code = code & "            If Not bCol.DataBodyRange Is Nothing Then" & vbCrLf
    code = code & "                Set rng = Intersect(Target, Union(bCol.DataBodyRange, dCol.DataBodyRange, cCol.DataBodyRange))" & vbCrLf
    code = code & "                If Not rng Is Nothing Then" & vbCrLf
    code = code & "                    bCol.DataBodyRange.Formula = f" & vbCrLf
    code = code & "                    bCol.DataBodyRange.Locked = True" & vbCrLf
    code = code & "                    For Each cell In bCol.DataBodyRange.Cells" & vbCrLf
    code = code & "                        For i = 1 To 9" & vbCrLf
    code = code & "                            cell.Errors.Item(i).Ignore = True" & vbCrLf
    code = code & "                        Next i" & vbCrLf
    code = code & "                    Next cell" & vbCrLf
    code = code & "                    ApplyKhataTableBordersInSheet Sh" & vbCrLf
    code = code & "                End If" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "CleanExit:" & vbCrLf
    code = code & "    Application.EnableEvents = True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub ApplyProtectionForEditing()" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Me.Unprotect Password:=SHEET_PW" & vbCrLf
    code = code & "    For Each ws In Me.Worksheets" & vbCrLf
    code = code & "        ProtectKhataSheet ws" & vbCrLf
    code = code & "    Next ws" & vbCrLf
    code = code & "    Me.Protect Password:=SHEET_PW, Structure:=True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub ProtectKhataSheet(ByVal ws As Worksheet)" & vbCrLf
    code = code & "    Dim lo As ListObject, bCol As ListColumn, cell As Range, i As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    ws.Unprotect Password:=SHEET_PW" & vbCrLf
    code = code & "    ws.Cells.Locked = True" & vbCrLf
    code = code & "    For Each lo In ws.ListObjects" & vbCrLf
    code = code & "        If Not lo.HeaderRowRange Is Nothing Then lo.HeaderRowRange.Locked = True" & vbCrLf
    code = code & "        If Not lo.DataBodyRange Is Nothing Then" & vbCrLf
    code = code & "            lo.DataBodyRange.Locked = False" & vbCrLf
    code = code & "            lo.DataBodyRange.EntireRow.Locked = False" & vbCrLf
    code = code & "            lo.DataBodyRange.Rows(lo.DataBodyRange.Rows.Count).Offset(1, 0).EntireRow.Locked = False" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "        Set bCol = Nothing" & vbCrLf
    code = code & "        Set bCol = lo.ListColumns(""Balance"")" & vbCrLf
    code = code & "        If Not bCol Is Nothing Then" & vbCrLf
    code = code & "            If Not bCol.DataBodyRange Is Nothing Then" & vbCrLf
    code = code & "                bCol.DataBodyRange.Formula = ""=IF(AND([@Debit]="""" ,[@Credit]=""""),"""",SUM($E$8,INDEX([Debit],1):[@Debit])-SUM(INDEX([Credit],1):[@Credit]))""" & vbCrLf
    code = code & "                bCol.DataBodyRange.Locked = True" & vbCrLf
    code = code & "                For Each cell In bCol.DataBodyRange.Cells" & vbCrLf
    code = code & "                    For i = 1 To 9" & vbCrLf
    code = code & "                        cell.Errors.Item(i).Ignore = True" & vbCrLf
    code = code & "                    Next i" & vbCrLf
    code = code & "                Next cell" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "    ApplyKhataTableBordersInSheet ws" & vbCrLf
    code = code & "    ws.Protect Password:=SHEET_PW, UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub ApplyKhataTableBordersInSheet(ByVal ws As Worksheet)" & vbCrLf
    code = code & "    Dim lo As ListObject" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    For Each lo In ws.ListObjects" & vbCrLf
    code = code & "        With lo.Range.Borders" & vbCrLf
    code = code & "            .LineStyle = xlContinuous" & vbCrLf
    code = code & "            .Color = RGB(0, 0, 0)" & vbCrLf
    code = code & "            .Weight = xlThin" & vbCrLf
    code = code & "        End With" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub SyncThisKhataOpeningFromPrior()" & vbCrLf
    code = code & "    Dim root As String, emp As String, myKey As Long, priorPath As String" & vbCrLf
    code = code & "    Dim priorClose As Double, oldEvents As Boolean" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    KhataRootAndEmployee root, emp" & vbCrLf
    code = code & "    myKey = KhataPeriodKeyFromFile(ThisWorkbook.Name)" & vbCrLf
    code = code & "    If myKey <= 0 Then myKey = KhataPeriodKeyFromSheet(Worksheets(1))" & vbCrLf
    code = code & "    If myKey <= 0 Then Exit Sub" & vbCrLf
    code = code & "    priorPath = FindPriorKhataPath(root, emp, myKey)" & vbCrLf
    code = code & "    oldEvents = Application.EnableEvents" & vbCrLf
    code = code & "    Application.EnableEvents = False" & vbCrLf
    code = code & "    Worksheets(1).Range(""E8"").MergeArea.NumberFormat = ""#,##0""" & vbCrLf
    code = code & "    If Len(priorPath) = 0 Then" & vbCrLf
    code = code & "        Worksheets(1).Range(""E8"").MergeArea.Value = 0" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        priorClose = ReadKhataClosingBalance(priorPath)" & vbCrLf
    code = code & "        Worksheets(1).Range(""E8"").MergeArea.Value = priorClose" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Application.Calculate" & vbCrLf
    code = code & "    Application.EnableEvents = oldEvents" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub SyncAllKhataOpeningsForEmployee()" & vbCrLf
    code = code & "    Dim root As String, emp As String, oldEvents As Boolean" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    oldEvents = Application.EnableEvents" & vbCrLf
    code = code & "    Application.EnableEvents = False" & vbCrLf
    code = code & "    KhataRootAndEmployee root, emp" & vbCrLf
    code = code & "    Err.Clear" & vbCrLf
    code = code & "    Application.Run ""'Recovery Manager.xlsm'!RebuildKhataOpeningChain"", root, emp" & vbCrLf
    code = code & "    If Err.Number <> 0 Then" & vbCrLf
    code = code & "        Err.Clear" & vbCrLf
    code = code & "        LocalSyncAllKhataOpenings root, emp" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Application.EnableEvents = oldEvents" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub KhataRootAndEmployee(ByRef root As String, ByRef emp As String)" & vbCrLf
    code = code & "    Dim FSO As Object, monthFolder As String, yearFolder As String" & vbCrLf
    code = code & "    Set FSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "    emp = FSO.GetFolder(ThisWorkbook.Path).Name" & vbCrLf
    code = code & "    monthFolder = FSO.GetParentFolderName(ThisWorkbook.Path)" & vbCrLf
    code = code & "    yearFolder = FSO.GetParentFolderName(monthFolder)" & vbCrLf
    code = code & "    root = FSO.GetParentFolderName(yearFolder)" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Function KhataPeriodKeyFromFile(ByVal fileName As String) As Long" & vbCrLf
    code = code & "    Dim baseName As String, parts() As String, midPart As String, tokens() As String" & vbCrLf
    code = code & "    Dim monName As String, y As Long, m As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    KhataPeriodKeyFromFile = 0" & vbCrLf
    code = code & "    baseName = fileName" & vbCrLf
    code = code & "    If InStrRev(baseName, ""."") > 0 Then baseName = Left(baseName, InStrRev(baseName, ""."") - 1)" & vbCrLf
    code = code & "    parts = Split(baseName, "" - "")" & vbCrLf
    code = code & "    If UBound(parts) < 2 Then Exit Function" & vbCrLf
    code = code & "    midPart = Replace(Trim(parts(1)), "","", """")" & vbCrLf
    code = code & "    tokens = Split(Application.WorksheetFunction.Trim(midPart), "" "")" & vbCrLf
    code = code & "    If UBound(tokens) < 1 Then Exit Function" & vbCrLf
    code = code & "    monName = Trim(tokens(0))" & vbCrLf
    code = code & "    y = CLng(Trim(tokens(UBound(tokens))))" & vbCrLf
    code = code & "    m = Month(DateValue(""1 "" & monName & "" "" & CStr(y)))" & vbCrLf
    code = code & "    If m >= 1 And y > 1900 Then KhataPeriodKeyFromFile = y * 12 + m" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function KhataPeriodKeyFromSheet(ByVal ws As Worksheet) As Long" & vbCrLf
    code = code & "    Dim txt As String, tokens() As String, monName As String, y As Long, m As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    KhataPeriodKeyFromSheet = 0" & vbCrLf
    code = code & "    txt = Replace(CStr(ws.Range(""C8"").MergeArea.Value), "","", """")" & vbCrLf
    code = code & "    tokens = Split(Application.WorksheetFunction.Trim(txt), "" "")" & vbCrLf
    code = code & "    If UBound(tokens) < 1 Then Exit Function" & vbCrLf
    code = code & "    monName = Trim(tokens(0))" & vbCrLf
    code = code & "    y = CLng(Trim(tokens(UBound(tokens))))" & vbCrLf
    code = code & "    m = Month(DateValue(""1 "" & monName & "" "" & CStr(y)))" & vbCrLf
    code = code & "    If m >= 1 And y > 1900 Then KhataPeriodKeyFromSheet = y * 12 + m" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function FindPriorKhataPath(ByVal root As String, ByVal emp As String, ByVal myKey As Long) As String" & vbCrLf
    code = code & "    Dim FSO As Object, yf As Object, mf As Object, ef As Object, f As Object" & vbCrLf
    code = code & "    Dim bestKey As Long, fileKey As Long, bestPath As String" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    FindPriorKhataPath = """"" & vbCrLf
    code = code & "    bestKey = -1" & vbCrLf
    code = code & "    Set FSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "    For Each yf In FSO.GetFolder(root).SubFolders" & vbCrLf
    code = code & "        If IsNumeric(yf.Name) Then" & vbCrLf
    code = code & "            For Each mf In yf.SubFolders" & vbCrLf
    code = code & "                For Each ef In mf.SubFolders" & vbCrLf
    code = code & "                    If StrComp(ef.Name, emp, vbTextCompare) = 0 Then" & vbCrLf
    code = code & "                        For Each f In ef.Files" & vbCrLf
    code = code & "                            If Left(f.Name, 2) <> ""~$"" And InStr(1, f.Name, ""Khata Register"", vbTextCompare) > 0 Then" & vbCrLf
    code = code & "                                fileKey = KhataPeriodKeyFromFile(f.Name)" & vbCrLf
    code = code & "                                If fileKey <= 0 Then fileKey = KhataPeriodKeyFromFolder(mf.Name, yf.Name)" & vbCrLf
    code = code & "                                If fileKey > 0 And fileKey < myKey And fileKey > bestKey Then" & vbCrLf
    code = code & "                                    bestKey = fileKey: bestPath = f.Path" & vbCrLf
    code = code & "                                End If" & vbCrLf
    code = code & "                            End If" & vbCrLf
    code = code & "                        Next f" & vbCrLf
    code = code & "                    End If" & vbCrLf
    code = code & "                Next ef" & vbCrLf
    code = code & "            Next mf" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next yf" & vbCrLf
    code = code & "    FindPriorKhataPath = bestPath" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function KhataPeriodKeyFromFolder(ByVal monthFolderName As String, ByVal yearFolderName As String) As Long" & vbCrLf
    code = code & "    Dim p As Long, monName As String, y As Long, m As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    KhataPeriodKeyFromFolder = 0" & vbCrLf
    code = code & "    If Not IsNumeric(yearFolderName) Then Exit Function" & vbCrLf
    code = code & "    y = CLng(yearFolderName)" & vbCrLf
    code = code & "    p = InStr(monthFolderName, ""."")" & vbCrLf
    code = code & "    If p > 0 Then" & vbCrLf
    code = code & "        monName = Trim(Mid(monthFolderName, p + 1))" & vbCrLf
    code = code & "        m = Month(DateValue(""1 "" & monName & "" "" & CStr(y)))" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        m = CLng(Val(monthFolderName))" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    If m >= 1 And m <= 12 Then KhataPeriodKeyFromFolder = y * 12 + m" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function ReadKhataClosingBalance(ByVal filePath As String) As Double" & vbCrLf
    code = code & "    Dim wb As Workbook, ws As Worksheet, already As Boolean, wbCheck As Workbook" & vbCrLf
    code = code & "    Dim oldSec As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    ReadKhataClosingBalance = 0" & vbCrLf
    code = code & "    already = False: Set wb = Nothing" & vbCrLf
    code = code & "    For Each wbCheck In Application.Workbooks" & vbCrLf
    code = code & "        If StrComp(wbCheck.FullName, filePath, vbTextCompare) = 0 Then Set wb = wbCheck: already = True: Exit For" & vbCrLf
    code = code & "    Next wbCheck" & vbCrLf
    code = code & "    If Not already Then" & vbCrLf
    code = code & "        oldSec = Application.AutomationSecurity" & vbCrLf
    code = code & "        Application.AutomationSecurity = 3" & vbCrLf
    code = code & "        Set wb = Workbooks.Open(FileName:=filePath, ReadOnly:=True, UpdateLinks:=0)" & vbCrLf
    code = code & "        Application.AutomationSecurity = oldSec" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    If wb Is Nothing Then Exit Function" & vbCrLf
    code = code & "    Set ws = wb.Worksheets(1)" & vbCrLf
    code = code & "    Application.Calculate" & vbCrLf
    code = code & "    If IsNumeric(ws.Range(""B37"").Value2) Then ReadKhataClosingBalance = CDbl(ws.Range(""B37"").Value2)" & vbCrLf
    code = code & "    If Not already Then wb.Close SaveChanges:=False" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Sub LocalSyncAllKhataOpenings(ByVal root As String, ByVal emp As String)" & vbCrLf
    code = code & "    Dim FSO As Object, yf As Object, mf As Object, ef As Object, f As Object" & vbCrLf
    code = code & "    Dim col As New Collection, arr() As String, i As Long, j As Long, tmp As String, n As Long" & vbCrLf
    code = code & "    Dim fileKey As Long, p As String, prev As Double" & vbCrLf
    code = code & "    Dim wb As Workbook, ws As Worksheet, already As Boolean, wbCheck As Workbook" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Set FSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "    For Each yf In FSO.GetFolder(root).SubFolders" & vbCrLf
    code = code & "        If IsNumeric(yf.Name) Then" & vbCrLf
    code = code & "            For Each mf In yf.SubFolders" & vbCrLf
    code = code & "                For Each ef In mf.SubFolders" & vbCrLf
    code = code & "                    If StrComp(ef.Name, emp, vbTextCompare) = 0 Then" & vbCrLf
    code = code & "                        For Each f In ef.Files" & vbCrLf
    code = code & "                            If Left(f.Name, 2) <> ""~$"" And InStr(1, f.Name, ""Khata Register"", vbTextCompare) > 0 Then" & vbCrLf
    code = code & "                                fileKey = KhataPeriodKeyFromFile(f.Name)" & vbCrLf
    code = code & "                                If fileKey <= 0 Then fileKey = KhataPeriodKeyFromFolder(mf.Name, yf.Name)" & vbCrLf
    code = code & "                                If fileKey > 0 Then col.Add Format(fileKey, ""000000"") & ""|"" & f.Path" & vbCrLf
    code = code & "                            End If" & vbCrLf
    code = code & "                        Next f" & vbCrLf
    code = code & "                    End If" & vbCrLf
    code = code & "                Next ef" & vbCrLf
    code = code & "            Next mf" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next yf" & vbCrLf
    code = code & "    n = col.Count: If n = 0 Then Exit Sub" & vbCrLf
    code = code & "    ReDim arr(1 To n)" & vbCrLf
    code = code & "    For i = 1 To n: arr(i) = CStr(col(i)): Next i" & vbCrLf
    code = code & "    For i = 1 To n - 1" & vbCrLf
    code = code & "        For j = i + 1 To n" & vbCrLf
    code = code & "            If arr(i) > arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp" & vbCrLf
    code = code & "        Next j" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    prev = 0" & vbCrLf
    code = code & "    For i = 1 To n" & vbCrLf
    code = code & "        p = Split(arr(i), ""|"")(1)" & vbCrLf
    code = code & "        already = False: Set wb = Nothing" & vbCrLf
    code = code & "        For Each wbCheck In Application.Workbooks" & vbCrLf
    code = code & "            If StrComp(wbCheck.FullName, p, vbTextCompare) = 0 Then Set wb = wbCheck: already = True: Exit For" & vbCrLf
    code = code & "        Next wbCheck" & vbCrLf
    code = code & "        If Not already Then Set wb = Workbooks.Open(FileName:=p, UpdateLinks:=0)" & vbCrLf
    code = code & "        If Not wb Is Nothing Then" & vbCrLf
    code = code & "            Set ws = wb.Worksheets(1)" & vbCrLf
    code = code & "            ws.Range(""E8"").MergeArea.NumberFormat = ""#,##0""" & vbCrLf
    code = code & "            If i = 1 Then" & vbCrLf
    code = code & "                ws.Range(""E8"").MergeArea.Value = 0" & vbCrLf
    code = code & "            Else" & vbCrLf
    code = code & "                ws.Range(""E8"").MergeArea.Value = prev" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "            Application.Calculate" & vbCrLf
    code = code & "            If IsNumeric(ws.Range(""B37"").Value2) Then prev = CDbl(ws.Range(""B37"").Value2) Else prev = 0" & vbCrLf
    code = code & "            If StrComp(wb.FullName, ThisWorkbook.FullName, vbTextCompare) <> 0 Then" & vbCrLf
    code = code & "                wb.Save" & vbCrLf
    code = code & "                If Not already Then wb.Close SaveChanges:=False" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "End Sub" & vbCrLf

    cm.AddFromString code
    Exit Sub

InstallFail:
End Sub

' Latest Khata Register for this employee with Year/Month earlier than current.
' Returns that file's Closing Balance (B37), or 0 if none.
Private Function GetPreviousKhataClosingBalance(ByVal RootFolder As String, _
                                                ByVal EmployeeName As String, _
                                                ByVal currentMonth As String, _
                                                ByVal currentYear As Long) As Double

    Dim FSO As Object
    Dim yearFolder As Object
    Dim monthFolder As Object
    Dim empFolder As Object
    Dim File As Object
    Dim bestPath As String
    Dim bestKey As Long
    Dim fileKey As Long
    Dim currentKey As Long
    Dim y As Long
    Dim m As Long
    Dim wbPrior As Workbook
    Dim ws As Worksheet
    Dim bal As Variant
    Dim oldSec As Long

    GetPreviousKhataClosingBalance = 0
    bestPath = ""
    bestKey = -1

    On Error Resume Next
    currentKey = currentYear * 12 + Month(DateValue("1 " & currentMonth & " " & CStr(currentYear)))
    If Err.Number <> 0 Or currentKey <= 0 Then Exit Function
    Err.Clear

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO Is Nothing Then Exit Function
    If Not FSO.FolderExists(RootFolder) Then Exit Function

    For Each yearFolder In FSO.GetFolder(RootFolder).SubFolders
        If IsNumeric(yearFolder.Name) Then
            For Each monthFolder In yearFolder.SubFolders
                For Each empFolder In monthFolder.SubFolders
                    If StrComp(empFolder.Name, EmployeeName, vbTextCompare) = 0 Then
                        For Each File In empFolder.Files
                            If Left(File.Name, 2) <> "~$" Then
                                If InStr(1, File.Name, "Khata Register", vbTextCompare) > 0 Then
                                    m = 0
                                    y = 0
                                    If Not ParseKhataFilePeriod(File.Name, m, y) Then
                                        ParseKhataPathPeriod monthFolder.Name, yearFolder.Name, m, y
                                    End If
                                    If m >= 1 And m <= 12 And y > 1900 Then
                                        fileKey = y * 12 + m
                                        If fileKey < currentKey And fileKey > bestKey Then
                                            bestKey = fileKey
                                            bestPath = File.Path
                                        End If
                                    End If
                                End If
                            End If
                        Next File
                    End If
                Next empFolder
            Next monthFolder
        End If
    Next yearFolder

    If Len(bestPath) = 0 Then Exit Function

    oldSec = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Set wbPrior = Nothing
    Set wbPrior = Workbooks.Open(FileName:=bestPath, ReadOnly:=True, UpdateLinks:=0)
    If Not wbPrior Is Nothing Then
        Application.CalculateFull
        Set ws = wbPrior.Worksheets(1)
        bal = ws.Range("B37").Value2
        If IsNumeric(bal) Then
            GetPreviousKhataClosingBalance = CDbl(bal)
        Else
            GetPreviousKhataClosingBalance = 0
        End If
        wbPrior.Close SaveChanges:=False
    End If
    Application.AutomationSecurity = oldSec
    On Error GoTo 0

End Function

' Parse "Khata Register - August 2026 - NAME.xlsm" -> monthNo, yearNo. Returns True on success.
Private Function ParseKhataFilePeriod(ByVal fileName As String, _
                                      ByRef monthNo As Long, _
                                      ByRef yearNo As Long) As Boolean
    Dim baseName As String
    Dim parts() As String
    Dim midPart As String
    Dim tokens() As String
    Dim monName As String

    ParseKhataFilePeriod = False
    monthNo = 0
    yearNo = 0

    On Error Resume Next
    baseName = fileName
    If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)

    parts = Split(baseName, " - ")
    If UBound(parts) < 2 Then Exit Function

    midPart = Trim(parts(1))   ' "August 2026" or "August, 2026"
    midPart = Replace(midPart, ",", "")
    tokens = Split(Application.WorksheetFunction.Trim(midPart), " ")
    If UBound(tokens) < 1 Then Exit Function

    monName = Trim(tokens(0))
    yearNo = CLng(Trim(tokens(UBound(tokens))))
    monthNo = Month(DateValue("1 " & monName & " " & CStr(yearNo)))
    If monthNo >= 1 And monthNo <= 12 And yearNo > 1900 Then ParseKhataFilePeriod = True
    On Error GoTo 0
End Function

' Folder "8. August" + year folder "2026"
Private Function ParseKhataPathPeriod(ByVal monthFolderName As String, _
                                      ByVal yearFolderName As String, _
                                      ByRef monthNo As Long, _
                                      ByRef yearNo As Long) As Boolean
    Dim p As Long
    Dim monName As String

    ParseKhataPathPeriod = False
    monthNo = 0
    yearNo = 0

    On Error Resume Next
    If Not IsNumeric(yearFolderName) Then Exit Function
    yearNo = CLng(yearFolderName)

    p = InStr(monthFolderName, ".")
    If p > 0 Then
        monName = Trim(Mid(monthFolderName, p + 1))
        monthNo = Month(DateValue("1 " & monName & " " & CStr(yearNo)))
    Else
        monthNo = CLng(Val(monthFolderName))
    End If

    If monthNo >= 1 And monthNo <= 12 Then ParseKhataPathPeriod = True
    On Error GoTo 0
End Function

Private Sub GenerateWorkingDaySheets(ByVal wb As Workbook, _
                                     ByVal seedSheet As Worksheet, _
                                     ByVal YearNo As Long, _
                                     ByVal MonthNo As Long, _
                                     ByVal EmployeeName As String)

    Dim NewSheet As Worksheet
    Dim anchor As Worksheet
    Dim CurrentDate As Date
    Dim LastDate As Date

    Set anchor = seedSheet
    CurrentDate = DateSerial(YearNo, MonthNo, 1)
    LastDate = DateSerial(YearNo, MonthNo + 1, 0)

    Do While CurrentDate <= LastDate
        If Weekday(CurrentDate, vbSunday) <> vbFriday Then
            ' Always copy the clean seed; place after last day sheet (never after HiddenLists)
            seedSheet.Copy After:=anchor
            Set NewSheet = anchor.Next

            On Error Resume Next
            NewSheet.Name = Format(CurrentDate, "dd-mmmm-yyyy")
            If Err.Number <> 0 Then
                Err.Clear
                NewSheet.Name = Format(CurrentDate, "dd-mmm-yyyy")
            End If
            On Error GoTo 0

            NewSheet.Range("A8").Value = EmployeeName
            NewSheet.Range("C8").Value = Format(CurrentDate, "dddd, dd mmmm yyyy")

            FitMergedCell NewSheet.Range("A8")
            FitMergedCell NewSheet.Range("C8")

            ' Sheet-copy sometimes drops Recovered Amount SUM â€” rewrite all summary formulas
            FixSummaryFormulas NewSheet

            Set anchor = NewSheet
        End If
        CurrentDate = CurrentDate + 1
    Loop

    ' Remove the seed "Recovery Template" sheet
    Application.DisplayAlerts = False
    seedSheet.Delete
    Application.DisplayAlerts = True

End Sub

' Rewrite summary formulas to the sheet's actual table name (tblRecovery2, tblRecovery3, ...)
Private Sub FixSummaryFormulas(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim t As String

    On Error Resume Next
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    t = lo.Name
    If Len(t) = 0 Then Exit Sub

    ws.Range("B46").Formula = "=SUM(" & t & "[Bill Amount])"
    ws.Range("B49").Formula = "=SUM(" & t & "[Recovered Amount])"
    ws.Range("B52").Formula = "=SUM(" & t & "[Remaining Amount])"
    ws.Range("B55").Formula = "=COUNTIF(" & t & "[Reference No.],""<>"")"
    ws.Range("B58").Formula = "=COUNTIF(" & t & "[Bill Status],""Returned"")"

    ws.Range("E46").Formula = "=COUNTIF(" & t & "[Bill Status],""Paid"")"
    ws.Range("E49").Formula = "=COUNTIF(" & t & "[Bill Status],""Partial"")"
    ws.Range("E52").Formula = "=COUNTIF(" & t & "[Bill Status],""PTO"")"
    ws.Range("E55").Formula = "=COUNTIF(" & t & "[Bill Status],""RETURNED"")"
    ws.Range("E58").Formula = "=COUNTIF(" & t & "[Bill Status],""HAND OVER TO SUPPLY MAN"")"

    ' Remaining: always numeric (avoids text/green-triangle). Blank when both
    ' Bill+Recovered empty is handled by conditional format ;;; on the column.
    If Not lo.ListColumns("Remaining Amount").DataBodyRange Is Nothing Then
        SetupRemainingAmountColumn lo.ListColumns("Remaining Amount").DataBodyRange
    End If
    ws.Range("B46,B49,B52").NumberFormat = "#,##0"

    ' FormulaHidden on merged cells must be cleared via MergeArea
    ' (B49 alone failed when Hidden=True; formula bar stayed blank when protected)
    ClearFormulaHidden ws.Range("B46")
    ClearFormulaHidden ws.Range("B49")
    ClearFormulaHidden ws.Range("B52")
    ClearFormulaHidden ws.Range("B55")
    ClearFormulaHidden ws.Range("B58")
    ClearFormulaHidden ws.Range("E46")
    ClearFormulaHidden ws.Range("E49")
    ClearFormulaHidden ws.Range("E52")
    ClearFormulaHidden ws.Range("E55")
    ClearFormulaHidden ws.Range("E58")

    On Error GoTo 0

End Sub

Private Sub ClearFormulaHidden(ByVal rng As Range)
    On Error Resume Next
    If rng.MergeCells Then
        rng.MergeArea.FormulaHidden = False
    Else
        rng.FormulaHidden = False
    End If
    On Error GoTo 0
End Sub

' Clear green error triangles (esp. unlocked formula cells after table unlock)
Private Sub IgnoreFormulaErrors(ByVal rng As Range)
    Dim cell As Range
    Dim i As Long
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    For Each cell In rng.Cells
        For i = 1 To 9
            cell.Errors.Item(i).Ignore = True
        Next i
    Next cell
    On Error GoTo 0
End Sub

' Numeric Remaining formula + hide 0 when Bill & Recovered both empty + lock + ignore errors
Private Sub SetupRemainingAmountColumn(ByVal remRng As Range)
    Dim fc As FormatCondition
    Dim firstRow As Long

    If remRng Is Nothing Then Exit Sub
    On Error Resume Next

    remRng.Formula = _
        "=IF([@[Bill Amount]]="""",0,[@[Bill Amount]])-IF([@[Recovered Amount]]="""",0,[@[Recovered Amount]])"
    remRng.NumberFormat = "#,##0"
    remRng.Locked = True

    Do While remRng.FormatConditions.Count > 0
        remRng.FormatConditions(1).Delete
    Loop

    firstRow = remRng.Row
    ' Hide value when both amount cells on this row are empty (shows blank, not 0)
    Set fc = remRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(C" & firstRow & "="""",D" & firstRow & "="""")")
    fc.NumberFormat = ";;;"

    IgnoreFormulaErrors remRng
    On Error GoTo 0
End Sub

' Current month -> today (if Friday -> next day). Past/future -> first sheet.
Public Sub ActivateStartupSheet(ByVal wb As Workbook, ByVal MonthNo As Long, ByVal YearNo As Long)

    Dim ws As Worksheet
    Dim targetName As String
    Dim d As Date

    On Error Resume Next

    If Month(Date) = MonthNo And Year(Date) = YearNo Then
        d = Date
        If Weekday(d, vbSunday) = vbFriday Then
            d = DateAdd("d", 1, d)
        End If

        Do While Month(d) = MonthNo And Year(d) = YearNo
            targetName = Format(d, "dd-mmmm-yyyy")
            Set ws = Nothing
            Set ws = wb.Worksheets(targetName)
            If Not ws Is Nothing Then
                If ws.Visible = xlSheetVisible Then
                    ws.Activate
                    Exit Sub
                End If
            End If
            d = DateAdd("d", 1, d)
        Loop
    End If

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ws.Name <> "HiddenLists" Then
                ws.Activate
                Exit Sub
            End If
        End If
    Next ws

End Sub

Private Sub InstallAutoOpenSheetCode(ByVal wb As Workbook, ByVal MonthNo As Long, ByVal YearNo As Long)

    Dim cm As Object
    Dim code As String

    On Error GoTo InstallFail

    Set cm = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    code = "Private Const SHEET_PW As String = ""2968""" & vbCrLf & vbCrLf

    ' Workbook_Open
    code = code & "Private Sub Workbook_Open()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Application.EnableEvents = True" & vbCrLf
    code = code & "    ApplyProtectionForEditing" & vbCrLf
    code = code & "    ActivateStartupDaySheet" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Protect sheets; Excel greys Table Insert/Delete while protected,
    ' so SheetSelectionChange unprotects only while cursor is in the table.
    code = code & "Private Sub ApplyProtectionForEditing()" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Me.Unprotect Password:=SHEET_PW" & vbCrLf
    code = code & "    For Each ws In Me.Worksheets" & vbCrLf
    code = code & "        If ws.Name <> ""HiddenLists"" Then ProtectOneSheet ws" & vbCrLf
    code = code & "    Next ws" & vbCrLf
    code = code & "    Me.Protect Password:=SHEET_PW, Structure:=True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub ProtectOneSheet(ByVal ws As Worksheet)" & vbCrLf
    code = code & "    Dim lo As ListObject" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    ws.Unprotect Password:=SHEET_PW" & vbCrLf
    code = code & "    ws.Cells.Locked = True" & vbCrLf
    code = code & "    ws.Range(""E8"").MergeArea.Locked = False" & vbCrLf
    code = code & "    ws.Range(""G8"").MergeArea.Locked = False" & vbCrLf
    code = code & "    For Each lo In ws.ListObjects" & vbCrLf
    code = code & "        If Not lo.HeaderRowRange Is Nothing Then lo.HeaderRowRange.Locked = True" & vbCrLf
    code = code & "        If Not lo.DataBodyRange Is Nothing Then" & vbCrLf
    code = code & "            lo.DataBodyRange.EntireRow.Locked = False" & vbCrLf
    code = code & "            lo.DataBodyRange.Rows(lo.DataBodyRange.Rows.Count).Offset(1, 0).EntireRow.Locked = False" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "        On Error Resume Next" & vbCrLf
    code = code & "        Dim remCol As ListColumn, remCell As Range, ej As Long" & vbCrLf
    code = code & "        Set remCol = lo.ListColumns(""Remaining Amount"")" & vbCrLf
    code = code & "        If Not remCol Is Nothing Then" & vbCrLf
    code = code & "            If Not remCol.DataBodyRange Is Nothing Then" & vbCrLf
    code = code & "                remCol.DataBodyRange.Locked = True" & vbCrLf
    code = code & "                For Each remCell In remCol.DataBodyRange.Cells" & vbCrLf
    code = code & "                    For ej = 1 To 9" & vbCrLf
    code = code & "                        remCell.Errors.Item(ej).Ignore = True" & vbCrLf
    code = code & "                    Next ej" & vbCrLf
    code = code & "                Next remCell" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "    ws.Protect Password:=SHEET_PW, UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Native Table row Insert/Delete only works while sheet is unprotected
    code = code & "Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    code = code & "    Dim lo As ListObject, zone As Range, inTable As Boolean" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    If Sh Is Nothing Or Target Is Nothing Then Exit Sub" & vbCrLf
    code = code & "    If Sh.Name = ""HiddenLists"" Then Exit Sub" & vbCrLf
    code = code & "    inTable = False" & vbCrLf
    code = code & "    For Each lo In Sh.ListObjects" & vbCrLf
    code = code & "        If Not lo.Range Is Nothing Then" & vbCrLf
    code = code & "            Set zone = lo.Range.Resize(lo.Range.Rows.Count + 1, lo.Range.Columns.Count)" & vbCrLf
    code = code & "            If Not Intersect(Target, zone) Is Nothing Then inTable = True: Exit For" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "    If inTable Then" & vbCrLf
    code = code & "        If Sh.ProtectContents Then Sh.Unprotect Password:=SHEET_PW" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        If Not Sh.ProtectContents Then ProtectOneSheet Sh" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Day sheet activation
    code = code & "Private Sub ActivateStartupDaySheet()" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    Dim targetName As String" & vbCrLf
    code = code & "    Dim wbMonth As Long" & vbCrLf
    code = code & "    Dim wbYear As Long" & vbCrLf
    code = code & "    Dim d As Date" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    wbMonth = " & CStr(MonthNo) & vbCrLf
    code = code & "    wbYear = " & CStr(YearNo) & vbCrLf
    code = code & "    If Month(Date) = wbMonth And Year(Date) = wbYear Then" & vbCrLf
    code = code & "        d = Date" & vbCrLf
    code = code & "        If Weekday(d, vbSunday) = vbFriday Then d = DateAdd(""d"", 1, d)" & vbCrLf
    code = code & "        Do While Month(d) = wbMonth And Year(d) = wbYear" & vbCrLf
    code = code & "            targetName = Format(d, ""dd-mmmm-yyyy"")" & vbCrLf
    code = code & "            Set ws = Nothing" & vbCrLf
    code = code & "            Set ws = Me.Worksheets(targetName)" & vbCrLf
    code = code & "            If Not ws Is Nothing Then" & vbCrLf
    code = code & "                If ws.Visible = xlSheetVisible Then" & vbCrLf
    code = code & "                    ws.Activate" & vbCrLf
    code = code & "                    If ws.ListObjects.Count > 0 Then" & vbCrLf
    code = code & "                        If Not ws.ListObjects(1).DataBodyRange Is Nothing Then ws.ListObjects(1).DataBodyRange.Cells(1, 1).Select" & vbCrLf
    code = code & "                    End If" & vbCrLf
    code = code & "                    Exit Sub" & vbCrLf
    code = code & "                End If" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "            d = DateAdd(""d"", 1, d)" & vbCrLf
    code = code & "        Loop" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    For Each ws In Me.Worksheets" & vbCrLf
    code = code & "        If ws.Visible = xlSheetVisible And ws.Name <> ""HiddenLists"" Then" & vbCrLf
    code = code & "            ws.Activate" & vbCrLf
    code = code & "            If ws.ListObjects.Count > 0 Then" & vbCrLf
    code = code & "                If Not ws.ListObjects(1).DataBodyRange Is Nothing Then ws.ListObjects(1).DataBodyRange.Cells(1, 1).Select" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "            Exit Sub" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next ws" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Restore Remaining Amount formula
    code = code & "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    code = code & "    Dim lo As ListObject, col As ListColumn, rng As Range, cell As Range" & vbCrLf
    code = code & "    Dim q As String, f As String, ei As Long" & vbCrLf
    code = code & "    On Error GoTo CleanExit" & vbCrLf
    code = code & "    If Sh Is Nothing Or Target Is Nothing Then GoTo CleanExit" & vbCrLf
    code = code & "    If Sh.Name = ""HiddenLists"" Then GoTo CleanExit" & vbCrLf
    code = code & "    q = Chr(34)" & vbCrLf
    code = code & "    f = ""=IF([@[Bill Amount]]="" & q & q & "",0,[@[Bill Amount]])-IF([@[Recovered Amount]]="" & q & q & "",0,[@[Recovered Amount]])""" & vbCrLf
    code = code & "    Application.EnableEvents = False" & vbCrLf
    code = code & "    For Each lo In Sh.ListObjects" & vbCrLf
    code = code & "        Set col = Nothing" & vbCrLf
    code = code & "        On Error Resume Next" & vbCrLf
    code = code & "        Set col = lo.ListColumns(""Remaining Amount"")" & vbCrLf
    code = code & "        On Error GoTo CleanExit" & vbCrLf
    code = code & "        If Not col Is Nothing Then" & vbCrLf
    code = code & "            If Not col.DataBodyRange Is Nothing Then" & vbCrLf
    code = code & "                Set rng = Intersect(Target, col.DataBodyRange)" & vbCrLf
    code = code & "                If Not rng Is Nothing Then" & vbCrLf
    code = code & "                    For Each cell In rng.Cells" & vbCrLf
    code = code & "                        cell.Formula = f" & vbCrLf
    code = code & "                        For ei = 1 To 9" & vbCrLf
    code = code & "                            cell.Errors.Item(ei).Ignore = True" & vbCrLf
    code = code & "                        Next ei" & vbCrLf
    code = code & "                    Next cell" & vbCrLf
    code = code & "                End If" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next lo" & vbCrLf
    code = code & "CleanExit:" & vbCrLf
    code = code & "    Application.EnableEvents = True" & vbCrLf
    code = code & "End Sub" & vbCrLf

    cm.AddFromString code
    Exit Sub

InstallFail:
End Sub
