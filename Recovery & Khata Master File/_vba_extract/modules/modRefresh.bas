Attribute VB_Name = "modRefresh"
Option Explicit

' Cached dropdown lists (loaded once for whole batch)
Private gAreas As Variant
Private gDayStatus As Variant
Private gRecoveryStatus As Variant
Private gAreasCount As Long
Private gDayCount As Long
Private gRecCount As Long
Private gLastRefreshError As String

Public Sub RefreshAllEmployeeWorkbooks()

    Dim RootFolder As String
    Dim fileList As Collection
    Dim i As Long
    Dim filePath As String
    Dim anyFail As Boolean
    Dim oldCalc As XlCalculation
    Dim oldSecurity As Long
    Dim oldStatusBar As Variant
    Dim oldWindowState As XlWindowState
    Dim okCount As Long
    Dim failCount As Long

    anyFail = False
    okCount = 0
    failCount = 0
    gLastRefreshError = ""

    oldCalc = Application.Calculation
    oldStatusBar = Application.DisplayStatusBar
    oldWindowState = Application.WindowState

    On Error Resume Next
    oldSecurity = Application.AutomationSecurity
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    Application.DisplayStatusBar = True
    Application.Cursor = xlWait
    ' Keep Excel interactive + visible so window does not drop behind other apps
    Application.Visible = True
    If Application.WindowState = xlMinimized Then Application.WindowState = xlNormal

    On Error Resume Next
    Application.PrintCommunication = False
    Application.AutomationSecurity = 3  ' msoAutomationSecurityForceDisable — skip Workbook_Open
    On Error GoTo ErrorHandler

    RootFolder = Trim(ThisWorkbook.Worksheets("Home").Range("txtRootFolder").Value)

    If RootFolder = "" Then
        BringManagerToFront oldWindowState
        MsgBox "Root Folder not found.", vbExclamation, "Recovery Manager"
        GoTo CleanExit
    End If

    If Dir(RootFolder, vbDirectory) = "" Then
        BringManagerToFront oldWindowState
        MsgBox "Root Folder path does not exist:" & vbCrLf & RootFolder, _
               vbExclamation, "Recovery Manager"
        GoTo CleanExit
    End If

    EnsureTrustedLocation RootFolder

    ' Close open registers once — avoids blink + re-open conflicts
    CloseOpenRecoveryRegisters
    BringManagerToFront oldWindowState

    ' Load Settings lists ONCE (not per workbook)
    If Not LoadDropdownCache() Then
        BringManagerToFront oldWindowState
        MsgBox "Could not read dropdown lists from Settings.", vbCritical, "Recovery Manager"
        GoTo CleanExit
    End If

    Set fileList = New Collection
    CollectRecoveryFiles RootFolder, fileList

    If fileList.Count = 0 Then
        BringManagerToFront oldWindowState
        MsgBox "No Recovery Register workbooks found.", vbExclamation, "Recovery Manager"
        GoTo CleanExit
    End If

    Application.StatusBar = "Refreshing dropdowns: 0 / " & fileList.Count

    For i = 1 To fileList.Count
        filePath = CStr(fileList(i))

        If FastRefreshOneWorkbook(filePath) Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
            anyFail = True
        End If

        ' Progress only — keep Manager in front (avoid DoEvents stealing focus every time)
        If (i Mod 25) = 0 Or i = fileList.Count Then
            Application.StatusBar = "Refreshing dropdowns: " & i & " / " & fileList.Count
            On Error Resume Next
            ThisWorkbook.Activate
            On Error GoTo ErrorHandler
        End If
    Next i

    BringManagerToFront oldWindowState

    If anyFail Then
        If Len(gLastRefreshError) > 0 Then
            MsgBox "Failed." & vbCrLf & vbCrLf & gLastRefreshError, vbCritical, "Recovery Manager"
        Else
            MsgBox "Failed.", vbCritical, "Recovery Manager"
        End If
    Else
        MsgBox "Success.", vbInformation, "Recovery Manager"
    End If

    BringManagerToFront oldWindowState

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.Cursor = xlDefault
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    Application.AutomationSecurity = oldSecurity
    BringManagerToFront oldWindowState
    On Error GoTo 0

    Erase gAreas
    Erase gDayStatus
    Erase gRecoveryStatus
    Exit Sub

ErrorHandler:
    On Error Resume Next
    BringManagerToFront oldWindowState
    On Error GoTo 0
    MsgBox "Failed.", vbCritical, "Recovery Manager"
    Resume CleanExit

End Sub

Private Sub BringManagerToFront(Optional ByVal preferredState As XlWindowState = xlNormal)

    On Error Resume Next

    Application.Visible = True

    If Application.WindowState = xlMinimized Then
        If preferredState = xlMinimized Then
            Application.WindowState = xlNormal
        Else
            Application.WindowState = preferredState
        End If
    End If

    ThisWorkbook.Activate
    If ThisWorkbook.Windows.Count > 0 Then
        ThisWorkbook.Windows(1).Visible = True
        ThisWorkbook.Windows(1).Activate
    End If

    AppActivate Application.Caption

    On Error GoTo 0

End Sub

Private Function LoadDropdownCache() As Boolean

    Dim wsSettings As Worksheet
    Dim tbl As ListObject
    Dim src As Range

    On Error GoTo FailLoad

    Set wsSettings = ThisWorkbook.Worksheets("Settings")

    gAreasCount = 0
    gDayCount = 0
    gRecCount = 0
    gAreas = Empty
    gDayStatus = Empty
    gRecoveryStatus = Empty

    Set tbl = wsSettings.ListObjects("tblAreas")
    Set src = Nothing
    On Error Resume Next
    Set src = tbl.ListColumns(1).DataBodyRange
    On Error GoTo FailLoad
    If Not src Is Nothing Then
        gAreas = src.Value
        gAreasCount = src.Rows.Count
    End If

    Set tbl = wsSettings.ListObjects("tblDayStatus")
    Set src = Nothing
    On Error Resume Next
    Set src = tbl.ListColumns(1).DataBodyRange
    On Error GoTo FailLoad
    If Not src Is Nothing Then
        gDayStatus = src.Value
        gDayCount = src.Rows.Count
    End If

    Set tbl = wsSettings.ListObjects("tblRecoveryStatus")
    Set src = Nothing
    On Error Resume Next
    Set src = tbl.ListColumns(1).DataBodyRange
    On Error GoTo FailLoad
    If Not src Is Nothing Then
        gRecoveryStatus = src.Value
        gRecCount = src.Rows.Count
    End If

    LoadDropdownCache = True
    Exit Function

FailLoad:
    LoadDropdownCache = False

End Function

' Ultra-light refresh: update HiddenLists only. No SecureWorkbook, no sheet loop, no Activate.
Private Function FastRefreshOneWorkbook(ByVal WorkbookPath As String) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim needNames As Boolean

    FastRefreshOneWorkbook = False

    On Error GoTo FailFast

    Set wb = Workbooks.Open(FileName:=WorkbookPath, _
                            UpdateLinks:=0, _
                            ReadOnly:=False, _
                            IgnoreReadOnlyRecommended:=True, _
                            Notify:=False, _
                            AddToMru:=False)

    ' Hide opened file window so Manager stays in front (must show again before Save)
    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo FailFast

    ' Structure-protect blocks sheet visibility changes — unlock once up front
    On Error Resume Next
    wb.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo FailFast

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("HiddenLists")
    On Error GoTo FailFast

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "HiddenLists"
        needNames = True
    Else
        needNames = Not NamedRangesExistPublic(wb)
    End If

    ws.Range("A2:C1000").ClearContents
    ws.Range("A1").Value = "Areas"
    ws.Range("B1").Value = "Day Status"
    ws.Range("C1").Value = "Recovery Status"

    If gAreasCount > 0 Then
        ws.Range("A2").Resize(gAreasCount, 1).Value = gAreas
    End If
    If gDayCount > 0 Then
        ws.Range("B2").Resize(gDayCount, 1).Value = gDayStatus
    End If
    If gRecCount > 0 Then
        ws.Range("C2").Resize(gRecCount, 1).Value = gRecoveryStatus
    End If

    ws.Visible = xlSheetVeryHidden

    If needNames Then
        CreateNamedRanges wb
    End If

    On Error Resume Next
    wb.Protect Password:=SHEET_PASSWORD, Structure:=True
    Err.Clear
    ' NEVER save while window is hidden — that makes double-click open with no window
    wb.Windows(1).Visible = True
    On Error GoTo FailFast

    wb.Save
    wb.Close SaveChanges:=False
    Set wb = Nothing

    FastRefreshOneWorkbook = True
    Exit Function

FailFast:
    gLastRefreshError = Err.Description & " | " & WorkbookPath

    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Windows(1).Visible = True
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    On Error GoTo 0
    FastRefreshOneWorkbook = False

End Function

Private Sub CollectRecoveryFiles(ByVal FolderPath As String, ByRef fileList As Collection)

    Dim FSO As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim ext As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(FolderPath) Then Exit Sub

    Set Folder = FSO.GetFolder(FolderPath)

    For Each File In Folder.Files
        If Left(File.Name, 2) <> "~$" Then
            ext = LCase(FSO.GetExtensionName(File.Name))
            If ext = "xlsx" Or ext = "xlsm" Then
                If InStr(1, File.Name, "Recovery Register", vbTextCompare) > 0 Then
                    fileList.Add File.Path
                End If
            End If
        End If
    Next File

    For Each SubFolder In Folder.SubFolders
        If LCase(SubFolder.Name) <> "recovery_backup" Then
            CollectRecoveryFiles SubFolder.Path, fileList
        End If
    Next SubFolder

End Sub
