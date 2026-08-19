Attribute VB_Name = "modDropdowns"
Option Explicit

' Returns True on success.
' showErrors:=False skips MsgBox (batch refresh).
' manageAppState:=False leaves ScreenUpdating/Events to the caller (batch refresh).
' applyValidations:=False = fast refresh (only update HiddenLists data; do not close already-open books).
Public Function UpdateWorkbookDropdowns(ByVal WorkbookPath As String, _
                                        Optional ByVal showErrors As Boolean = True, _
                                        Optional ByVal manageAppState As Boolean = True, _
                                        Optional ByVal applyValidations As Boolean = True) As Boolean

    Dim wb As Workbook
    Dim errMsg As String
    Dim succeeded As Boolean
    Dim alreadyOpen As Boolean

    succeeded = False
    alreadyOpen = False
    UpdateWorkbookDropdowns = False

    If manageAppState Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
    End If

    On Error GoTo ErrorHandler

    Set wb = FindWorkbookByPath(WorkbookPath)
    alreadyOpen = Not wb Is Nothing

    If Not alreadyOpen Then
        Set wb = Workbooks.Open(FileName:=WorkbookPath, _
                                UpdateLinks:=False, _
                                ReadOnly:=False, _
                                IgnoreReadOnlyRecommended:=True, _
                                Notify:=False, _
                                AddToMru:=False)
        On Error Resume Next
        ThisWorkbook.Activate
        On Error GoTo ErrorHandler
    End If

    EnsureWorkbookWindowsVisible wb

    On Error Resume Next
    wb.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo ErrorHandler

    If applyValidations Then
        UnprotectAllSheets wb
        CreateHiddenListsSheet wb
        CreateNamedRanges wb
        ApplyDropdownsToWorkbook wb
        SecureWorkbook wb
    Else
        ' Fast path for single-file calls: lists only, no full SecureWorkbook
        CreateHiddenListsSheet wb
        If Not NamedRangesExist(wb) Then CreateNamedRanges wb
        On Error Resume Next
        wb.Protect Password:=SHEET_PASSWORD, Structure:=True
        On Error GoTo ErrorHandler
    End If

    EnsureWorkbookWindowsVisible wb
    wb.Save

    ' Never close a workbook the user already had open
    If Not alreadyOpen Then
        wb.Close SaveChanges:=False
    End If
    Set wb = Nothing

    succeeded = True

CleanExit:
    UpdateWorkbookDropdowns = succeeded

    On Error Resume Next
    ThisWorkbook.Activate
    On Error GoTo 0

    If manageAppState Then
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
    Exit Function

ErrorHandler:
    errMsg = Err.Description
    succeeded = False

    On Error Resume Next
    If Not wb Is Nothing Then
        EnsureWorkbookWindowsVisible wb
        If Not alreadyOpen Then wb.Close SaveChanges:=False
    End If
    Set wb = Nothing
    On Error GoTo 0

    If showErrors Then
        MsgBox "Dropdown update failed:" & vbCrLf & errMsg, _
               vbCritical, "Recovery Manager"
    End If

    Resume CleanExit

End Function

Public Function FindWorkbookByPath(ByVal FullPath As String) As Workbook

    Dim wb As Workbook
    Dim target As String

    target = LCase$(FullPath)

    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = target Then
            Set FindWorkbookByPath = wb
            Exit Function
        End If
    Next wb

End Function

Private Function NamedRangesExist(ByVal wb As Workbook) As Boolean
    NamedRangesExist = NamedRangesExistPublic(wb)
End Function

Public Function NamedRangesExistPublic(ByVal wb As Workbook) As Boolean

    Dim n As Name

    On Error Resume Next
    Err.Clear
    Set n = wb.Names("AreasList")
    If Err.Number <> 0 Then GoTo NoRanges

    Err.Clear
    Set n = wb.Names("DayStatusList")
    If Err.Number <> 0 Then GoTo NoRanges

    Err.Clear
    Set n = wb.Names("RecoveryStatusList")
    If Err.Number <> 0 Then GoTo NoRanges

    NamedRangesExistPublic = True
    On Error GoTo 0
    Exit Function

NoRanges:
    NamedRangesExistPublic = False
    On Error GoTo 0

End Function

Private Sub UnprotectAllSheets(ByVal wb As Workbook)

    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=SHEET_PASSWORD
        On Error GoTo 0
    Next ws

End Sub

Public Sub CreateHiddenListsSheet(ByVal wb As Workbook)

    Dim ws As Worksheet
    Dim wsSettings As Worksheet
    Dim tblAreas As ListObject
    Dim tblDayStatus As ListObject
    Dim tblRecoveryStatus As ListObject
    Dim src As Range

    On Error Resume Next
    Set ws = wb.Worksheets("HiddenLists")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "HiddenLists"
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "Areas"
    ws.Range("B1").Value = "Day Status"
    ws.Range("C1").Value = "Recovery Status"

    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    Set tblAreas = wsSettings.ListObjects("tblAreas")
    Set tblDayStatus = wsSettings.ListObjects("tblDayStatus")
    Set tblRecoveryStatus = wsSettings.ListObjects("tblRecoveryStatus")

    Set src = Nothing
    On Error Resume Next
    Set src = tblAreas.ListColumns(1).DataBodyRange
    On Error GoTo 0
    If Not src Is Nothing Then
        ws.Range("A2").Resize(src.Rows.Count, 1).Value = src.Value
    End If

    Set src = Nothing
    On Error Resume Next
    Set src = tblDayStatus.ListColumns(1).DataBodyRange
    On Error GoTo 0
    If Not src Is Nothing Then
        ws.Range("B2").Resize(src.Rows.Count, 1).Value = src.Value
    End If

    Set src = Nothing
    On Error Resume Next
    Set src = tblRecoveryStatus.ListColumns(1).DataBodyRange
    On Error GoTo 0
    If Not src Is Nothing Then
        ws.Range("C2").Resize(src.Rows.Count, 1).Value = src.Value
    End If

    ws.Visible = xlSheetVeryHidden

End Sub

Public Sub CreateNamedRanges(ByVal wb As Workbook)

    On Error Resume Next
    wb.Names("AreasList").Delete
    wb.Names("DayStatusList").Delete
    wb.Names("RecoveryStatusList").Delete
    On Error GoTo 0

    wb.Names.Add Name:="AreasList", _
        RefersTo:="=HiddenLists!$A$2:INDEX(HiddenLists!$A:$A,COUNTA(HiddenLists!$A:$A))"

    wb.Names.Add Name:="DayStatusList", _
        RefersTo:="=HiddenLists!$B$2:INDEX(HiddenLists!$B:$B,COUNTA(HiddenLists!$B:$B))"

    wb.Names.Add Name:="RecoveryStatusList", _
        RefersTo:="=HiddenLists!$C$2:INDEX(HiddenLists!$C:$C,COUNTA(HiddenLists!$C:$C))"

End Sub

Public Sub ApplyDropdownsToWorkbook(ByVal wb As Workbook)

    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If ws.Name <> "HiddenLists" Then
            ApplyDropdownsToSheet ws
        End If
    Next ws

End Sub

' Apply Area / Day Status / Bill Status dropdowns to one sheet
Public Sub ApplyDropdownsToSheet(ByVal ws As Worksheet)

    Dim rngArea As Range
    Dim rngDay As Range
    Dim lo As ListObject
    Dim statusCol As ListColumn
    Dim rngStatus As Range

    On Error Resume Next
    Set rngArea = ws.Range("E8").MergeArea
    Set rngDay = ws.Range("G8").MergeArea
    On Error GoTo 0

    ' IMPORTANT: Delete and Add must NOT share the same With block
    ' (Excel invalidates Validation after Delete)
    If Not rngArea Is Nothing Then
        On Error Resume Next
        rngArea.Validation.Delete
        On Error GoTo 0
        rngArea.Validation.Add Type:=xlValidateList, _
                               AlertStyle:=xlValidAlertStop, _
                               Operator:=xlBetween, _
                               Formula1:="=AreasList"
        With rngArea.Validation
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
        End With
    End If

    If Not rngDay Is Nothing Then
        On Error Resume Next
        rngDay.Validation.Delete
        On Error GoTo 0
        rngDay.Validation.Add Type:=xlValidateList, _
                              AlertStyle:=xlValidAlertStop, _
                              Operator:=xlBetween, _
                              Formula1:="=DayStatusList"
        With rngDay.Validation
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
        End With
    End If

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        Set statusCol = Nothing
        On Error Resume Next
        Set statusCol = lo.ListColumns("Bill Status")
        On Error GoTo 0

        If Not statusCol Is Nothing Then
            If Not statusCol.DataBodyRange Is Nothing Then
                Set rngStatus = statusCol.DataBodyRange
                On Error Resume Next
                rngStatus.Validation.Delete
                On Error GoTo 0
                rngStatus.Validation.Add Type:=xlValidateList, _
                                         AlertStyle:=xlValidAlertStop, _
                                         Operator:=xlBetween, _
                                         Formula1:="=RecoveryStatusList"
                With rngStatus.Validation
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowError = True
                End With
            End If
        End If
    End If

End Sub
