Attribute VB_Name = "modFolders"
Option Explicit

Public Sub CreateNewYear()

    Dim selectedYear As String
    Dim RootFolder As String
    Dim yearFolder As String

    selectedYear = Trim(Range("selYear").Value)
    RootFolder = Trim(Range("txtRootFolder").Value)

    If RootFolder = "" Then
        MsgBox "Please select a Root Folder.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If selectedYear = "" Then
        MsgBox "Please select a Year.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If Dir(RootFolder, vbDirectory) = "" Then
        MkDir RootFolder
    End If

    yearFolder = RootFolder & "\" & selectedYear

    If Dir(yearFolder, vbDirectory) <> "" Then
        MsgBox "Year folder already exists.", vbInformation, "Recovery Manager"
        Exit Sub
    End If

    MkDir yearFolder

    MsgBox "Year folder created successfully!" & vbCrLf & yearFolder, vbInformation, "Recovery Manager"

End Sub

Public Sub CreateNewMonth()

    Dim RootFolder As String
    Dim selectedYear As String
    Dim selectedMonth As String

    Dim monthFolderName As String
    Dim yearFolder As String
    Dim monthFolder As String

    RootFolder = Trim(Range("txtRootFolder").Value)
    selectedYear = Trim(Range("selYear").Value)
    selectedMonth = Trim(Range("selMonth").Value)

    If RootFolder = "" Then
        MsgBox "Please select a Root Folder.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If selectedYear = "" Then
        MsgBox "Please select a Year.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If selectedMonth = "" Then
        MsgBox "Please select a Month.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    yearFolder = RootFolder & "\" & selectedYear

    If Dir(yearFolder, vbDirectory) = "" Then
        MsgBox "Please create the Year folder first.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    monthFolderName = GetMonthFolderName(selectedMonth)

    If monthFolderName = "" Then
        MsgBox "Selected month was not found in Settings.", vbCritical, "Recovery Manager"
        Exit Sub
    End If

    monthFolder = yearFolder & "\" & monthFolderName

    If Dir(monthFolder, vbDirectory) <> "" Then
        MsgBox "Month folder already exists.", vbInformation, "Recovery Manager"
        Exit Sub
    End If

    MkDir monthFolder

    MsgBox "Month folder created successfully!" & vbCrLf & monthFolder, vbInformation, "Recovery Manager"

End Sub

' Builds folder name from Settings!tblMonths: "1. January"
Public Function GetMonthFolderName(ByVal monthName As String) As String

    Dim tbl As ListObject
    Dim rw As ListRow
    Dim monthNo As String
    Dim nameVal As String

    GetMonthFolderName = ""

    Set tbl = ThisWorkbook.Worksheets("Settings").ListObjects("tblMonths")

    For Each rw In tbl.ListRows
        monthNo = Trim(CStr(rw.Range.Cells(1, 1).Value))
        nameVal = Trim(CStr(rw.Range.Cells(1, 2).Value))

        If StrComp(nameVal, Trim(monthName), vbTextCompare) = 0 Then
            GetMonthFolderName = monthNo & ". " & nameVal
            Exit Function
        End If
    Next rw

End Function

Public Sub CreateEmployeeFolder()

    Dim RootFolder As String
    Dim selectedYear As String
    Dim selectedMonth As String
    Dim selectedEmployee As String

    Dim monthFolderName As String
    Dim employeeFolder As String
    Dim monthFolder As String

    RootFolder = Trim(Range("txtRootFolder").Value)
    selectedYear = Trim(Range("selYear").Value)
    selectedMonth = Trim(Range("selMonth").Value)
    selectedEmployee = Trim(Range("selEmployee").Value)

    If RootFolder = "" Then
        MsgBox "Please select a Root Folder.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If selectedYear = "" Then
        MsgBox "Please select a Year.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If selectedMonth = "" Then
        MsgBox "Please select a Month.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If selectedEmployee = "" Then
        MsgBox "Please select an Employee.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    monthFolderName = GetMonthFolderName(selectedMonth)

    If monthFolderName = "" Then
        MsgBox "Selected month was not found in Settings.", vbCritical, "Recovery Manager"
        Exit Sub
    End If

    monthFolder = RootFolder & "\" & selectedYear & "\" & monthFolderName

    If Dir(monthFolder, vbDirectory) = "" Then
        MsgBox "Please create the Month folder first.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    employeeFolder = monthFolder & "\" & selectedEmployee

    If Dir(employeeFolder, vbDirectory) <> "" Then
        MsgBox "Employee folder already exists.", vbInformation, "Recovery Manager"
        Exit Sub
    End If

    MkDir employeeFolder

    MsgBox "Employee folder created successfully!" & vbCrLf & employeeFolder, vbInformation, "Recovery Manager"

End Sub
