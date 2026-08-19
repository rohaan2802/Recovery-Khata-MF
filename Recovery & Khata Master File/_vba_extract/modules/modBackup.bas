Attribute VB_Name = "modBackup"
Option Explicit

Private FilesProcessed As Long
Private FilesBackedUp As Long
Private FilesSkipped As Long
Private anyFail As Boolean

Private RootFolder As String
Private BackupRoot As String
Private gFSO As Object

Public Sub WorkDoneCreateBackup()

    Dim fileList As Collection
    Dim i As Long
    Dim filePath As String
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldUpdating As Boolean
    Dim oldSecurity As Long
    Dim oldWindowState As XlWindowState

    RootFolder = Trim(ThisWorkbook.Worksheets("Home").Range("txtRootFolder").Value)

    If RootFolder = "" Then
        MsgBox "Please select Root Folder first.", vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    If Dir(RootFolder, vbDirectory) = "" Then
        MsgBox "Root Folder path does not exist:" & vbCrLf & RootFolder, _
               vbExclamation, "Recovery Manager"
        Exit Sub
    End If

    BackupRoot = Left(RootFolder, 2) & "\Recovery_Backup"

    FilesProcessed = 0
    FilesBackedUp = 0
    FilesSkipped = 0
    anyFail = False

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldUpdating = Application.ScreenUpdating
    oldWindowState = Application.WindowState

    On Error Resume Next
    oldSecurity = Application.AutomationSecurity
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    Application.Cursor = xlWait
    Application.Visible = True
    If Application.WindowState = xlMinimized Then Application.WindowState = xlNormal

    On Error Resume Next
    Application.PrintCommunication = False
    Application.AutomationSecurity = 3  ' skip Workbook_Open macros
    On Error GoTo ErrorHandler

    EnsureTrustedLocation RootFolder
    CreateBackupRoot

    CloseOpenRecoveryRegisters

    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set fileList = New Collection
    CollectWorkbookFiles RootFolder, fileList

    Application.StatusBar = "Work Done & Backup: 0 / " & fileList.Count

    For i = 1 To fileList.Count
        filePath = CStr(fileList(i))
        ProcessWorkbook filePath

        If (i Mod 25) = 0 Or i = fileList.Count Then
            Application.StatusBar = "Work Done & Backup: " & i & " / " & fileList.Count
        End If
    Next i

    BringBackupManagerToFront oldWindowState

    If anyFail Then
        MsgBox "Failed Back Up", vbCritical, "Recovery Manager"
    Else
        MsgBox "Work Done & Sucessfull Back Up", vbInformation, "Recovery Manager"
    End If

    BringBackupManagerToFront oldWindowState

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldUpdating
    Application.PrintCommunication = True
    Application.AutomationSecurity = oldSecurity
    Set gFSO = Nothing
    BringBackupManagerToFront oldWindowState
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    anyFail = True
    On Error Resume Next
    BringBackupManagerToFront oldWindowState
    On Error GoTo 0
    MsgBox "Failed Back Up", vbCritical, "Recovery Manager"
    Resume CleanExit

End Sub

Private Sub BringBackupManagerToFront(Optional ByVal preferredState As XlWindowState = xlNormal)

    On Error Resume Next
    Application.Visible = True
    If Application.WindowState = xlMinimized Then
        Application.WindowState = IIf(preferredState = xlMinimized, xlNormal, preferredState)
    End If
    ThisWorkbook.Activate
    If ThisWorkbook.Windows.Count > 0 Then
        ThisWorkbook.Windows(1).Visible = True
        ThisWorkbook.Windows(1).Activate
    End If
    AppActivate Application.Caption
    On Error GoTo 0

End Sub

Private Sub CreateBackupRoot()

    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.FolderExists(BackupRoot) Then
        FSO.CreateFolder BackupRoot
        On Error Resume Next
        SetAttr BackupRoot, vbHidden
        On Error GoTo 0
    End If

End Sub

Private Sub CollectWorkbookFiles(ByVal FolderPath As String, ByRef fileList As Collection)

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
                If InStr(1, File.Name, "Recovery Register", vbTextCompare) > 0 _
                   Or InStr(1, File.Name, "Khata Register", vbTextCompare) > 0 Then
                    fileList.Add File.Path
                End If
            End If
        End If
    Next File

    For Each SubFolder In Folder.SubFolders
        If LCase(SubFolder.Path) <> LCase(BackupRoot) Then
            If LCase(SubFolder.Name) <> "recovery_backup" Then
                CollectWorkbookFiles SubFolder.Path, fileList
            End If
        End If
    Next SubFolder

End Sub

Private Function GetRelativePath(ByVal FullPath As String) As String

    If Len(FullPath) <= Len(RootFolder) Then
        GetRelativePath = ""
    Else
        GetRelativePath = Mid(FullPath, Len(RootFolder) + 2)
    End If

End Function

Private Sub ProcessWorkbook(ByVal FilePath As String)

    Dim wb As Workbook

    FilesProcessed = FilesProcessed + 1

    On Error GoTo ProcessError

    ' Always open closed (registers already closed) — hide window to stop master blink
    Set wb = Workbooks.Open( _
                FileName:=FilePath, _
                UpdateLinks:=0, _
                ReadOnly:=False, _
                IgnoreReadOnlyRecommended:=True, _
                Notify:=False, _
                AddToMru:=False)

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo ProcessError

    FastFitWorkbook wb

    On Error Resume Next
    wb.Windows(1).Visible = True
    On Error GoTo ProcessError

    wb.Save
    wb.Close SaveChanges:=False
    Set wb = Nothing

    CopyWorkbook FilePath
    Exit Sub

ProcessError:
    FilesSkipped = FilesSkipped + 1
    anyFail = True

    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Windows(1).Visible = True
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    On Error GoTo 0

End Sub

' Light fit only — no full SecureWorkbook rebuild (was main blink + lag)
Private Sub FastFitWorkbook(ByVal wb As Workbook)

    Dim ws As Worksheet

    On Error Resume Next
    wb.Unprotect Password:=SHEET_PASSWORD

    For Each ws In wb.Worksheets
        If ws.Name <> "HiddenLists" Then
            If ws.Visible = xlSheetVisible Then
                ws.Unprotect Password:=SHEET_PASSWORD
                FitMergedCell ws.Range("A8")
                FitMergedCell ws.Range("C8")
                FitMergedCell ws.Range("E8")
                FitMergedCell ws.Range("G8")
                FastFitTableColumns ws
                ' Re-apply same protection: sheet locked, table insert/delete allowed
                ProtectSheetAllowTableEdit ws
            End If
        End If
    Next ws

    wb.Protect Password:=SHEET_PASSWORD, Structure:=True
    On Error GoTo 0

End Sub

Private Sub FastFitTableColumns(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim col As ListColumn
    Dim minW As Double
    Dim needed As Double

    On Error Resume Next
    For Each lo In ws.ListObjects
        For Each col In lo.ListColumns
            If Not col.Range Is Nothing Then
                minW = col.Range.Columns(1).ColumnWidth
                col.Range.Columns(1).AutoFit
                needed = col.Range.Columns(1).ColumnWidth
                If needed < minW Then col.Range.Columns(1).ColumnWidth = minW
            End If
        Next col
    Next lo
    On Error GoTo 0

End Sub

Private Sub CopyWorkbook(ByVal SourceFile As String)

    Dim RelativePath As String
    Dim DestinationFile As String
    Dim DestinationFolder As String

    If gFSO Is Nothing Then Set gFSO = CreateObject("Scripting.FileSystemObject")

    RelativePath = GetRelativePath(SourceFile)
    If RelativePath = "" Then
        FilesSkipped = FilesSkipped + 1
        Exit Sub
    End If

    DestinationFile = BackupRoot & "\" & RelativePath
    DestinationFolder = Left(DestinationFile, InStrRev(DestinationFile, "\") - 1)

    If Not gFSO.FolderExists(DestinationFolder) Then
        CreateFolderTree DestinationFolder
    End If

    If gFSO.FileExists(DestinationFile) Then
        gFSO.DeleteFile DestinationFile, True
    End If

    gFSO.CopyFile SourceFile, DestinationFile, True

    If gFSO.FileExists(DestinationFile) Then
        FilesBackedUp = FilesBackedUp + 1
    Else
        FilesSkipped = FilesSkipped + 1
        anyFail = True
    End If

End Sub

Private Sub CreateFolderTree(ByVal FolderPath As String)

    Dim Parts() As String
    Dim CurrentPath As String
    Dim i As Long

    If gFSO Is Nothing Then Set gFSO = CreateObject("Scripting.FileSystemObject")

    Parts = Split(FolderPath, "\")
    CurrentPath = Parts(0)

    For i = 1 To UBound(Parts)
        If Len(Parts(i)) > 0 Then
            CurrentPath = CurrentPath & "\" & Parts(i)
            If Not gFSO.FolderExists(CurrentPath) Then
                gFSO.CreateFolder CurrentPath
            End If
        End If
    Next i

End Sub
