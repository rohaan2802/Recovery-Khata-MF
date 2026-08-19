VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ' Excel does not persist Ignore-Error for Remaining Amount formulas;
    ' clear green triangles every open on the template.
    ClearRemainingGreenTriangles Me.Worksheets("Recovery Template")
    ClearKhataGreenTriangles Me.Worksheets("Khata Template")
    On Error GoTo 0
End Sub

Private Sub ClearRemainingGreenTriangles(ByVal ws As Worksheet)
    Dim lo As ListObject
    Dim remRng As Range
    Dim cell As Range
    Dim i As Long

    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub

    Set lo = ws.ListObjects(1)
    Set remRng = lo.ListColumns("Remaining Amount").DataBodyRange
    If remRng Is Nothing Then Exit Sub

    remRng.Locked = True
    For Each cell In remRng.Cells
        For i = 1 To 9
            cell.Errors.Item(i).Ignore = True
        Next i
    Next cell
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If Sh.Name = "Khata Template" Then
        ClearKhataGreenTriangles Sh
        EnsureKhataTemplateBorders Sh
    End If
    On Error GoTo 0
End Sub

Private Sub ClearKhataGreenTriangles(ByVal ws As Worksheet)
    Dim lo As ListObject
    Dim balRng As Range
    Dim cell As Range
    Dim i As Long
    Dim f As String

    If ws Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub

    Set lo = ws.ListObjects(1)
    Set balRng = lo.ListColumns("Balance").DataBodyRange
    If balRng Is Nothing Then Exit Sub

    f = "=IF(AND([@Debit]="""",[@Credit]=""""),"""",SUM($E$8,INDEX([Debit],1):[@Debit])-SUM(INDEX([Credit],1):[@Credit]))"
    balRng.Formula = f
    balRng.Locked = True

    For Each cell In balRng.Cells
        For i = 1 To 9
            cell.Errors.Item(i).Ignore = True
        Next i
    Next cell
End Sub

Private Sub EnsureKhataTemplateBorders(ByVal ws As Worksheet)
    Dim lo As ListObject
    On Error Resume Next
    For Each lo In ws.ListObjects
        With lo.Range.Borders
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
    Next lo
End Sub
