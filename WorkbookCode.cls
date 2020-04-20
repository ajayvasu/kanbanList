VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    Dim changedCell As String
    changedCell = Target.Address(False, False)
    'Range("A5").Value = changedCell
    Application.EnableEvents = False
    If Sh.Index = 1 Then
        If Not Intersect(Target, Sheets(1).Range("B2:B600")) Is Nothing Then
            
            If Range(changedCell).Value <> "" Then
                Range(changedCell).Interior.ColorIndex = 6
            Else
                Range(changedCell).Interior.ColorIndex = 0
            End If
            
        End If
    End If
    Application.EnableEvents = True
End Sub



