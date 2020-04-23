Public Sub CommandButton1_Click()
    Worksheets(1).Range("A1").AutoFilter _
    Field:=4, _
    Criteria1:="R4", _
    VisibleDropDown:=False

End Sub

Public Sub CommandButton2_Click()
    Worksheets(1).Range("A1").AutoFilter _
    Field:=4, _
    Criteria1:="R5", _
    VisibleDropDown:=False
End Sub

Public Sub CommandButton3_Click()
    Worksheets(1).Range("A1").AutoFilter _
    Field:=4, _
    Criteria1:="R6", _
    VisibleDropDown:=False
End Sub

Public Sub CommandButton4_Click()
    Worksheets(1).Range("A1").AutoFilter _
    Field:=4, _
    Criteria1:="maint", _
    VisibleDropDown:=False
End Sub

Public Sub CommandButton5_Click()
    Worksheets(1).Range("A1").AutoFilter
    Worksheets(1).Range("A1").AutoFilter
End Sub



Private Sub CommandButton6_Click()
    Worksheets(1).Range("A1").AutoFilter _
    Field:=6, _
    Criteria1:="RemedyForce", _
    VisibleDropDown:=False
End Sub

Private Sub CommandButton7_Click()
    Worksheets(1).Range("A1").AutoFilter _
    Field:=6, _
    Criteria1:="ChaRM", _
    VisibleDropDown:=False
End Sub

Private Sub CommandButton8_Click()
    Worksheets(1).Range("A1").AutoFilter
    Worksheets(1).Range("A1").AutoFilter
End Sub
