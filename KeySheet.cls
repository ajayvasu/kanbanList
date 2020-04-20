VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'**********************************************************
' Purpose: i.Consolidate  ii.Merge
'   i.Consolidate columns of RemedyForce & ChaRM exports
'       Input: 1.xlsx & 2.xlsx (RF & ChaRM workbooks in current directory)
'       Output: KanbanList.xlsm (This workbook)
'
'   ii.Merge yesterday's & today's Kanban List
'       Input: Kanbanlist.xlsm (Sheet 1 & 2 of this workbook)
'       Output: KanbanList.xlsm (with 1 merged sheet)
' Author: Ajay Vasu
' Company: Danfoss Engineering
' Date Created: 4/07/2020
'**********************************************************


Private Sub CommandButton1_Click()
     
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set kanbanList = ActiveWorkbook

    'The abstract ballet ensues..
    refreshAll
    addSheet kanbanList
    getTodayRows kanbanList
    formatColumns kanbanList
    Merge kanbanList
    deleteYesterday kanbanList
        
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Sub refreshAll()

    'Refresh RemedyForce
    ActiveWorkbook.Connections("Query - Table 2").Refresh
    
    'Refresh ChaRM
    
    
End Sub
Private Sub addSheet(kList)
    
    kList.Sheets.Add Before:=kList.Sheets(1)
    kList.Sheets(1).Name = Format(Now(), "DD-MM-YYYY hh-nn-ss")
    kList.Sheets(2).Range("A1:Q1").Copy Destination:=kList.Sheets(1).Range("A1")
    kList.Sheets(1).Range("A1").AutoFilter
    
End Sub

Private Sub getTodayRows(kList)
    
    Dim i As Integer
    
    'Combine all rows from today & yesterday.
    kList.Sheets(3).Range("I2:I300").Copy Destination:=kList.Sheets(1).Range("G2")
    kList.Sheets(3).Range("F2:F300").Copy Destination:=kList.Sheets(1).Range("P2")
    kList.Sheets(3).Range("A2:B300").Copy Destination:=kList.Sheets(1).Range("H2")
    kList.Sheets(3).Range("G2:G200").Copy Destination:=kList.Sheets(1).Range("J2")
    kList.Sheets(3).Range("C2:C300").Copy Destination:=kList.Sheets(1).Range("K2")
    kList.Sheets(3).Range("E2:E300").Copy Destination:=kList.Sheets(1).Range("O2")
    kList.Sheets(3).Range("H2:H300").Copy Destination:=kList.Sheets(1).Range("N2")
    kList.Sheets(3).Range("D2:D300").Copy Destination:=kList.Sheets(1).Range("M2")
    
    Dim last As Integer
    last = kList.Sheets(1).Range("G2").End(xlDown).Row
    
    For i = 2 To last
        kList.Sheets(1).Cells(i, 6).Value = "RemedyForce"
    Next i
    
    kList.Sheets(4).Range("A2:C300").Copy Destination:=kList.Sheets(1).Range("G" & last + 1)
    kList.Sheets(4).Range("AJ2:AJ300").Copy Destination:=kList.Sheets(1).Range("J" & last + 1)
    kList.Sheets(4).Range("J2:J300").Copy Destination:=kList.Sheets(1).Range("K" & last + 1)
    kList.Sheets(4).Range("Q2:Q300").Copy Destination:=kList.Sheets(1).Range("M" & last + 1)
    kList.Sheets(4).Range("Z2:Z300").Copy Destination:=kList.Sheets(1).Range("N" & last + 1)
    kList.Sheets(4).Range("G2:G300").Copy Destination:=kList.Sheets(1).Range("P" & last + 1)
    kList.Sheets(4).Range("AA2:AA300").Copy Destination:=kList.Sheets(1).Range("Q" & last + 1)
    
    For i = last + 1 To kList.Sheets(1).Range("G2").End(xlDown).Row
        kList.Sheets(1).Cells(i, 6).Value = "ChaRM"
    Next i
    
    'Remaining Change Manager Blanks filled with Klaus Karkov
    kList.Sheets(1).Range("Q2:Q" & last).Value = "Klaus Karkov"
    
End Sub

Private Sub formatColumns(kList)

    Dim i As Integer
    
    'Date formatted, status consolidated; Source filled.
    For i = 2 To 500
        kList.Sheets(1).Range("K" & CStr(i)).NumberFormat = "mm-dd-yyyy"
        kList.Sheets(1).Cells(i, 11).Value = Format(kList.Sheets(1).Cells(i, 11).Value, "dd-mm-yyyy")
        If (kList.Sheets(1).Cells(i, 9).Value = "REGISTERED") Or (kList.Sheets(1).Cells(i, 9).Value = "ASSIGNED") Or (kList.Sheets(1).Cells(i, 9).Value = "Created") Or (kList.Sheets(1).Cells(i, 9).Value = "In Developer QA") Then
            kList.Sheets(1).Cells(i, 9).Value = "Entry"
        ElseIf (kList.Sheets(1).Cells(i, 9).Value = "PLANNED") Or (kList.Sheets(1).Cells(i, 9).Value = "IN PROGRESS") Or (kList.Sheets(1).Cells(i, 9).Value = "Developer QA Completed") Or (kList.Sheets(1).Cells(i, 9).Value = "In Build") Or (kList.Sheets(1).Cells(i, 9).Value = "Build Completed") Or (kList.Sheets(1).Cells(i, 9).Value = "Unit Testing In Process") Or (kList.Sheets(1).Cells(i, 9).Value = "Functional SignOff Failed") Then
            kList.Sheets(1).Cells(i, 9).Value = "In Progress"
        ElseIf (kList.Sheets(1).Cells(i, 9).Value = "WAITING FOR?") Or (kList.Sheets(1).Cells(i, 9).Value = "WAITING FOR…") Or (kList.Sheets(1).Cells(i, 9).Value = "In Build - Awaiting Input") Then
            kList.Sheets(1).Cells(i, 9).Value = "Blocked"
        ElseIf (kList.Sheets(1).Cells(i, 9).Value = "Unit Testing Passed") Or (kList.Sheets(1).Cells(i, 9).Value = "Functional SignOff in Process") Or (kList.Sheets(1).Cells(i, 9).Value = "Functional SignOff Completed") Then
            kList.Sheets(1).Cells(i, 9).Value = "Exit"
        End If
    
    Next i
    
    'Fill Workloads for developers
    For i = 2 To kList.Sheets(1).Range("G2").End(xlDown).Row
        kList.Sheets(1).Range("L" & CStr(i)).Value = "=COUNTIF(M2:M600, M" & CStr(i) & ")"
    Next i
    
    'Autofit width of columns (Description is manual
    kList.Sheets(1).Range("A:G").Columns.AutoFit
    Worksheets(1).Columns("H").ColumnWidth = 30
    'Worksheets(1).Range("H2").WrapText = True
    kList.Sheets(1).Range("I:J").Columns.AutoFit
    kList.Sheets(1).Range("L:Q").Columns.AutoFit
    'kList.Sheets(1).Rows("2:2").Select
    'ActiveWindow.FreezePanes = True
    kList.Sheets(1).Range("A2").Select
    
End Sub
Private Sub Merge(kList)
    Dim i, j As Integer
    For i = 2 To kList.Sheets(1).Range("G2").End(xlDown).Row
        For j = 2 To kList.Sheets(1).Range("G2").End(xlDown).Row
            If kList.Sheets(1).Cells(j, 9).Value = "Exit" Then
                kList.Sheets(1).Rows(j).Delete
                j = j - 1
            ElseIf kList.Sheets(1).Cells(i, 7).Value = kList.Sheets(2).Cells(j, 7).Value Then
                kList.Sheets(2).Range("A" & CStr(j) & ":E" & CStr(j)).Copy Destination:=kList.Sheets(1).Range("A" & CStr(i) & ":E" & CStr(i))
            End If
        Next j
    Next i
        
End Sub

Private Sub deleteYesterday(kList)

    Application.DisplayAlerts = False
    kList.Sheets(2).Delete
    Application.DisplayAlerts = True

End Sub

