
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
    moveYesterdayRows kanbanList
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

Private Sub moveYesterdayRows(kList)

    Dim last As Integer
    last = kList.Sheets(1).Range("F2").End(xlDown).Row
    
    kList.Sheets(1).Range("A2:Q" & CStr(last)).Cut kList.Sheets(1).Range("R2")

End Sub
Private Sub getTodayRows(kList)
    
    Dim i As Integer
    
    'Combine all rows from today & yesterday.
    kList.Sheets(2).Range("I2:I300").Copy Destination:=kList.Sheets(1).Range("G2")
    kList.Sheets(2).Range("F2:F300").Copy Destination:=kList.Sheets(1).Range("P2")
    kList.Sheets(2).Range("A2:B300").Copy Destination:=kList.Sheets(1).Range("H2")
    kList.Sheets(2).Range("G2:G200").Copy Destination:=kList.Sheets(1).Range("J2")
    kList.Sheets(2).Range("C2:C300").Copy Destination:=kList.Sheets(1).Range("K2")
    kList.Sheets(2).Range("E2:E300").Copy Destination:=kList.Sheets(1).Range("O2")
    kList.Sheets(2).Range("H2:H300").Copy Destination:=kList.Sheets(1).Range("N2")
    kList.Sheets(2).Range("D2:D300").Copy Destination:=kList.Sheets(1).Range("M2")
    
    Dim last As Integer
    last = kList.Sheets(1).Range("G2").End(xlDown).Row
    
    For i = 2 To last
        kList.Sheets(1).Cells(i, 6).Value = "RemedyForce"
    Next i
    
    kList.Sheets(3).Range("AT2:AT300").Copy Destination:=kList.Sheets(1).Range("D" & last + 1)
    kList.Sheets(3).Range("A2:C300").Copy Destination:=kList.Sheets(1).Range("G" & last + 1)
    kList.Sheets(3).Range("AJ2:AJ300").Copy Destination:=kList.Sheets(1).Range("J" & last + 1)
    kList.Sheets(3).Range("J2:J300").Copy Destination:=kList.Sheets(1).Range("K" & last + 1)
    kList.Sheets(3).Range("Q2:Q300").Copy Destination:=kList.Sheets(1).Range("M" & last + 1)
    kList.Sheets(3).Range("Z2:Z300").Copy Destination:=kList.Sheets(1).Range("N" & last + 1)
    kList.Sheets(3).Range("G2:G300").Copy Destination:=kList.Sheets(1).Range("P" & last + 1)
    kList.Sheets(3).Range("AA2:AA300").Copy Destination:=kList.Sheets(1).Range("Q" & last + 1)
    
    For i = last + 1 To kList.Sheets(1).Range("G2").End(xlDown).Row
        kList.Sheets(1).Cells(i, 6).Value = "ChaRM"
    Next i
    
    'Remaining Change Manager Blanks filled with Klaus Karkov
    kList.Sheets(1).Range("Q2:Q" & last).Value = "Klaus Karkov"
    
End Sub

Private Sub formatColumns(kList)

    Dim i As Integer
    Dim last As Integer
    last = kList.Sheets(1).Range("G2").End(xlDown).Row
    
    'Date formatted, status consolidated; Source filled.
    For i = 2 To last
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
    For i = 2 To last
        kList.Sheets(1).Range("L" & CStr(i)).Value = "=COUNTIF(M2:M600, M" & CStr(i) & ")"
    Next i
    
    'Enter Release info for ChaRM records
    For i = 2 To last
        If kList.Sheets(1).Cells(i, 4).Value = 4 Or kList.Sheets(1).Cells(i, 4).Value = 5 Or kList.Sheets(1).Cells(i, 4).Value = 6 Then
            kList.Sheets(1).Cells(i, 4).Value = "R" & CStr(kList.Sheets(1).Cells(i, 4).Value)
        ElseIf InStr(1, kList.Sheets(1).Cells(i, 4).Value, "MT") = 1 Then
            kList.Sheets(1).Cells(i, 4).Value = "maint"
        End If
    Next i
    
    'Resolve blank dates
    For i = 2 To last
        If kList.Sheets(1).Cells(i, 11).Value = "-" Then
            kList.Sheets(1).Cells(i, 11).Value = ""
        End If
    Next i
    
    'Color critical dates
    Dim days As Long
    Dim limit As Long
    Dim current As Date
    Dim utdate As Date
    limit = 3
    current = now()
    For i = 2 To last
        If kList.Sheets(1).Cells(i, 11).Value = "" Then
            If kList.Sheets(1).Cells(i, 9).Value <> "Entry" Then
                kList.Sheets(1).Cells(i, 11).Interior.ColorIndex = 46
            End If
        Else
            utdate = kList.Sheets(1).Cells(i, 11).Value
            days = DateDiff("d", current, utdate)
            If days <= limit And days >= 0 Then
                kList.Sheets(1).Cells(i, 11).Interior.ColorIndex = 36
            ElseIf days < 0 And kList.Sheets(1).Cells(i, 9).Value <> "Entry" Then
                kList.Sheets(1).Cells(i, 11).Interior.ColorIndex = 46
            End If
        End If
    Next i
    
    'Autofit width of columns (Description is manual)
    kList.Sheets(1).Range("A:G").Columns.AutoFit
    Worksheets(1).Columns("H").ColumnWidth = 30
    Worksheets(1).Columns("D").ColumnWidth = 16
    kList.Sheets(1).Range("I:J").Columns.AutoFit
    kList.Sheets(1).Range("L:Q").Columns.AutoFit
    'kList.Sheets(1).Range("A2").Select
    
End Sub
Private Sub Merge(kList)
    Dim i, j As Integer
    Dim lastNew As Integer
    Dim lastOld As Integer
    lastNew = kList.Sheets(1).Range("G2").End(xlDown).Row
    lastOld = kList.Sheets(1).Range("X2").End(xlDown).Row
    
    For i = 2 To lastNew
        For j = 2 To lastOld
            If kList.Sheets(1).Cells(j, 9).Value = "Exit" Then
                kList.Sheets(1).Rows(j).Delete
                j = j - 1
            ElseIf kList.Sheets(1).Cells(i, 7).Value = kList.Sheets(1).Cells(j, 24).Value Then
                kList.Sheets(1).Range("R" & CStr(j) & ":T" & CStr(j)).Copy Destination:=kList.Sheets(1).Range("A" & CStr(i) & ":C" & CStr(i))
                kList.Sheets(1).Range("V" & CStr(j)).Copy Destination:=kList.Sheets(1).Range("E" & CStr(i))
            End If
        Next j
    Next i
        
End Sub

Private Sub deleteYesterday(kList)

    Application.DisplayAlerts = False
    
    'Remove old data
    Dim last As Integer
    last = kList.Sheets(1).Range("X2").End(xlDown).Row
    kList.Sheets(1).Range("R2:AH" & CStr(last)).Delete
    
    'Rename the sheet with timestamp
    kList.Sheets(1).Name = Format(now(), "DD-MM-YYYY hh-nn-ss")
    
    Application.DisplayAlerts = True

End Sub

