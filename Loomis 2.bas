Attribute VB_Name = "Loomis_Cash_Order"
Public LastRow As Long
Sub Main()

    TurnOnPerformanceEnhancers
        SortAndRemoveUnwantedColumns
        RemoveUnwantedRows
        SpaceOutRows
        AutoSum
        MoveTotalsToEachCustomerFirstLine
        RemoveDuplicateAndBlankRows
        MoveDataToNewSheet
        GrandTotals
        Goodbye
        TurnOffPerformanceEnhancers

End Sub

Sub TurnOnPerformanceEnhancers()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    ActiveSheet.DisplayPageBreaks = False
                
End Sub

Sub TurnOffPerformanceEnhancers()
 
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayAlerts = True
                
End Sub

Sub GetLastRow()

    LastRow = Cells(Rows.Count, "A").End(xlUp).Row

End Sub

Sub SortAndRemoveUnwantedColumns()

    GetLastRow
        
    Sheets(1).Range("B:D,F:H,K:K,N:N,Q:BN").EntireColumn.Delete
        
    Columns("B:B").Cut
    Columns("A:A").Insert Shift:=xlToRight
        
    Cells.EntireColumn.AutoFit
    Sheets(1).Sort.SortFields.Add Key:=Range("C2", "C" & LastRow + 1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheets(1).Sort
        .SetRange Range("A1", "H" & LastRow + 1)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub RemoveUnwantedRows()

    GetLastRow
    
    For i = LastRow To 1 Step -1
        If Range("C" & i).Value < 100000 Then
            Range("C" & i).EntireRow.Delete
        ElseIf Range("C" & i).Value = "123456789" Then
            Range("C" & i).EntireRow.Delete
        ElseIf Range("C" & i).Value = "1555559999" Then
            Range("C" & i).EntireRow.Delete
        End If
    Next i

End Sub

Sub SpaceOutRows()

    Dim i As Long, j As String, k As String
        
    GetLastRow
    i = 2
        
    While i < LastRow
    j = Range("C" & i).Value
    k = Range("C" & i + 1).Value
        If j = k Then
            i = i + 1
        Else: Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
              Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
              i = i + 3
        End If
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Wend

End Sub

Sub AutoSum()
     
    sColumn = "G"
    i = 1

    While i < 3
        For Each NumRange In Columns(sColumn).SpecialCells(xlConstants, xlNumbers).Areas
          SumAddr = NumRange.Address(False, False)
         NumRange.Offset(NumRange.Count, 0).Resize(1, 1).Formula = "=SUM(" & SumAddr & ")"
        Next NumRange
     
    sColumn = "H"
    i = i + 1

    Wend

End Sub

Sub MoveTotalsToEachCustomerFirstLine()

    Dim i As Long, sFirstLine As String

    GetLastRow

    i = 2
    
        While i < LastRow
            sFirstLine = Range("G" & i).Address
            While Range("G" & i).Value <> vbNullString
                i = i + 1
            Wend
        
        Range(sFirstLine) = Range("G" & i - 1).Value
        i = i + 1
        Wend

    i = 2
        While i < LastRow
            sFirstLine = Range("H" & i).Address
                While Range("H" & i).Value <> vbNullString
                i = i + 1
                Wend
        
        Range(sFirstLine) = Range("H" & i - 1).Value
        i = i + 1
        
        Wend
    
End Sub

Sub RemoveDuplicateAndBlankRows()

    GetLastRow
        
    ActiveSheet.Range("A1", "H" & LastRow).RemoveDuplicates Columns:=3, Header:= _
    xlYes
    
    For i = LastRow To 1 Step -1
        If Range("A" & i).Value = vbNullString Then
        Range("A" & i).EntireRow.Delete
        End If
    Next i

End Sub

Sub MoveDataToNewSheet() 'To remove unwanted blank cells(haven't found another way to do this)

    GetLastRow
    
    Sheets.Add After:=ActiveSheet
    Sheets(1).Activate
    Sheets(1).Range("A1", "H" & LastRow).Copy Destination:=Sheets(2).Range("A1")
    Sheets(1).Delete
    Sheets(1).Name = "Loomis Cash Order CT"
    Range("A1:H1").VerticalAlignment = xlCenter
    Range("A1:H1").HorizontalAlignment = xlCenter

End Sub

Sub GrandTotals()

GetLastRow

    For Each NumRange In Columns("G").SpecialCells(xlConstants, xlNumbers).Areas
        SumAddr = NumRange.Address(False, False)
        NumRange.Offset(NumRange.Count, 0).Resize(1, 1).Formula = "=SUM(" & SumAddr & ")"
    Next NumRange
    
    For Each NumRange In Columns("H").SpecialCells(xlConstants, xlNumbers).Areas
        SumAddr = NumRange.Address(False, False)
        NumRange.Offset(NumRange.Count, 0).Resize(1, 1).Formula = "=SUM(" & SumAddr & ")"
    Next NumRange
    
        Cells.Font.Size = 10
        Cells.Font.Name = "Calibri"
        
        Range("D" & LastRow + 1) = "GRAND TOTALS"
    Rows(LastRow + 1).Font.Name = "Calibri"
    Rows(LastRow + 1).Font.Size = 11
    Rows(LastRow + 1).Font.Bold = True
        
        Range("A1", "H" & LastRow + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("A1", "H" & LastRow + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        Range("A" & LastRow + 1, "H" & LastRow + 1).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
        
        Range("G" & LastRow, "H" & LastRow + 1).NumberFormat = "$#,##0.00"
    
        Cells.EntireColumn.AutoFit

        Range("A1").Activate
        
End Sub

Sub Goodbye()

    sPrompt = "[" & ChrW(&H2022) & "_" & ChrW(&H2022) & "]" & "   I'm done        "
    MsgBox sPrompt + "", , "Loomis Cash Order™    "
    
End Sub
