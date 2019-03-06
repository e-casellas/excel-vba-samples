Public LastRow As Long
Sub Main()

    TurnOnPerformanceEnhancers
    SheetKiller
    SheetCleaner
    CopyPaste
    RemoveUnwantedSheets
    SortColumns
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

Sub SheetKiller()
    Dim sName As String
    Dim i As Long, wsCount As Long
        
    wsCount = Sheets.Count
    For i = wsCount To 1 Step -1
        sName = Sheets(i).Name
        If sName Like "*SUMMARY*" Then
                Sheets(i).Delete
        End If
    Next i
        
End Sub

Sub SheetCleaner()

    Dim i As Long, j As Long, wsCount As Long
    
    wsCount = Sheets.Count
    For i = wsCount To 1 Step -1
        Sheets(i).Range("B:D,H:I,K:L,N:N,Q:R,T:AG,AI:IH").EntireColumn.Delete
    Next i
        
        For j = wsCount To 2 Step -1
                Sheets(j).Range("1:1").EntireRow.Delete
        Next j

End Sub

Sub CopyPasta()

    Dim i As Long, wsCount As Long, PasteRow As Long, CopyRow As Long
    
    wsCount = Sheets.Count
    For i = wsCount To 2 Step -1
        PasteRow = Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row + 1
        CopyRow = Sheets(i).Cells(Rows.Count, "A").End(xlUp).Row
        Sheets(i).Range("A1", "J" & CopyRow).Copy Destination:=Sheets(1).Range("A" & PasteRow)
    Next i
        
End Sub

Sub RemoveUnwantedSheets()

    Dim i As Long, wsCount As Long
        
    wsCount = Sheets.Count
    For i = wsCount To 2 Step -1
                Sheets(i).Delete
    Next i

End Sub

Sub SortColumns()

    Columns("G:G").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Columns("G:G").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Columns("H:H").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("G:G").Cut
    Columns("E:E").Insert Shift:=xlToRight

End Sub

Sub RemoveUnwantedRows()

    GetLastRow
    
    For i = LastRow To 1 Step -1
        If Range("H" & i).Value = vbNullString Then
                        Range("H" & i).EntireRow.Delete
        ElseIf Range("J" & i).Value = "0" Then
                        Range("J" & i).EntireRow.Delete
        End If
    Next i

End Sub

Sub MoveDataToNewSheet() 'To remove unwanted blank cells(haven't found another way to do this)

        Sheets.Add After:=ActiveSheet
        Sheets(1).Activate
        GetLastRow
        Sheets(1).Range("A1", "J" & LastRow).Copy Destination:=Sheets(2).Range("A1")
        Sheets(1).Delete
        Sheets(1).Name = "CT"
		Range("A1:J1").VerticalAlignment = xlCenter
		Range("A1:J1").HorizontalAlignment = xlCenter

End Sub

Sub SpaceOutRows()

        Dim i As Long, j As String, k As String
        
        GetLastRow
    i = 2
        
    While i <> LastRow
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
     
sColumn = "I"
i = 1

While i < 3

    For Each NumRange In Columns(sColumn).SpecialCells(xlConstants, xlNumbers).Areas
        SumAddr = NumRange.Address(False, False)
        NumRange.Offset(NumRange.Count, 0).Resize(1, 1).Formula = "=SUM(" & SumAddr & ")"
    Next NumRange
     
sColumn = "J"
i = i + 1

Wend

End Sub

Sub MoveTotalsToEachCustomerFirstLine()

Dim i As Long, sFirstLine As String

GetLastRow

i = 2
    While i < LastRow
        sFirstLine = Range("I" & i).Address
            While Range("I" & i).Value <> vbNullString
                i = i + 1
            Wend
        Range(sFirstLine) = Range("I" & i - 1).Value
    i = i + 1
Wend

i = 2
    While i < LastRow
        sFirstLine = Range("J" & i).Address
            While Range("J" & i).Value <> vbNullString
                i = i + 1
            Wend
        Range(sFirstLine) = Range("J" & i - 1).Value
    i = i + 1
Wend

End Sub

Sub RemoveDuplicateAndBlankRows()

        GetLastRow
        
    ActiveSheet.Range("A1", "J" & LastRow).RemoveDuplicates Columns:=3, Header:= _
    xlYes
    
    For i = LastRow To 1 Step -1
        If Range("A" & i).Value = vbNullString Then
        Range("A" & i).EntireRow.Delete
        End If
    Next i

End Sub

Sub GrandTotals()

GetLastRow

    For Each NumRange In Columns("I").SpecialCells(xlConstants, xlNumbers).Areas
        SumAddr = NumRange.Address(False, False)
        NumRange.Offset(NumRange.Count, 0).Resize(1, 1).Formula = "=SUM(" & SumAddr & ")"
    Next NumRange
    
    For Each NumRange In Columns("J").SpecialCells(xlConstants, xlNumbers).Areas
        SumAddr = NumRange.Address(False, False)
        NumRange.Offset(NumRange.Count, 0).Resize(1, 1).Formula = "=SUM(" & SumAddr & ")"
    Next NumRange
    
        Range("E" & LastRow + 1) = "GRAND TOTALS"
    Rows(LastRow + 1).Font.Name = "Calibri"
    Rows(LastRow + 1).Font.Size = 9
    Rows(LastRow + 1).Font.Bold = True
        
        Range("A1", "J" & LastRow + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("A1", "J" & LastRow + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        Range("A" & LastRow + 1, "J" & LastRow + 1).Select
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
    Cells.EntireColumn.AutoFit
    ActiveWorkbook.Worksheets("CT").Sort.SortFields.Add Key:=Range("C2", "C" & LastRow + 1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("CT").Sort
        .SetRange Range("A1", "J" & LastRow + 1)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

        Range("A1").Activate
        
End Sub

Sub Goodbye()

        sPrompt = "[" & ChrW(&H2022) & "_" & ChrW(&H2022) & "]" & "   I'm done        "
        MsgBox sPrompt + "", , "Loomis OCR™    "
        
End Sub