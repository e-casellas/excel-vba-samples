Option Explicit
Sub IRD_Rejector()

    TurnOnPerformanceEnhancers
    DeleteAllRowsExcept771s
    MoveDataToSingleRows
    TextToColumns
    DeleteAndReorderColumns
    SumVolumesAndRemoveDuplicateAccounts
    TransferDataToNewSheet
    SetHeadersAndCalculateTotalVolume
    FormatSheet
    EnterTotalsInMasterWorkbook
    Goodbye
    TurnOffPerformanceEnhancers

End Sub

Function LastRow() As Long

    LastRow = Workbooks(2).Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row

End Function

Private Sub TurnOnPerformanceEnhancers()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    ActiveSheet.DisplayPageBreaks = False
                
End Sub

Private Sub TurnOffPerformanceEnhancers()
 
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayAlerts = True
                
End Sub

Private Sub Goodbye()

    Dim sPrompt As String
    
    sPrompt = "[" & ChrW(&H2022) & "_" & ChrW(&H2022) & "]" & "   I'm done        "
    MsgBox sPrompt + "", , "IRD Rejector™    "
        
End Sub

Private Sub DeleteAllRowsExcept771s()

    Dim DestinationRow As Long, i As Long
    
    DestinationRow = 1
    
    For i = 1 To LastRow Step 1
        If Range("A" & i).Text Like "*REMOTE DEPOSIT PER ITEM FEE*" And Range("A" & i).Text Like "*REJECT*" Then
        Range("A" & i, "A" & i + 1).Copy Destination:=Range("B" & DestinationRow)
        DestinationRow = DestinationRow + 2
        End If
    Next i
    
    Columns("A:A").Delete Shift:=xlToLeft
    
End Sub

Private Sub MoveDataToSingleRows()

    Dim i As Long
    
    For i = 2 To LastRow Step 1
        Range("A" & i).Cut Destination:=Range("H" & i - 1)
        Rows(i & ":" & i).Delete Shift:=xlUp
    Next i
    
End Sub

Private Sub TextToColumns()

    Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(8, 1), Array(70, 1)), TrailingMinusNumbers:= _
        True
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(4, 1), Array(21, 1), Array(32, 1), Array(69, 1), _
        Array(83, 1), Array(116, 1)), TrailingMinusNumbers:=True
    Cells.EntireColumn.AutoFit
    
End Sub

Private Sub DeleteAndReorderColumns()

    Columns("I:I").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Range("D:D,F:F,I:I,J:J").Delete Shift:=xlToLeft

End Sub

Private Sub SumVolumesAndRemoveDuplicateAccounts()

    Dim i As Long

    For i = 1 To LastRow Step 1
        Range("G" & i).Formula = "=SUMIF(B1:B" & LastRow & ",B" & i & ",C1:C" & LastRow & ")"
    Next i
    
    Range("C1", "C" & LastRow).Value = Range("G1", "G" & LastRow).Value
    Columns("G:G").Delete Shift:=xlToLeft
    
    Range("A1:F" & LastRow).RemoveDuplicates Columns:=2, Header:=xlNo

End Sub

Private Sub TransferDataToNewSheet()

    Dim sSheetName As String

    sSheetName = Sheets(1).Name
    Sheets.Add After:=ActiveSheet
    Sheets(1).Range("A1", "F" & LastRow).Copy Destination:=Sheets(2).Range("A1")
    Sheets(1).Delete
    Sheets(1).Name = sSheetName
    
End Sub

Private Sub SetHeadersAndCalculateTotalVolume()

    Rows("1:1").Insert Shift:=xlDown
    
    Range("A1") = "SOURCE"
    Range("B1") = "ACCOUNT NUMBER"
    Range("C1") = "VOLUME"
    Range("D1") = "DESCRIPTION"
    Range("E1") = "MESSAGE"
    Range("F1") = "EFFECTIVE DATE"
    Range("B" & LastRow + 1) = "TOTAL"
    Range("C" & LastRow + 1).Formula = "=SUM(C1:C" & LastRow & ")"
    
    
End Sub

Private Sub FormatSheet()

    Range("A1:F1").Font.Bold = True
    Range("A1:F1").HorizontalAlignment = xlCenter
    Range("A1:F1").VerticalAlignment = xlCenter
    Range("A1:F1").WrapText = True
    Range("B" & LastRow + 1, "C" & LastRow + 1).Font.Bold = True
    
    Columns("A:A").ColumnWidth = 7.43
    Columns("B:B").ColumnWidth = 10.29
    Columns("C:C").ColumnWidth = 8.14
    Columns("F:F").ColumnWidth = 9.29
    Cells.EntireColumn.AutoFit
    Rows("1:1").EntireRow.AutoFit

End Sub

Private Sub EnterTotalsInMasterWorkbook()

    Dim i As Long, sAccountNumber As String, iVolume As Integer, rMasterCellRow As Long, _
    rMasterCellColumn As Long

    For i = 2 To LastRow Step 1
        On Error GoTo ErrorHandler
        
        sAccountNumber = Workbooks(2).Sheets(1).Range("B" & i).Value
        
        iVolume = Workbooks(2).Sheets(1).Range("C" & i).Value
        
        rMasterCellRow = Workbooks(3).Sheets(1).Cells.Find(What:=sAccountNumber, After:=Range("A1"), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Row
        
        rMasterCellColumn = Workbooks(3).Sheets(1).Cells(2, Columns.Count).End(xlToLeft).Column
        
        If IsEmpty(Workbooks(3).Sheets(1).Cells(rMasterCellRow, rMasterCellColumn)) Then
              Workbooks(3).Sheets(1).Cells(rMasterCellRow, rMasterCellColumn) = iVolume
        Else: Workbooks(3).Sheets(1).Cells(rMasterCellRow, rMasterCellColumn) = _
              Workbooks(3).Sheets(1).Cells(rMasterCellRow, rMasterCellColumn).Value + iVolume
        End If
        
Continue:
    Next i

Exit Sub

ErrorHandler:
    On Error GoTo -1
    Workbooks(2).Sheets(1).Range("B" & i, "C" & i).Interior.Color = RGB(255, 255, 0)
    GoTo Continue
    
End Sub
