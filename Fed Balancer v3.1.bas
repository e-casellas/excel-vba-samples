Public bRejected As Boolean
Public LastRow493 As Long
Public LastRow513 As Long

Sub FED_Balancer()
bRejected = False
    
    Call TurnOnPerformanceEnhancers
    Call HighlightAndRemoveCells
    Call Highlighter2
    Call Highlighter3
    Call Totals
    Call Rejects
    Call GetLastRow493
    Call SeparateByTransactionType493
    Call DeleteBlankRows493
    Call TextToColumns493
    Call Titles493
    Call ReSortColumns493
    Call GetLastRow513
    Call SeparateByTransactionType513
    Call DeleteBlankRows513
    Call TextToColumns513
    Call Titles513
    Call ReSortColumns513
    Call Goodbye
    Call TurnOffPerformanceEnhancers

End Sub

Function LastRow() As Long

    LastRow = Workbooks(2).Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row

End Function

Private Sub GetLastRow493()
    
    LastRow493 = Workbooks(2).Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row

End Sub

Private Sub GetLastRow513()
    
    LastRow513 = Workbooks(2).Sheets(3).Cells(Rows.Count, "A").End(xlUp).Row

End Sub

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

Private Sub HighlightAndRemoveCells()

    For i = LastRow To 1 Step -1
    
        If Range("A" & i).Text Like "*STATEMENT OF YOUR OTHER SECONDARY RTNS' ACTIVITY*" Then
            Range("A" & i).Interior.Color = RGB(255, 255, 0)
        ElseIf Range("A" & i).Text Like "*STATEMENT OF YOUR SUBACCOUNTS' ACTIVITY*" Then
            Range("A" & i).Interior.Color = RGB(255, 255, 0)
        ElseIf Range("A" & i).Text Like "*DETAIL OF OWN ACTIVITY*" Then
            Range("A" & i).Interior.Color = RGB(255, 255, 0)
        ElseIf Range("A" & i).Text = "                                                                                " Then
            Range("A" & i).ClearContents
        ElseIf Range("A" & i).Text = "" Then
            Range("A" & i).ClearContents
        End If
    
    Next i
    
End Sub

Private Sub Highlighter2()

        Range("A1").Select

        
Line1:
        If ActiveCell.Row >= LastRow Then GoTo Omega
        If Selection.Interior.Color = RGB(255, 255, 0) Then GoTo Line2
        ActiveCell.Offset(1, 0).Select
        GoTo Line1
        
Line2:
        ActiveCell.Offset(2, 0).Select
        GoTo Line3
        
Line3:
        If ActiveCell.Text Like "*1130-0783-5   POPULAR BANK*" Then GoTo FOT
        If ActiveCell.Text Like "*0631-1542-4   POPULAR BANK*" Then GoTo FOT
        If ActiveCell.Text Like "*1119-2535-9   POPULAR BANK*" Then GoTo FOT
        If ActiveCell.Row >= LastRow Then GoTo Omega
        If Selection.Text Like "*[***]*" Then GoTo Line1
        ActiveCell.Offset(1, 0).Select
        If Not ActiveCell.Text Like "*7500*" Then GoTo Line3
        If ActiveCell.Text Like "*7500*" Then GoTo Line4
        
Line4:
        ActiveCell.Interior.Color = RGB(255, 255, 0)
        ActiveCell.Offset(1, 0).Select
        If IsEmpty(ActiveCell.Value) Then
                GoTo Line3
                Else: GoTo Line4
        End If
        
FOT:
        If ActiveCell.Row >= LastRow Then GoTo Omega
        If Selection.Text Like "*[***]*" Then GoTo Line1
        ActiveCell.Offset(1, 0).Select
        If Not ActiveCell.Text Like "*7500*" Then GoTo FOT
        If ActiveCell.Text Like "*7500*" Then GoTo FOTL1
        
FOTL1:
        ActiveCell.Interior.Color = RGB(255, 102, 255)
        ActiveCell.Offset(1, 0).Select
        If IsEmpty(ActiveCell.Value) Then
                GoTo FOT
                Else: GoTo FOTL1
        End If
        
Omega:
End Sub

Private Sub Highlighter3()

        Range("A" & LastRow).Select
Line1:
        ActiveCell.Offset(-1, 0).Select
        If ActiveCell.Address = ("$A$1") Then GoTo Omega
        If Selection.Interior.Color = RGB(255, 102, 255) Then GoTo Line2
        If Selection.Interior.Color <> RGB(255, 102, 255) Then GoTo Line1
        
Line2:
        If Selection.Interior.Color = RGB(255, 255, 0) Then GoTo Line3:
        If Selection.Interior.Color <> RGB(255, 255, 0) Then
                ActiveCell.Offset(-1, 0).Select
                If ActiveCell.Address = ("$A$1") Then GoTo Omega
                GoTo Line2
        End If
        
Line3:
        ActiveCell.Interior.Color = RGB(255, 102, 255)
        GoTo Line1
        
Omega:
End Sub

Private Sub Totals()

    Rows("1:1").Delete Shift:=xlUp
    Range("A1").Select
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Name = "493"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Name = "513"
    Sheets(1).Select
    Columns("A:A").Select
    Selection.AutoFilter
    Range("A1").Select
    ActiveSheet.Range("A1", "A" & LastRow).AutoFilter Field:=1, Criteria1:=RGB(255, _
        255, 0), Operator:=xlFilterCellColor
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy
    Sheets("493").Select
    ActiveSheet.Paste
    Cells.EntireColumn.AutoFit
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Range("A1", "A" & LastRow).AutoFilter Field:=1, Criteria1:=RGB(255, _
        102, 255), Operator:=xlFilterCellColor
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("513").Select
    ActiveSheet.Paste
    Cells.EntireColumn.AutoFit
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    Sheets(1).Select
    ActiveSheet.Range("A1", "A" & LastRow).AutoFilter Field:=1
    Range("A1").Select
    Cells.EntireColumn.AutoFit
End Sub

Private Sub Rejects()

    Dim i As Long

    For i = LastRow To 1 Step -1
    
        If Range("A" & i).Text Like "*Rejected*" Then
            Range("A" & i).Interior.Color = RGB(255, 0, 0)
            Range("A" & i + 1).Activate
                While ActiveCell.Value <> vbNullString
                    ActiveCell.Interior.Color = RGB(255, 0, 0)
                    ActiveCell.Offset(1, 0).Activate
                Wend
        
        bRejected = True
        
        End If
        
    Next i

End Sub

Private Sub SeparateByTransactionType493()

    Dim i As Long
    
    Workbooks(2).Sheets(2).Activate
    
    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*[*]*" Then
            Range("A" & i).ClearContents
        End If
    Next i
    
    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Credit Transaction Originated*" Then
            Do
                Range("A" & i).Cut Destination:=Range("C" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Debit Originated*" Then
            Do
                Range("A" & i).Cut Destination:=Range("E" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Credit Originated*" Then
            Do
                Range("A" & i).Cut Destination:=Range("G" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Debit Transaction Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("I" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i
    
    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Credit Transaction Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("K" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i
        
    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Debit Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("M" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Credit Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("O" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i
        
    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Immediate*" Then
            Do
                Range("A" & i).Cut Destination:=Range("Q" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Debit Transaction Rejected*" Then
            Do
                Range("A" & i).Cut Destination:=Range("T" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow493 Step 1
        If Range("A" & i).Text Like "*Credit Transaction Rejected*" Then
            Do
                Range("A" & i).Cut Destination:=Range("V" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

End Sub

Private Sub DeleteBlankRows493()

    Dim i As Long
    
    Workbooks(2).Sheets(2).Activate

    For i = LastRow493 To 1 Step -1
        If Range("A" & i).Value = vbNullString Or Range("A" & i).Text Like "*7500 (*" Then
            Range("A" & i).Delete Shift:=xlUp
        End If
        If Range("C" & i).Value = vbNullString Or Range("C" & i).Text Like "*7500 (*" Then
            Range("C" & i).Delete Shift:=xlUp
        End If
        If Range("G" & i).Value = vbNullString Or Range("G" & i).Text Like "*7500 (*" Then
            Range("G" & i).Delete Shift:=xlUp
        End If
        If Range("I" & i).Value = vbNullString Or Range("I" & i).Text Like "*7500 (*" Then
            Range("I" & i).Delete Shift:=xlUp
        End If
        If Range("K" & i).Value = vbNullString Or Range("K" & i).Text Like "*7500 (*" Then
            Range("K" & i).Delete Shift:=xlUp
        End If
        If Range("M" & i).Value = vbNullString Or Range("M" & i).Text Like "*7500 (*" Then
            Range("M" & i).Delete Shift:=xlUp
        End If
        If Range("O" & i).Value = vbNullString Or Range("O" & i).Text Like "*7500 (*" Then
            Range("O" & i).Delete Shift:=xlUp
        End If
        If Range("Q" & i).Value = vbNullString Or Range("Q" & i).Text Like "*7500 (*" Then
            Range("Q" & i).Delete Shift:=xlUp
        End If
        If Range("T" & i).Value = vbNullString Or Range("T" & i).Text Like "*7500 (*" Then
            Range("T" & i).Delete Shift:=xlUp
        End If
        If Range("V" & i).Value = vbNullString Or Range("V" & i).Text Like "*7500 (*" Then
            Range("V" & i).Delete Shift:=xlUp
        End If
    Next i
End Sub

Private Sub TextToColumns493()

    Workbooks(2).Sheets(2).Activate
    
    On Error Resume Next

    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(45, 1)), TrailingMinusNumbers:=True
    Columns("A:A").Delete Shift:=xlToLeft
    
    Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(37, 1)), TrailingMinusNumbers:=True
    Columns("B:B").Delete Shift:=xlToLeft
    
    Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(37, 1)), TrailingMinusNumbers:=True
    Columns("C:C").Delete Shift:=xlToLeft
    
    Columns("D:D").TextToColumns Destination:=Range("D1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(34, 1)), TrailingMinusNumbers:=True
    Columns("D:D").Delete Shift:=xlToLeft
    
    Columns("E:E").TextToColumns Destination:=Range("E1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(35, 1)), TrailingMinusNumbers:=True
    Columns("E:E").Delete Shift:=xlToLeft
    
    Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(47, 1)), TrailingMinusNumbers:=True
    Columns("F:F").Delete Shift:=xlToLeft
    
    Columns("G:G").TextToColumns Destination:=Range("G1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(36, 1)), TrailingMinusNumbers:=True
    Columns("G:G").Delete Shift:=xlToLeft
    
    Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(46, 1)), TrailingMinusNumbers:=True
    Columns("H:H").Delete Shift:=xlToLeft
    
    Columns("I:I").TextToColumns Destination:=Range("I1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(35, 1), Array(65, 1)), TrailingMinusNumbers _
        :=True
    Columns("I:I").Delete Shift:=xlToLeft

    Columns("K:K").TextToColumns Destination:=Range("K1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(36, 1)), TrailingMinusNumbers:=True
    Columns("K:K").Delete Shift:=xlToLeft
    
    Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(35, 1)), TrailingMinusNumbers:=True
    Columns("L:L").Delete Shift:=xlToLeft
    
    For i = LastRow493 To 1 Step -1
        If Range("A" & i).Value = vbNullString Then
            Range("A" & i).Delete Shift:=xlUp
        End If
        If Range("B" & i).Value = vbNullString Then
            Range("B" & i).Delete Shift:=xlUp
        End If
        If Range("C" & i).Value = vbNullString Then
            Range("C" & i).Delete Shift:=xlUp
        End If
        If Range("D" & i).Value = vbNullString Then
            Range("D" & i).Delete Shift:=xlUp
        End If
        If Range("E" & i).Value = vbNullString Then
            Range("E" & i).Delete Shift:=xlUp
        End If
        If Range("F" & i).Value = vbNullString Then
            Range("F" & i).Delete Shift:=xlUp
        End If
        If Range("G" & i).Value = vbNullString Then
            Range("G" & i).Delete Shift:=xlUp
        End If
        If Range("H" & i).Value = vbNullString Then
            Range("H" & i).Delete Shift:=xlUp
        End If
        If Range("I" & i).Value = vbNullString Then
            Range("I" & i).Delete Shift:=xlUp
        End If
        If Range("J" & i).Value = vbNullString Then
            Range("J" & i).Delete Shift:=xlUp
        End If
        If Range("K" & i).Value = vbNullString Then
            Range("K" & i).Delete Shift:=xlUp
        End If
        If Range("L" & i).Value = vbNullString Then
            Range("L" & i).Delete Shift:=xlUp
        End If
    Next i
    
    Columns("A:L").NumberFormat = "#,##0.00"
    Cells.EntireColumn.AutoFit

End Sub

Private Sub Titles493()

    Workbooks(2).Sheets(2).Activate

    Rows("1:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "ACH Transactions Originated"
    Range("A1:B1").HorizontalAlignment = xlCenter
    Range("A1:B1").Merge

    Range("C1").Value = "Same Day ACH Originated"
    Range("C1:D1").HorizontalAlignment = xlCenter
    Range("C1:D1").Merge
    
    Range("E1").Value = "ACH Transactions Received"
    Range("E1:F1").HorizontalAlignment = xlCenter
    Range("E1:F1").Merge

    Range("G1").Value = "Same Day ACH Received"
    Range("G1:H1").HorizontalAlignment = xlCenter
    Range("G1:H1").Merge

    Range("I1").Value = "ACH Immediate Transactions"
    Range("I1:J1").HorizontalAlignment = xlCenter
    Range("I1:J1").Merge

    Range("K1").Value = "ACH Transactions Rejected"
    Range("K1:L1").HorizontalAlignment = xlCenter
    Range("K1:L1").Merge

    Range("A2").Value = "Debits"
    Range("B2").Value = "Credits"
    Range("C2").Value = "Debits"
    Range("D2").Value = "Credits"
    Range("E2").Value = "Debits"
    Range("F2").Value = "Credits"
    Range("G2").Value = "Debits"
    Range("H2").Value = "Credits"
    Range("I2").Value = "Debits"
    Range("J2").Value = "Credits"
    Range("K2").Value = "Debits"
    Range("L2").Value = "Credits"
    If bRejected = False Then
        Range("K3").Value = "None"
        Range("L3").Value = "None"
    End If
    Columns("A:L").ColumnWidth = 12.71
    Range("A1:L2").Font.Bold = True

    Range("A1:L2").Borders.LineStyle = xlContinuous
    Range("A1:L2").Interior.Color = RGB(255, 255, 0)

End Sub

Private Sub ReSortColumns493()

    Workbooks(2).Sheets(2).Activate

    Range("B3", "B" & LastRow493).Cut
    Range("A3").Insert Shift:=xlToRight
    Range("D3", "D" & LastRow493).Cut
    Range("C3").Insert Shift:=xlToRight
    Application.CutCopyMode = False

End Sub

Private Sub SeparateByTransactionType513()

    Dim i As Long
    
    Workbooks(2).Sheets(3).Activate
    
    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*[*]*" Then
            Range("A" & i).ClearContents
        End If
    Next i
    
    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Credit Transaction Originated*" Then
            Do
                Range("A" & i).Cut Destination:=Range("C" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Debit Originated*" Then
            Do
                Range("A" & i).Cut Destination:=Range("E" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Credit Originated*" Then
            Do
                Range("A" & i).Cut Destination:=Range("G" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Debit Transaction Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("I" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i
    
    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Credit Transaction Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("K" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i
        
    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Debit Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("M" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Same Day ACH Credit Received*" Then
            Do
                Range("A" & i).Cut Destination:=Range("O" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i
        
    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Immediate*" Then
            Do
                Range("A" & i).Cut Destination:=Range("Q" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Debit Transaction Rejected*" Then
            Do
                Range("A" & i).Cut Destination:=Range("T" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

    For i = 1 To LastRow513 Step 1
        If Range("A" & i).Text Like "*Credit Transaction Rejected*" Then
            Do
                Range("A" & i).Cut Destination:=Range("V" & i)
                i = i + 1
            Loop Until Range("A" & i).Text Like "*7500 (*" Or Range("A" & i).Value = vbNullString
            i = i - 1
        End If
    Next i

End Sub

Private Sub DeleteBlankRows513()

    Workbooks(2).Sheets(3).Activate

    Dim i As Long

    For i = LastRow513 To 1 Step -1
        If Range("A" & i).Value = vbNullString Or Range("A" & i).Text Like "*7500 (*" Then
            Range("A" & i).Delete Shift:=xlUp
        End If
        If Range("C" & i).Value = vbNullString Or Range("C" & i).Text Like "*7500 (*" Then
            Range("C" & i).Delete Shift:=xlUp
        End If
        If Range("G" & i).Value = vbNullString Or Range("G" & i).Text Like "*7500 (*" Then
            Range("G" & i).Delete Shift:=xlUp
        End If
        If Range("I" & i).Value = vbNullString Or Range("I" & i).Text Like "*7500 (*" Then
            Range("I" & i).Delete Shift:=xlUp
        End If
        If Range("K" & i).Value = vbNullString Or Range("K" & i).Text Like "*7500 (*" Then
            Range("K" & i).Delete Shift:=xlUp
        End If
        If Range("M" & i).Value = vbNullString Or Range("M" & i).Text Like "*7500 (*" Then
            Range("M" & i).Delete Shift:=xlUp
        End If
        If Range("O" & i).Value = vbNullString Or Range("O" & i).Text Like "*7500 (*" Then
            Range("O" & i).Delete Shift:=xlUp
        End If
        If Range("Q" & i).Value = vbNullString Or Range("Q" & i).Text Like "*7500 (*" Then
            Range("Q" & i).Delete Shift:=xlUp
        End If
        If Range("T" & i).Value = vbNullString Or Range("T" & i).Text Like "*7500 (*" Then
            Range("T" & i).Delete Shift:=xlUp
        End If
        If Range("V" & i).Value = vbNullString Or Range("V" & i).Text Like "*7500 (*" Then
            Range("V" & i).Delete Shift:=xlUp
        End If
    Next i
End Sub

Private Sub TextToColumns513()

    Workbooks(2).Sheets(3).Activate
    
    On Error Resume Next

    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(45, 1)), TrailingMinusNumbers:=True
    Columns("A:A").Delete Shift:=xlToLeft
    
    Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(37, 1)), TrailingMinusNumbers:=True
    Columns("B:B").Delete Shift:=xlToLeft
    
    Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(37, 1)), TrailingMinusNumbers:=True
    Columns("C:C").Delete Shift:=xlToLeft
    
    Columns("D:D").TextToColumns Destination:=Range("D1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(34, 1)), TrailingMinusNumbers:=True
    Columns("D:D").Delete Shift:=xlToLeft
    
    Columns("E:E").TextToColumns Destination:=Range("E1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(35, 1)), TrailingMinusNumbers:=True
    Columns("E:E").Delete Shift:=xlToLeft
    
    Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(47, 1)), TrailingMinusNumbers:=True
    Columns("F:F").Delete Shift:=xlToLeft
    
    Columns("G:G").TextToColumns Destination:=Range("G1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(36, 1)), TrailingMinusNumbers:=True
    Columns("G:G").Delete Shift:=xlToLeft
    
    Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(46, 1)), TrailingMinusNumbers:=True
    Columns("H:H").Delete Shift:=xlToLeft
    
    Columns("I:I").TextToColumns Destination:=Range("I1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(35, 1), Array(65, 1)), TrailingMinusNumbers _
        :=True
    Columns("I:I").Delete Shift:=xlToLeft

    Columns("K:K").TextToColumns Destination:=Range("K1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(36, 1)), TrailingMinusNumbers:=True
    Columns("K:K").Delete Shift:=xlToLeft
    
    Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(35, 1)), TrailingMinusNumbers:=True
    Columns("L:L").Delete Shift:=xlToLeft
    
    For i = LastRow513 To 1 Step -1
        If Range("A" & i).Value = vbNullString Then
            Range("A" & i).Delete Shift:=xlUp
        End If
        If Range("B" & i).Value = vbNullString Then
            Range("B" & i).Delete Shift:=xlUp
        End If
        If Range("C" & i).Value = vbNullString Then
            Range("C" & i).Delete Shift:=xlUp
        End If
        If Range("D" & i).Value = vbNullString Then
            Range("D" & i).Delete Shift:=xlUp
        End If
        If Range("E" & i).Value = vbNullString Then
            Range("E" & i).Delete Shift:=xlUp
        End If
        If Range("F" & i).Value = vbNullString Then
            Range("F" & i).Delete Shift:=xlUp
        End If
        If Range("G" & i).Value = vbNullString Then
            Range("G" & i).Delete Shift:=xlUp
        End If
        If Range("H" & i).Value = vbNullString Then
            Range("H" & i).Delete Shift:=xlUp
        End If
        If Range("I" & i).Value = vbNullString Then
            Range("I" & i).Delete Shift:=xlUp
        End If
        If Range("J" & i).Value = vbNullString Then
            Range("J" & i).Delete Shift:=xlUp
        End If
        If Range("K" & i).Value = vbNullString Then
            Range("K" & i).Delete Shift:=xlUp
        End If
        If Range("L" & i).Value = vbNullString Then
            Range("L" & i).Delete Shift:=xlUp
        End If
    Next i
    
    Columns("A:L").NumberFormat = "#,##0.00"
    Cells.EntireColumn.AutoFit

End Sub

Private Sub Titles513()

    Workbooks(2).Sheets(3).Activate

    Rows("1:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "ACH Transactions Originated"
    Range("A1:B1").HorizontalAlignment = xlCenter
    Range("A1:B1").Merge

    Range("C1").Value = "Same Day ACH Originated"
    Range("C1:D1").HorizontalAlignment = xlCenter
    Range("C1:D1").Merge
    
    Range("E1").Value = "ACH Transactions Received"
    Range("E1:F1").HorizontalAlignment = xlCenter
    Range("E1:F1").Merge

    Range("G1").Value = "Same Day ACH Received"
    Range("G1:H1").HorizontalAlignment = xlCenter
    Range("G1:H1").Merge

    Range("I1").Value = "ACH Immediate Transactions"
    Range("I1:J1").HorizontalAlignment = xlCenter
    Range("I1:J1").Merge

    Range("K1").Value = "ACH Transactions Rejected"
    Range("K1:L1").HorizontalAlignment = xlCenter
    Range("K1:L1").Merge

    Range("A2").Value = "Debits"
    Range("B2").Value = "Credits"
    Range("C2").Value = "Debits"
    Range("D2").Value = "Credits"
    Range("E2").Value = "Debits"
    Range("F2").Value = "Credits"
    Range("G2").Value = "Debits"
    Range("H2").Value = "Credits"
    Range("I2").Value = "Debits"
    Range("J2").Value = "Credits"
    Range("K2").Value = "Debits"
    Range("L2").Value = "Credits"
    If bRejected = False Then
        Range("K3").Value = "None"
        Range("L3").Value = "None"
    End If
    Columns("A:L").ColumnWidth = 12.71
    Range("A1:L2").Font.Bold = True

    Range("A1:L2").Borders.LineStyle = xlContinuous
    Range("A1:L2").Interior.Color = RGB(255, 255, 0)

End Sub

Private Sub ReSortColumns513()

    Workbooks(2).Sheets(3).Activate

    Range("B3", "B" & LastRow493).Cut
    Range("A3").Insert Shift:=xlToRight
    Range("D3", "D" & LastRow493).Cut
    Range("C3").Insert Shift:=xlToRight
    Application.CutCopyMode = False

End Sub

Private Sub Goodbye()

    Workbooks(2).Sheets(1).Activate

    If bRejected = True Then
        MsgBox "Rejected transaction(s) found!", vbExclamation, "Attention"
    End If
    
    sPrompt = "[" & ChrW(&H2022) & "_" & ChrW(&H2022) & "]" & "   I'm done        "
    MsgBox sPrompt + "", , "Fed Balancer™    "
        
End Sub
