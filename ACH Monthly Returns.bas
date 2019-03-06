Attribute VB_Name = "ACHMonthlyReturns"
Option Explicit
Sub ACHMonthlyReturns()
Dim LastRow As Long, i As Long
Application.DisplayAlerts = False
Application.ScreenUpdating = False
LastRow = Cells(Rows.Count, "A").End(xlUp).Row

'//Removes all unwanted data, only leaves Company IDs and their Debits
    On Error GoTo Continue
    Rows("1:1").Insert (xlDown)
    Columns("A:A").AutoFilter
    Sheets(1).Range("A1:A" & LastRow).AutoFilter Field:=1, Criteria1:= _
        "=*COMPANY ID:*", Operator:=xlOr, Criteria2:="=*DEBIT*"
    Selection.SpecialCells(xlCellTypeVisible).Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("MONTHLY ACH ORIG RETURN PERCENT").Delete
    Cells.EntireColumn.AutoFit
    Rows("1:1").Delete (xlUp)
    Range("A1").Select

'//Move even rows to column C
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    For i = LastRow To 1 Step -2
        Range("A" & i).Cut Destination:=Range("C" & i - 1)
        Range("A" & i, "C" & i).Delete (xlUp)
    Next i
    
'//Text to Columns
    Cells.Replace What:="COMPANY ID:", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="COMPANY NAME:", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(14, 1)), TrailingMinusNumbers:=True
    Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(5, 1), Array(24, 1), Array(43, 1), Array(58, 1), _
        Array(72, 1), Array(91, 1), Array(102, 1), Array(118, 1)), TrailingMinusNumbers:= _
        True
    Columns("C:C").Delete (xlToLeft)
    Cells.EntireColumn.AutoFit

'//Move data to new sheet to remove empty rows
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Sheets.Add After:=ActiveSheet
    Sheets(1).Range("A1", "J" & LastRow).Copy Destination:=Sheets(2).Range("A1")
    Sheets(1).Delete
    Sheets(1).Name = "Sheet1"
    
'//Move Totals(first row) to the end
    Range("A" & LastRow + 1, "J" & LastRow + 1).Value = Range("A1:J1").Value
    Range("A1:J1").Delete (xlUp)
    Cells.EntireColumn.AutoFit
    
'//Create return rate % sheets
    Rows("1:1").Insert (xlDown)
    Columns("A:J").AutoFilter
    Sheets.Add After:=ActiveSheet
    Sheets.Add After:=ActiveSheet
    Sheets.Add After:=ActiveSheet
    Sheets(1).Select
    ActiveSheet.Range("A1:J" & LastRow).AutoFilter Field:=8, Criteria1:=">=0.5%" _
        , Operator:=xlAnd
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets(2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Delete (xlUp)
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Range("A1:J" & LastRow).AutoFilter Field:=8
    ActiveSheet.Range("A1:J" & LastRow).AutoFilter Field:=10, Criteria1:=">=3%", _
        Operator:=xlAnd
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets(3).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Delete (xlUp)
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Range("A1:J" & LastRow).AutoFilter Field:=10
    ActiveSheet.Range("A1:J" & LastRow).AutoFilter Field:=6, Criteria1:=">=15%", _
        Operator:=xlAnd
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets(4).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Delete (xlUp)
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets(1).Select
    Range("A1").Select
    Rows("1:1").Delete (xlUp)
    
'//Format sheets
    Cells.Copy
    Sheets(2).Activate
    Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("H1:H" & LastRow).Interior.Color = RGB(255, 255, 0)
    Sheets(1).Activate
    Cells.Copy
    Sheets(3).Activate
    Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("J1:J" & LastRow).Interior.Color = RGB(255, 255, 0)
    Sheets(1).Activate
    Cells.Copy
    Sheets(4).Activate
    Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Range("F1:F" & LastRow).Interior.Color = RGB(255, 255, 0)
    Sheets(1).Activate
    Sheets(1).Name = "Totals"
    Sheets(2).Name = "Unauthorized Debit Entries"
    Sheets(3).Name = "Administrative Debit Returns"
    Sheets(4).Name = "Overall Debit Returns"
    
Continue:
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
