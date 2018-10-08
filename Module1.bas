Attribute VB_Name = "Module1"
Option Explicit

Sub CreateWorksheets()
Attribute CreateWorksheets.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CreateWorksheets Macro
'
    Workbooks.Add
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Corporate Summary"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Monthly Rentals"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "Assets in Inertia"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "Monthly ONs"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "Monthly OFFs"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet6").Select
    Sheets("Sheet6").Name = "Raw"
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\CaltexOutput", FileFormat:=51
    Windows("CaltexMacroMaster.xlsm").Activate
End Sub

Sub Copy()
'
' Copy data
'
    Application.DisplayAlerts = False
    Workbooks.Open (ThisWorkbook.Path & "\CaltexData.xlsx")
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("CaltexOutput.xlsx").Activate
    Sheets("Raw").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("U:U").Select
    Selection.TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
    Columns("V:V").Select
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
    Workbooks("CaltexData.xlsx").Close
    Application.DisplayAlerts = True
    Windows("CaltexMacroMaster.xlsm").Activate
End Sub
Sub Split()
'
' Split data up
'
Dim d1 As Date
Dim d2 As Date
Dim d3 As Date
    d1 = Workbooks("CaltexMacroMaster.xlsm").Sheets("Macro").Range("B3")
    d2 = Format(CDate(d1), "mm/dd/yyyy")
    d3 = Format(CDate(d1), "dd/mm/yyyy")
    Windows("CaltexOutput.xlsx").Activate
    Sheets("Raw").Range("A1:AR1").AutoFilter Field:=22, Criteria1:="<" & d2
    Sheets("Raw").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Assets in Inertia").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Raw").Activate
    ActiveSheet.ShowAllData
    Sheets("Raw").Range("A1:AR1").AutoFilter Field:=22, Criteria1:=">=" & d2
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Monthly Rentals").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Raw").Activate
    ActiveSheet.ShowAllData
    Sheets("Raw").Range("A1:AR1").AutoFilter Field:=21, Criteria1:="=" & d3
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Monthly ONs").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Windows("CaltexMacroMaster.xlsm").Activate
End Sub
Sub FormatTotal()
'
' Format the data and total the sheets
'
Dim L1 As Long
Dim L2 As Long
Dim L3 As Long
Dim L4 As Long
Dim L5 As Long
Dim L6 As Long
Dim L7 As Long
Dim L8 As Long
Dim L9 As Long
    Windows("CaltexOutput.xlsx").Activate
    Sheets("Monthly Rentals").Activate
    Union(Range("AN:AN,AO:AO,AP:AP,AQ:AQ,AR:AR,C:C,D:D,E:E,F:F,G:G,H:H,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q,R:R,T:T,W:W,X:X,Y:Y,Z:Z,AA:AA,AB:AB,AC:AC,AD:AD,AF:AF,AG:AG,AH:AH") _
    , Range("AI:AI,AJ:AJ,AK:AK,AL:AL,AM:AM")).Select
    Range("AR1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "RentGST"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*0.1"
    Range("E2").Select
    Selection.AutoFill Destination:=Range(ActiveCell, ActiveCell.Offset(0, -1).End(xlDown).Offset(0, 1))
    Range("E2", Range("E2").End(xlDown)).Select
    Selection.NumberFormat = "0.00"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("E2").Select
    Application.CutCopyMode = False
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Rent(Inc GST)"
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("F2").Select
    Selection.AutoFill Destination:=Range(ActiveCell, ActiveCell.Offset(0, -1).End(xlDown).Offset(0, 1))
    Range("F2", Range("F2").End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    L1 = Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & L1 + 1).Formula = "=SUM(D2:D" & L1 & ")"
    L2 = Range("E" & Rows.Count).End(xlUp).Row
    Range("E" & L2 + 1).Formula = "=SUM(E2:E" & L2 & ")"
    L3 = Range("F" & Rows.Count).End(xlUp).Row
    Range("F" & L3 + 1).Formula = "=SUM(F2:F" & L3 & ")"
    Cells.Select
    With Selection.Font
        .Name = "Verdana"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("A2").Select
'
'
    Sheets("Assets in Inertia").Activate
    Union(Range("AN:AN,AO:AO,AP:AP,AQ:AQ,AR:AR,C:C,D:D,E:E,F:F,G:G,H:H,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q,R:R,T:T,W:W,X:X,Y:Y,Z:Z,AA:AA,AB:AB,AC:AC,AD:AD,AF:AF,AG:AG,AH:AH") _
    , Range("AI:AI,AJ:AJ,AK:AK,AL:AL,AM:AM")).Select
    Range("AR1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "RentGST"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*0.1"
    Range("E2").Select
    Selection.AutoFill Destination:=Range(ActiveCell, ActiveCell.Offset(0, -1).End(xlDown).Offset(0, 1))
    Range("E2", Range("E2").End(xlDown)).Select
    Selection.NumberFormat = "0.00"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("E2").Select
    Application.CutCopyMode = False
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Rent(Inc GST)"
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("F2").Select
    Selection.AutoFill Destination:=Range(ActiveCell, ActiveCell.Offset(0, -1).End(xlDown).Offset(0, 1))
    Range("F2", Range("F2").End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    L4 = Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & L4 + 1).Formula = "=SUM(D2:D" & L4 & ")"
    L5 = Range("E" & Rows.Count).End(xlUp).Row
    Range("E" & L5 + 1).Formula = "=SUM(E2:E" & L5 & ")"
    L6 = Range("F" & Rows.Count).End(xlUp).Row
    Range("F" & L6 + 1).Formula = "=SUM(F2:F" & L6 & ")"
    Cells.Select
    With Selection.Font
        .Name = "Verdana"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("A2").Select
'
'
    Sheets("Monthly ONs").Activate
    Union(Range("AN:AN,AO:AO,AP:AP,AQ:AQ,AR:AR,C:C,D:D,E:E,F:F,G:G,H:H,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q,R:R,T:T,W:W,X:X,Y:Y,Z:Z,AA:AA,AB:AB,AC:AC,AD:AD,AF:AF,AG:AG,AH:AH") _
    , Range("AI:AI,AJ:AJ,AK:AK,AL:AL,AM:AM")).Select
    Range("AR1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "RentGST"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*0.1"
    Range("E2").Select
    Selection.AutoFill Destination:=Range(ActiveCell, ActiveCell.Offset(0, -1).End(xlDown).Offset(0, 1))
    Range("E2", Range("E2").End(xlDown)).Select
    Selection.NumberFormat = "0.00"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("E2").Select
    Application.CutCopyMode = False
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Rent(Inc GST)"
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("F2").Select
    Selection.AutoFill Destination:=Range(ActiveCell, ActiveCell.Offset(0, -1).End(xlDown).Offset(0, 1))
    Range("F2", Range("F2").End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    L7 = Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & L7 + 1).Formula = "=SUM(D2:D" & L7 & ")"
    L8 = Range("E" & Rows.Count).End(xlUp).Row
    Range("E" & L8 + 1).Formula = "=SUM(E2:E" & L8 & ")"
    L9 = Range("F" & Rows.Count).End(xlUp).Row
    Range("F" & L9 + 1).Formula = "=SUM(F2:F" & L9 & ")"
    Cells.Select
    With Selection.Font
        .Name = "Verdana"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("A2").Select
    Sheets("Monthly Rentals").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("Assets in Inertia").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("Monthly ONs").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("Monthly OFFs").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    With ActiveWorkbook.Sheets("Monthly Rentals").Tab
        .Color = 65535
        .TintAndShade = 0
    End With
    With ActiveWorkbook.Sheets("Assets in Inertia").Tab
        .Color = 65535
        .TintAndShade = 0
    End With
    With ActiveWorkbook.Sheets("Monthly ONs").Tab
        .Color = 65535
        .TintAndShade = 0
    End With
    With ActiveWorkbook.Sheets("Monthly OFFs").Tab
        .Color = 65535
        .TintAndShade = 0
    End With
    Application.DisplayAlerts = False
    Worksheets("Raw").Delete
    Application.DisplayAlerts = True
    Windows("CaltexMacroMaster.xlsm").Activate
End Sub
Sub CorpSummary()
'
' Create Corporate Summary
'
    Windows("CaltexOutput.xlsx").Activate
    Sheets("Corporate Summary").Activate
    Range("B2").Select
    Selection.Font.Bold = True
    ActiveCell.FormulaR1C1 = _
        "Summary of Invoicing for rental period xxxx to xxxx for Caltex Corporate (i.e. Region = Caltex Australia Petroleum Pty Ltd)"
    Range("B4:I14").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("B4:I4").Select
    Selection.Font.Bold = True
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "Date Sent"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "Invoice #"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Date Due"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Amount Due Inc GST"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "Invoice Details"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "Overdue Status"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "Notes"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "Amount Received"
    Columns("B:I").Select
    Columns("B:I").EntireColumn.AutoFit
    Range("B1").Select
    Columns("B:B").ColumnWidth = 14.43
    Range("B19").Select
    Selection.Font.Bold = True
    ActiveCell.FormulaR1C1 = _
        "Summary of Monthly IT Internal Billing Report (EQG Draft) for xxxx for Caltex Corporate (i.e. Region = Caltex Australia Petroleum Pty Ltd)"
    Range("B20").Select
        Range("C22:D22").Select
    Selection.Merge
    Selection.Copy
    Range("C23:D27").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C22:E27").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("C22:D22").Select
    ActiveCell.FormulaR1C1 = "Monthly Rentals"
    Range("C23:D23").Select
    ActiveCell.FormulaR1C1 = "Assets in Inertia"
    Range("C24:D24").Select
    ActiveCell.FormulaR1C1 = "Monthly Ons"
    Range("C25:D25").Select
    ActiveCell.FormulaR1C1 = "Service Invoice"
    Range("C26:D26").Select
    ActiveCell.FormulaR1C1 = "Monthly OFFs"
    Range("C27:D27").Select
    ActiveCell.FormulaR1C1 = "Refund"
    Sheets("Monthly Rentals").Select
    Range("F1").Select
    Selection.End(xlDown).Select
    Selection.Copy
    Sheets("Corporate Summary").Select
    Range("E22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Assets in Inertia").Select
    Range("F1").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Corporate Summary").Select
    Range("E23").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Monthly ONs").Select
    Range("F1").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Corporate Summary").Select
    Range("E24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("C22:E27").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub ExecuteAll()
'
' Call all macros
'

Call CreateWorksheets
Call Copy
Call Split
Call FormatTotal
Call CorpSummary
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\CaltexOutput", FileFormat:=51
    Application.DisplayAlerts = True
End Sub
