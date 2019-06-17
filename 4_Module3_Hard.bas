Attribute VB_Name = "Module3"
Sub Unit2_3_VBAHard()
Attribute Unit2_3_VBAHard.VB_ProcData.VB_Invoke_Func = " \n14"

Call Unit2_2_VBAModerate



'Create Headers and Line Items
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Selection.Font.Bold = True
    
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Selection.Font.Bold = True
    
    Range("P1:Q1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Greatest % Increase"
    Selection.Font.Bold = True
    
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "Greatest % Decrease"
    Selection.Font.Bold = True
    
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "Greatest Total Volume"
    Selection.Font.Bold = True
    
    Columns("O:O").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("O:O").EntireColumn.AutoFit
  
    Range("P2").Select
    
'Identify Min Max % Changes
    Dim TickerLR2 As Long
    TickerLR2 = Cells(Rows.Count, 10).End(xlUp).Row
    
'FindMinMaxVal
    Dim GreatInc As Double
    Dim GreatDec As Double
    Dim GreatVol As Double
    
    MyRangeIncDec = ActiveSheet.Range(("L2"), ("L" & TickerLR2))
    MyRangeVol = ActiveSheet.Range(("M2"), ("M" & TickerLR2))
    
    GreatInc = Application.WorksheetFunction.Max(MyRangeIncDec)
    GreatDec = Application.WorksheetFunction.Min(MyRangeIncDec)
    GreatVol = Application.WorksheetFunction.Max(MyRangeVol)
            
    For k = 2 To TickerLR2
            
    If Cells(k, 12) = GreatInc Then
        GreatIncTicker = Cells(k, 10)
            
    ElseIf Cells(k, 12) = GreatDec Then
        GreatDecTicker = Cells(k, 10)
                
    ElseIf Cells(k, 13) = GreatVol Then
        GreatVolTicker = Cells(k, 10)
            
    End If
            
    Next k
    
    'Populate values
    Range("P2") = GreatIncTicker
    Range("P3") = GreatDecTicker
    Range("P4") = GreatVolTicker
    
    Range("Q2") = GreatInc
    Range("Q3") = GreatDec
    Range("Q4") = GreatVol
    
            
'Format Results
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Range("Q4").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
    Columns("P:Q").Select
    Columns("P:Q").EntireColumn.AutoFit
    
    Range("P1").Select
            
    Range("A1").Select
    
MsgBox ("Done!!!")

    
End Sub
