Attribute VB_Name = "Module2"
Sub Unit2_2_VBAModerate()

Call Unit2_1_VBAEasy

'Insert New Columns and Titles
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"
    Range("K3").Select
    
'Sort DataSet
    Dim LastRow As Long
    Dim ASName As String
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ASName = ActiveSheet.Name
    
    Columns("A:G").Select
    'ActiveWorkbook.Worksheets("A").Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets("A").Sort.SortFields.Add2 Key:=Range("A2:A" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'ActiveWorkbook.Worksheets("A").Sort.SortFields.Add2 Key:=Range("B2:B" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
    ActiveWorkbook.Worksheets(ASName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ASName).Sort.SortFields.Add2 Key:=Range("A2:A" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(ASName).Sort.SortFields.Add2 Key:=Range("B2:B" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
    'With ActiveWorkbook.Worksheets("A").Sort
    With ActiveWorkbook.Worksheets(ASName).Sort
        .SetRange Range("A1:G" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Caculate Yearly Change and % Change for each ticker
    'TickerLR = Column 10 (J)
    Dim TickerLR As Long
    TickerLR = Cells(Rows.Count, 10).End(xlUp).Row
    For I = 2 To TickerLR
    
        'Count Ticker
            Dim TickerCount As Variant
            TickerCount = Application.WorksheetFunction.CountIf(Range("A1:A" & LastRow), Cells(I, 10))

        'Determine Last Cell of Ticker Range (Column A)
            Dim FirstRowRange As Double
            Dim LastRowRange As Double
            
            Columns("A:A").Select
        
            Selection.Find(What:=Cells(I, 10), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
               :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
               False, SearchFormat:=False).Activate
               
            FirstRowRange = ActiveCell.Row
            LastRowRange = FirstRowRange + TickerCount - 1
   
        'FindMinMaxVal
            Dim MinVal As Integer
            Dim MaxVal As Integer
    
            MyRange = ActiveSheet.Range(("B" & FirstRowRange), ("B" & LastRowRange))
    
            BeginDate = Application.WorksheetFunction.Min(MyRange)
            EndDate = Application.WorksheetFunction.Max(MyRange)
       
        'Determine Open | Close Price
           Dim OpenPrice As Double
           Dim ClosePrice As Double
           
           For x = FirstRowRange To LastRowRange
           
           If Cells(x, 2) = BeginDate Then
                OpenPrice = Cells(x, 3)
                
                ElseIf Cells(x, 2) = EndDate Then
                    ClosePrice = Cells(x, 6)
                End If
            Next x
    
       
        'Yearly Change
            Dim YrChng As Double
    
            YrChng = ClosePrice - OpenPrice
            Cells(I, 11).Value = YrChng
    
        '% Change
            Dim PercentChng As Double
            
            If OpenPrice = 0 Then
                Cells(I, 12).Value = 0
                
            Else
                PercentChng = (YrChng / OpenPrice)
                Cells(I, 12).Value = PercentChng
            End If
    
    Next I
    
    
'Format Yearly Change and % Change
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = _
        "_(* #,##0.00000000000_);_(* (#,##0.00000000000);_(* ""-""??_);_(@_)"
    Columns("K:K").EntireColumn.AutoFit
    
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Columns("L:L").Select
    Columns("L:L").EntireColumn.AutoFit
    Range("L1").Select

'Conditional Format Yearly Change
    
    For j = 2 To TickerLR

    If Cells(j, 11).Value > 0 Then
        Cells(j, 11).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65280
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    ElseIf Cells(j, 11).Value < 0 Then
         Cells(j, 11).Select
         With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
         End With
         
         With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
         End With
    
    End If
    
    Next j
    
Range("J1").Select
  
    
    
End Sub


