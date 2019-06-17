Attribute VB_Name = "Module1"
Sub Unit2_1_VBAEasy()
  
 'Set up Unique Ticker Population
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("J:J").RemoveDuplicates Columns:=1, Header:= _
        xlNo

'Calculate Volume for each Ticker
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
    
    For x = 2 To LastRow
    Dim VolSum As Double

    If Cells(x, 10) = Cells(x, 10).Text Then
    VolSum = Application.SumIf(Range("A:A"), Cells(x, 10), Range("G:G"))
    Cells(x, 11).Value = VolSum
    
    End If
    
    Next x


'Populate Column Titles and Format
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Range("J1:K1").Select
    Selection.Font.Bold = True
 
    Columns("J:K").Select
    Columns("J:K").EntireColumn.AutoFit
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    Range("J1").Select
    
    
End Sub

