Attribute VB_Name = "Module4"


Sub ShowUserForm()
'https://www.excel-easy.com/vba/examples/multiple-list-box-selections.html


OptionBox.Show

End Sub

Sub Auto_Open()

OptionBox.Show

End Sub


Sub ResetWS()
'
' Reset Worksheet

    Columns("J:Q").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveWindow.FreezePanes = False
End Sub


Sub ProcessAll()
    Dim WS_Count As Integer
    Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
        For I = 1 To WS_Count
        Worksheets(I).Select

            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
        
        Call Unit2_3_VBAHard

        Next I
        
    MsgBox ("All Processed")

End Sub


Sub ResetWSAll()
    Dim WS_Count As Integer
    Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
        For I = 1 To WS_Count
        Worksheets(I).Select

            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
        
        Call ResetWS

        Next I
        
    MsgBox ("All Reset")

End Sub
