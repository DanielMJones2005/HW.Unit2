VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionBox 
   Caption         =   "The VBA of Wall Street Process Selection"
   ClientHeight    =   6204
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6660
   OleObjectBlob   =   "1_OptionBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton3_Click()

Dim LB2 As Long

    Unload Me
    
    For LB2 = 0 To ListBox2.ListCount - 1
        
        'MsgBox (ListBox2.List(LB2))
        Worksheets(ListBox2.List(LB2)).Select
        Worksheets(ListBox2.List(LB2)).Activate
        Call Unit2_3_VBAHard
        
    Next LB2
 
    
    

    
End Sub

Private Sub CommandButton4_Click()
     
     Unload Me
     
     
     ActiveSheet.Range("A1").Select
        
        
End Sub

Private Sub CommandButton5_Click()
    Unload Me
    
    Call ProcessAll
End Sub

Private Sub CommandButton6_Click()
    
    Unload Me
    
    Call ResetWSAll
    
End Sub

Private Sub CommandButton7_Click()
    Unload Me
    Call ResetWS
End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub UserForm_Initialize()
Dim I As Integer
Dim N As Long

    For N = 1 To ActiveWorkbook.Sheets.Count
    ListBox1.AddItem ActiveWorkbook.Sheets(N).Name
    Next N
    
End Sub

Private Sub CommandButton1_Click()
Dim I As Integer

  

ListBox1.MultiSelect = 1

For I = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(I) = True Then ListBox2.AddItem ListBox1.List(I)
Next I

    
    

End Sub

Private Sub CommandButton2_Click()

Dim counter As Integer
counter = 0

For I = 0 To ListBox2.ListCount - 1
    If ListBox2.Selected(I - counter) Then
        ListBox2.RemoveItem (I - counter)
        counter = counter + 1
    End If
Next I


End Sub





