VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sifre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command2_Click()
On Error GoTo Err_Command2_Click


    DoCmd.Close

Exit_Command2_Click:
    Exit Sub

Err_Command2_Click:
    MsgBox Err.Description
    Resume Exit_Command2_Click
    
End Sub
Private Sub Command6_Click()
On Error GoTo Err_Command6_Click


    DoCmd.Close

Exit_Command6_Click:
    Exit Sub

Err_Command6_Click:
    MsgBox Err.Description
    Resume Exit_Command6_Click
    
End Sub
Private Sub Command12_Click()
On Error GoTo Err_Command12_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
    
End Sub



Private Sub Command13_Click()
On Error GoTo Err_Command13_Click


    DoCmd.Quit

Exit_Command13_Click:
    Exit Sub

Err_Command13_Click:
    MsgBox Err.Description
    Resume Exit_Command13_Click
    
End Sub

Private Sub Command41_Click()

End Sub

Private Sub sifre_Exit(Cancel As Integer)

Dim db As Database, rs As Recordset, sql As String, ara As String, ara1 As String
    sql = "select * from sifre"
    ara = "adisoyadi = '" & Trim(Adisoyadi.Value) & "'"
    Set db = CodeDb
    Set rs = db.OpenRecordset(sql)
    rs.FindFirst (ara)
    If Not rs.NoMatch Then
    rs.MoveFirst
    ara1 = "adisoyadi = '" & Trim(Adisoyadi.Value) & "' and sifre = '" & Trim(sifre.Value) & "'"
    rs.FindFirst (ara1)
     If rs.NoMatch Then
     MsgBox ("Hatalý Þifre Girdiniz....")
     Else
        nsifreislemleri = rs.Fields("sifreislemleri")
        ndegisiklikislemleri = rs.Fields("duzeltmeislemleri")
     DoCmd.OpenForm "oncephe"
     End If
     Else
     MsgBox ("Kullanýcý Adý Hatalý..")
     End If
    
End Sub
