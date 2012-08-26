VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_raporlar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
On Error GoTo Err_Command2_Click


    DoCmd.Close

Exit_Command2_Click:
    Exit Sub

Err_Command2_Click:
    MsgBox Err.Description
    Resume Exit_Command2_Click
    
End Sub
Private Sub Command14_Click()
On Error GoTo Err_Command14_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command14_Click:
    Exit Sub

Err_Command14_Click:
    MsgBox Err.Description
    Resume Exit_Command14_Click
    
End Sub
Private Sub Command50_Click()
On Error GoTo Err_Command50_Click

    Dim stDocName As String

    stDocName = "msl_rapor_kisisel"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command50_Click:
    Exit Sub

Err_Command50_Click:
    MsgBox Err.Description
    Resume Exit_Command50_Click
    
End Sub
Private Sub Command51_Click()
On Error GoTo Err_Command51_Click

    Dim stDocName As String

    stDocName = "msl_rapor_kisisel"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command51_Click:
    Exit Sub

Err_Command51_Click:
    MsgBox Err.Description
    Resume Exit_Command51_Click
    
End Sub
Private Sub Command62_Click()
On Error GoTo Err_Command62_Click

    Dim stDocName As String

    stDocName = "msl_rapor_tumfabrika"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command62_Click:
    Exit Sub

Err_Command62_Click:
    MsgBox Err.Description
    Resume Exit_Command62_Click
    
End Sub
Private Sub Command72_Click()
On Error GoTo Err_Command72_Click

    Dim stDocName As String

    stDocName = "karne"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command72_Click:
    Exit Sub

Err_Command72_Click:
    MsgBox Err.Description
    Resume Exit_Command72_Click
    
End Sub
Private Sub Command87_Click()
On Error GoTo Err_Command87_Click

    Dim stDocName As String

    stDocName = "msl_rapor_takim"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command87_Click:
    Exit Sub

Err_Command87_Click:
    MsgBox Err.Description
    Resume Exit_Command87_Click
    
End Sub
Private Sub Command88_Click()
On Error GoTo Err_Command88_Click

    Dim stDocName As String

    stDocName = "msl_rapor_takim"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command88_Click:
    Exit Sub

Err_Command88_Click:
    MsgBox Err.Description
    Resume Exit_Command88_Click
    
End Sub
Private Sub Command89_Click()
On Error GoTo Err_Command89_Click

    Dim stDocName As String

    stDocName = "karne_ozet"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command89_Click:
    Exit Sub

Err_Command89_Click:
    MsgBox Err.Description
    Resume Exit_Command89_Click
    
End Sub
Private Sub Command90_Click()
On Error GoTo Err_Command90_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command90_Click:
    Exit Sub

Err_Command90_Click:
    MsgBox Err.Description
    Resume Exit_Command90_Click
    
End Sub
Private Sub Command91_Click()
On Error GoTo Err_Command91_Click

    Dim stDocName As String

    stDocName = ChrW(97) & ChrW(121) & ChrW(108) & ChrW(305) & ChrW(107) & ChrW(32) & ChrW(112) & ChrW(117) & ChrW(97) & ChrW(110)
    DoCmd.OpenReport stDocName, acPreview

Exit_Command91_Click:
    Exit Sub

Err_Command91_Click:
    MsgBox Err.Description
    Resume Exit_Command91_Click
    
End Sub
Private Sub Command92_Click()
On Error GoTo Err_Command92_Click

    Dim stDocName As String

    stDocName = "karne_2011"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command92_Click:
    Exit Sub

Err_Command92_Click:
    MsgBox Err.Description
    Resume Exit_Command92_Click
    
End Sub
Private Sub Command93_Click()
On Error GoTo Err_Command93_Click

    Dim stDocName As String

    stDocName = "msl_rapor_takim_ozet"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command93_Click:
    Exit Sub

Err_Command93_Click:
    MsgBox Err.Description
    Resume Exit_Command93_Click
    
End Sub
