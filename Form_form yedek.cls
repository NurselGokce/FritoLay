VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_form yedek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Ay_BeforeUpdate(Cancel As Integer)
End Sub

Private Sub Option101_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text37.Visible = True
Text36.Visible = False
End Sub

Private Sub Option106_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text55.Visible = True
Text56.Visible = False
End Sub

Private Sub Option108_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text55.Visible = False
Text56.Visible = False
End Sub

Private Sub Option110_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text55.Visible = False
Text56.Visible = False
End Sub

Private Sub Option112_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text55.Visible = False
Text56.Visible = False
End Sub

Private Sub Option114_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text55.Visible = False
Text56.Visible = True
End Sub

Private Sub Option121_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text131.Visible = True
Text132.Visible = False
End Sub

Private Sub Option123_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text131.Visible = False
Text132.Visible = False
End Sub

Private Sub Option125_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text131.Visible = False
Text132.Visible = False
End Sub

Private Sub Option127_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text131.Visible = False
Text132.Visible = False
End Sub

Private Sub Option129_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text131.Visible = False
Text132.Visible = True
End Sub

Private Sub Option136_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text146.Visible = True
Text147.Visible = False
End Sub

Private Sub Option138_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text146.Visible = False
Text147.Visible = False
End Sub

Private Sub Option140_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text146.Visible = False
Text147.Visible = False
End Sub

Private Sub Option142_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text146.Visible = False
Text147.Visible = False
End Sub

Private Sub Option144_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text146.Visible = False
Text147.Visible = True
End Sub

Private Sub Option80_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text19.Visible = True
Text21.Visible = False
End Sub

Private Sub Option82_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text21.Visible = False
Text19.Visible = False
End Sub

Private Sub Option84_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text21.Visible = False
Text19.Visible = False
End Sub

Private Sub Option86_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text21.Visible = False
Text19.Visible = False
End Sub

Private Sub Option88_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text21.Visible = True
Text19.Visible = False
End Sub

Private Sub Option93_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text37.Visible = False
Text36.Visible = True
End Sub

Private Sub Option95_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text37.Visible = False
Text36.Visible = False
End Sub

Private Sub Option97_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text37.Visible = False
Text36.Visible = False
End Sub

Private Sub Option99_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text36.Visible = False
Text37.Visible = False
End Sub
Private Sub Command163_Click()
On Error GoTo Err_Command163_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command163_Click:
    Exit Sub

Err_Command163_Click:
    MsgBox Err.Description
    Resume Exit_Command163_Click
    
End Sub

Private Sub Command181_Click()
On Error GoTo Err_Command181_Click

    Dim stDocName As String

    stDocName = "aktarim"
    DoCmd.RunMacro stDocName

Exit_Command181_Click:
    Exit Sub

Err_Command181_Click:
    MsgBox Err.Description
    Resume Exit_Command181_Click

    
End Sub



Private Sub Image428_Click()
On Error GoTo Err_Image428_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Urunguvenligi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Image428_Click:
    Exit Sub

Err_Image428_Click:
    MsgBox Err.Description
    Resume Exit_Image428_Click
End Sub

Private Sub Image429_Click()
On Error GoTo Err_Image428_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Ýs Güvenliði_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Image428_Click:
    Exit Sub

Err_Image428_Click:
    MsgBox Err.Description
    Resume Exit_Image428_Click
End Sub

Private Sub Image430_Click()
On Error GoTo Err_Image428_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "takimcalismasi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Image428_Click:
    Exit Sub

Err_Image428_Click:
    MsgBox Err.Description
    Resume Exit_Image428_Click
End Sub

Private Sub Image431_Click()
On Error GoTo Err_Image428_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "disiplin_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Image428_Click:
    Exit Sub

Err_Image428_Click:
    MsgBox Err.Description
    Resume Exit_Image428_Click
End Sub

Private Sub Label404_Click()
Label233.Visible = True
Label245.Visible = True
Label246.Visible = True
Frame220.Visible = True
Frame234.Visible = True
Frame247.Visible = True
Box450.Visible = True
Box451.Visible = True
Box452.Visible = True
OLEUnbound424.Visible = True
End Sub

Private Sub Label406_Click()
Label260.Visible = True
Label273.Visible = True
Label285.Visible = True
Frame261.Visible = True
Frame274.Visible = True
Frame286.Visible = True
Box453.Visible = True
Box454.Visible = True
Box455.Visible = True
Image428.Visible = True
End Sub

Private Sub Label408_Click()
Label299.Visible = True
Label320.Visible = True
Label330.Visible = True
Frame311.Visible = True
Frame321.Visible = True
Frame331.Visible = True
Box456.Visible = True
Box457.Visible = True
Box458.Visible = True
Image429.Visible = True
End Sub

Private Sub Label414_Click()
Label360.Visible = True
Label361.Visible = True
Frame362.Visible = True
Frame369.Visible = True
Box462.Visible = True
Box463.Visible = True
Image431.Visible = True
End Sub

Private Sub Label418_Click()
Label342.Visible = True
Label358.Visible = True
Label359.Visible = True
Frame345.Visible = True
Frame380.Visible = True
Frame389.Visible = True
Box459.Visible = True
Box460.Visible = True
Box461.Visible = True
Image430.Visible = True
End Sub

Private Sub OLEUnbound424_Click()
On Error GoTo Err_OLEUnbound424_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "isbilgisi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OLEUnbound424_Click:
    Exit Sub

Err_OLEUnbound424_Click:
    MsgBox Err.Description
    Resume Exit_OLEUnbound424_Click
End Sub

Private Sub Option223_KeyPress(KeyAscii As Integer)


End Sub

Private Sub Option223_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = True
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option225_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = True
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option227_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = True
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False

End Sub

Private Sub Option229_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = True
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option237_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = True
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option239_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = True
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option241_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = True
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False

End Sub

Private Sub Option243_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = True
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option250_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = True
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option252_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = True
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option254_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = True
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option256_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = True
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option264_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = True
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option266_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = True
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False

End Sub

Private Sub Option268_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = True
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option277_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = True
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False

End Sub

Private Sub Option279_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = True
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False

End Sub

Private Sub Option281_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = True
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False

End Sub

Private Sub Option289_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = True
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False

End Sub

Private Sub Option291_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = True
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False

End Sub

Private Sub Option293_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = True
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option314_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = True
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False

End Sub

Private Sub Option316_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = True
End Sub

Private Sub Option318_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = True
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option324_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = True
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option326_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = True
End Sub

Private Sub Option328_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = True
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option334_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = True
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option336_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = True
End Sub

Private Sub Option338_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = True
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option348_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = True
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option350_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = True
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option352_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = True
disiplinkotu.Visible = False
End Sub

Private Sub Option365_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = True
End Sub

Private Sub Option367_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = True
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option372_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = True
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option374_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = False
disiplinaciklama.Visible = True
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option376_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = True
End Sub

Private Sub Option383_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = True
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option385_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = True
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option387_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = True
disiplinkotu.Visible = False
End Sub

Private Sub Option392_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = True
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = False
disiplinkotu.Visible = False
End Sub

Private Sub Option394_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbilgisiaciklama.Visible = False
kaliteaciklama.Visible = False
takimaciklama.Visible = True
disiplinaciklama.Visible = False
isguvenlilgiaciklama.Visible = False
End Sub

Private Sub Option396_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isbigisimukemmel.Visible = False
urunguvenligimukemmel.Visible = False
isguvenlilgimukemmel.Visible = False
takimmukkemmel.Visible = False
disiplinmukkemmel.Visible = False
isbigisikotu.Visible = False
urunguvenligikotu.Visible = False
isguvenlilgikotu.Visible = False
takimkotu.Visible = True
disiplinkotu.Visible = False
End Sub

Private Sub sicilno_Exit(Cancel As Integer)
    Rem Dim db As Database, rs As Recordset, sql As String, ara As String
    sql = "Select * from toplam"
    ara = "sicilno = " & Str(sicilno.Text)
    Set db = CodeDb
    Set rs = db.OpenRecordset(sql)
    rs.FindFirst (ara)
    If Not rs.NoMatch Then
        Adisoyadi.SetFocus
        Adisoyadi.Text = rs.Fields("Adisoyadi")
        vardiya_sefi.SetFocus
        vardiya_sefi.Text = rs.Fields("vardiya_sefi")
        bolum.SetFocus
        bolum.Text = rs.Fields("bolum")
    Else
        Adisoyadi.SetFocus
        Adisoyadi.Text = ""
        vardiya_sefi.SetFocus
        vardiya_sefi.Text = ""
        bolum.SetFocus
        bolum.Text = ""
     
    End If
        
    Ay.SetFocus
    Ay.Text = Month(Date) - 1
    yil.SetFocus
    yil.Text = Year(Date)
      
    Form.Refresh
    
End Sub
Private Sub Command188_Click()
On Error GoTo Err_Command188_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "isbilgisi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command188_Click:
    Exit Sub

Err_Command188_Click:
    MsgBox Err.Description
    Resume Exit_Command188_Click
    
End Sub
Private Sub Command189_Click()
On Error GoTo Err_Command189_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "takimcalismasi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command189_Click:
    Exit Sub

Err_Command189_Click:
    MsgBox Err.Description
    Resume Exit_Command189_Click
    
End Sub
Private Sub Command190_Click()
On Error GoTo Err_Command190_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Disiplin_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command190_Click:
    Exit Sub

Err_Command190_Click:
    MsgBox Err.Description
    Resume Exit_Command190_Click
    
End Sub
Private Sub Command191_Click()
On Error GoTo Err_Command191_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Urunguvenligi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command191_Click:
    Exit Sub

Err_Command191_Click:
    MsgBox Err.Description
    Resume Exit_Command191_Click
    
End Sub
Private Sub Command192_Click()
On Error GoTo Err_Command192_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = ChrW(304) & ChrW(115) & ChrW(32) & ChrW(71) & ChrW(252) & ChrW(118) & ChrW(101) & ChrW(110) & ChrW(108) & ChrW(105) & ChrW(287) & ChrW(105) & ChrW(95) & ChrW(104) & ChrW(97) & ChrW(107) & ChrW(107) & ChrW(105) & ChrW(110) & ChrW(100) & ChrW(97)
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command192_Click:
    Exit Sub

Err_Command192_Click:
    MsgBox Err.Description
    Resume Exit_Command192_Click
    
End Sub
Private Sub Command203_Click()
On Error GoTo Err_Command203_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "raporlar"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command203_Click:
    Exit Sub

Err_Command203_Click:
    MsgBox Err.Description
    Resume Exit_Command203_Click
    
End Sub
Private Sub Command205_Click()
On Error GoTo Err_Command205_Click


    DoCmd.Close

Exit_Command205_Click:
    Exit Sub

Err_Command205_Click:
    MsgBox Err.Description
    Resume Exit_Command205_Click
    
End Sub
Private Sub Command215_Click()
On Error GoTo Err_Command215_Click

    Dim stDocName As String

    stDocName = "karne"
    DoCmd.SendObject acReport, stDocName

Exit_Command215_Click:
    Exit Sub

Err_Command215_Click:
    MsgBox Err.Description
    Resume Exit_Command215_Click
    
End Sub
Private Sub Command216_Click()
On Error GoTo Err_Command216_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.SendObject acReport, stDocName

Exit_Command216_Click:
    Exit Sub

Err_Command216_Click:
    MsgBox Err.Description
    Resume Exit_Command216_Click
    
End Sub
Private Sub Command217_Click()
On Error GoTo Err_Command217_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.OutputTo acReport, stDocName

Exit_Command217_Click:
    Exit Sub

Err_Command217_Click:
    MsgBox Err.Description
    Resume Exit_Command217_Click
    
End Sub
Private Sub Command218_Click()
On Error GoTo Err_Command218_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.OutputTo acReport, stDocName

Exit_Command218_Click:
    Exit Sub

Err_Command218_Click:
    MsgBox Err.Description
    Resume Exit_Command218_Click
    
End Sub
Private Sub Command219_Click()
On Error GoTo Err_Command219_Click

    Dim stDocName As String

    stDocName = "msl_rapor"
    DoCmd.SendObject acReport, stDocName

Exit_Command219_Click:
    Exit Sub

Err_Command219_Click:
    MsgBox Err.Description
    Resume Exit_Command219_Click
    
End Sub
Private Sub Command426_Click()
On Error GoTo Err_Command426_Click

    Dim stDocName As String

    stDocName = "aktarim"
    DoCmd.RunMacro stDocName

Exit_Command426_Click:
    Exit Sub

Err_Command426_Click:
    MsgBox Err.Description
    Resume Exit_Command426_Click
    
End Sub
Private Sub Command427_Click()
On Error GoTo Err_Command427_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "isbilgisi_hakkinda"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command427_Click:
    Exit Sub

Err_Command427_Click:
    MsgBox Err.Description
    Resume Exit_Command427_Click
    
End Sub
Private Sub Command495_Click()
On Error GoTo Err_Command495_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "devam"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command495_Click:
    Exit Sub

Err_Command495_Click:
    MsgBox Err.Description
    Resume Exit_Command495_Click
    
End Sub
