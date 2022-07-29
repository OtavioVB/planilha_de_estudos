Private Sub btn_ConfigFecharMetas_Click()
Unload UserForm_ConfigMETAS
UserForm_Config.Show
End Sub

Private Sub btn_SalvarConfigMetas_Click()

Application.ScreenUpdating = False
Sheets("CONFIGURAÇÃO").Unprotect Password:="ENDM10707045"

If tb_MetasSemanaisHoras <> "" And tb_MetasSemanaisQuest <> "" Then
Dim Quests As Double
horassemanais = Format(tb_MetasSemanaisHoras.Value, "hh:mm:ss")
Quests = tb_MetasSemanaisQuest


Sheets("CONFIGURAÇÃO").Range("C32").Value = horassemanais
Sheets("CONFIGURAÇÃO").Range("C33").Value = Quests
If MsgBox("Configurações realizadas com sucesso!", vbOKOnly + vbInformation) = vbOK Then
Unload UserForm_ConfigMETAS
UserForm_Config.Show
End If
ElseIf tb_MetasSemanaisHoras = "" And tb_MetasSemanaisQuest = "" Then
If MsgBox("Você não preencheu todos os espaços!", vbOKOnly + vbCritical) = vbOK Then
End If
End If

Sheets("CONFIGURAÇÃO").Protect Password:="ENDM10707045"
Application.ScreenUpdating = True

End Sub


Private Sub tb_MetasSemanaisHoras_Change()

If Len(tb_MetasSemanaisHoras.Text) = 2 Then
tb_MetasSemanaisHoras = tb_MetasSemanaisHoras + ":00:00"
End If

End Sub

Private Sub tb_MetasSemanaisHoras_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Len(tb_MetasSemanaisHoras.Text) = 3 And KeyCode = 8 Then
tb_MetasSemanaisHoras = ""
End If
If Len(tb_MetasSemanaisHoras.Text) = 6 And KeyCode = 8 Then
tb_MetasSemanaisHoras = ""
End If

End Sub

Private Sub UserForm_Initialize()
    tb_MetasSemanaisHoras.MaxLength = 8
    tb_MetasSemanaisQuest.MaxLength = 3
End Sub