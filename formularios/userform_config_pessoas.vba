Private Sub btn_CancelarConfigPessoais_Click()
Unload UserForm_ConfigPessoais
UserForm_Config.Show
End Sub


Private Sub btn_SalvarConfigPessoais_Click()

Application.ScreenUpdating = False
Sheets("CONFIGURAÇÃO").Unprotect Password:="ENDM10707045"

If tb_UniversidadePessoal <> "" And tb_NomePessoal <> "" And tb_FrasePessoal <> "" And tb_CursoPessoal <> "" Then
Sheets("CONFIGURAÇÃO").Range("C45").Value = "JA ABRIU PELA PRIMEIRA VEZ"
Sheets("CONFIGURAÇÃO").Range("C46").Value = "Olá " + tb_NomePessoal + ", como você está hoje?"
Sheets("CONFIGURAÇÃO").Range("C47").Value = tb_FrasePessoal
Sheets("CONFIGURAÇÃO").Range("C48").Value = tb_CursoPessoal
Sheets("CONFIGURAÇÃO").Range("C49").Value = tb_UniversidadePessoal
ElseIf MsgBox("Você não preencheu todos os campos!", vbCritical) = vbOK Then
End If
Unload UserForm_ConfigPessoais
UserForm_Config.Show

Sheets("CONFIGURAÇÃO").Protect Password:="ENDM10707045"
Application.ScreenUpdating = True

End Sub

Private Sub tb_NomePessoal_Enter()
tb_NomePessoal.Value = ""
tb_NomePessoal.ForeColor = &H0&
End Sub
Private Sub tb_CursoPessoal_Enter()
tb_CursoPessoal.Value = ""
tb_CursoPessoal.ForeColor = &H0&
End Sub
Private Sub tb_UniversidadePessoal_Enter()
tb_UniversidadePessoal.Value = ""
tb_UniversidadePessoal.ForeColor = &H0&
End Sub
Private Sub tb_FrasePessoal_Enter()
tb_FrasePessoal.Value = ""
tb_FrasePessoal.ForeColor = &H0&
End Sub

Private Sub UserForm_Initialize()
tb_FrasePessoal.MultiLine = True
tb_CursoPessoal.Value = Sheets("CONFIGURAÇÃO").Range("C48").Value
tb_UniversidadePessoal.Value = Sheets("CONFIGURAÇÃO").Range("C49").Value
tb_NomePessoal.Value = Sheets("CONFIGURAÇÃO").Range("C46").Value
tb_FrasePessoal.Value = Sheets("CONFIGURAÇÃO").Range("C47").Value
End Sub
