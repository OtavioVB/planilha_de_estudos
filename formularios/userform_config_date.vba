
Private Sub btn_ConfigFechar_Click()
Unload UserForm_ConfigDATES
UserForm_Config.Show
End Sub

Private Sub btn_SalvarConfig_Click()

Application.ScreenUpdating = False
Sheets("CONFIGURAÇÃO").Unprotect Password:="ENDM10707045"

If tb_TerminoINSC <> "" And tb_1DIA <> "" And tb_2DIA <> "" Then
Sheets("CONFIGURAÇÃO").Range("C26").Value = tb_TerminoINSC
Sheets("CONFIGURAÇÃO").Range("C27").Value = tb_1DIA
Sheets("CONFIGURAÇÃO").Range("C28").Value = tb_2DIA
Unload UserForm_ConfigDATES
UserForm_Config.Show
ElseIf Not tb_TerminoINSC <> "" Or Not tb_1DIA <> "" Or Not tb_2DIA <> "" Then
If MsgBox("ERRO" & Chr(13) & "Você não preencheu todos os espaços de forma adequada!", vbCritical + vbOKOnly) = vbOK Then
End If
End If

Sheets("CONFIGURAÇÃO").Protect Password:="ENDM10707045"
Application.ScreenUpdating = True

End Sub

Private Sub tb_2DIA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If Len(tb_2DIA.Text) = 6 And KeyCode = 8 Then
tb_2DIA.Text = ""
End If
If Len(tb_2DIA.Text) = 3 And KeyCode = 8 Then
tb_2DIA.Text = ""
End If
End Sub
Private Sub tb_1DIA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If Len(tb_1DIA.Text) = 6 And KeyCode = 8 Then
tb_1DIA.Text = ""
End If
If Len(tb_1DIA.Text) = 3 And KeyCode = 8 Then
tb_1DIA.Text = ""
End If
End Sub
Private Sub tb_TerminoINSC_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If Len(tb_TerminoINSC.Text) = 6 And KeyCode = 8 Then
tb_TerminoINSC.Text = ""
End If
If Len(tb_TerminoINSC.Text) = 3 And KeyCode = 8 Then
tb_TerminoINSC.Text = ""
End If
End Sub

Private Sub tb_TerminoINSC_Change()
Me.tb_TerminoINSC.MaxLength = 10
If Len(tb_TerminoINSC.Text) = 2 Then
tb_TerminoINSC = tb_TerminoINSC + "/"
End If
If Len(tb_TerminoINSC.Text) = 5 Then
tb_TerminoINSC = tb_TerminoINSC + "/"
End If
End Sub

Private Sub tb_1DIA_Change()
Me.tb_1DIA.MaxLength = 10
If Len(tb_1DIA.Text) = 2 Then
tb_1DIA = tb_1DIA + "/"
End If
If Len(tb_1DIA.Text) = 5 Then
tb_1DIA = tb_1DIA + "/"
End If
End Sub
Private Sub tb_2DIA_Change()
Me.tb_2DIA.MaxLength = 10
If Len(tb_2DIA.Text) = 2 Then
tb_2DIA = tb_2DIA + "/"
End If
If Len(tb_2DIA.Text) = 5 Then
tb_2DIA = tb_2DIA + "/"
End If
End Sub

Private Sub UserForm_Click()

End Sub