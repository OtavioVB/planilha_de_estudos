Private Sub btn_EstudoCancelar_Click()

Unload UserForm_IniciarEstudos

End Sub

Private Sub btn_EstudoIniciar_Click()

Application.ScreenUpdating = False
Sheets("ESTUDOS").Unprotect Password:="ENDM10707045"

If cb_EstudoMateria.Value <> "" And tb_EstudoConteudo.Value <> "" And cb_TipoDeEstudo <> "" And IsNumeric(tb_PAGE) = True Then
Dim PageI As Double 'Converte para número
PageI = tb_PAGE 'Variável
Range("A5").Select
'-----------------------------------
If Range("A6").Value <> "" Then
Selection.End(xlDown).Select
End If
'-----------------------------------
If ActiveCell.Offset(0, 3).Value <> "" Then
ActiveCell.Offset(1, 0).Value = cb_EstudoMateria
ActiveCell.Offset(1, 1).Value = tb_EstudoConteudo
ActiveCell.Offset(1, 2).Value = Time$
ActiveCell.Offset(1, 5).Value = cb_TipoDeEstudo
ActiveCell.Offset(1, 6).Value = Date$
ActiveCell.Offset(1, 13).Value = PageI
Unload UserForm_IniciarEstudos
If cb_TipoDeEstudo = "Revisão" Or cb_TipoDeEstudo = "Estudo" Then
ActiveCell.Offset(1, 7).Value = "0"
ActiveCell.Offset(1, 8).Value = "0"
ActiveCell.Offset(1, 9).Value = "0"
End If
ElseIf MsgBox("Você não finalizou o último estudo!", vbCritical) = vbOK Then
End If
ElseIf IsNumeric(tb_PAGE) = False And Not tb_PAGE <> "" Then
If MsgBox("Você não colocou as páginas como um número", vbCritical) = vbOK Then
End If
ElseIf cb_EstudoMateria.Value = "" Or tb_EstudoConteudo.Value = "" Or cb_TipoDeEstudo = "" Or tb_PAGE <> "" Then
Unload UserForm_IniciarEstudos
If MsgBox("Você não preencheu todos os campos!", vbCritical) = vbOK Then
End If
End If

Range("A1").Select
Sheets("ESTUDOS").Protect Password:="ENDM10707045"
Application.ScreenUpdating = True
End Sub


Private Sub UserForm_Click()

End Sub
