Sub AbrirIniciar()
  UserForm_IniciarEstudos.Show
End Sub

Sub IniciarFinal()
Sheets("ESTUDOS").Unprotect Password:="ENDM10707045"

Range("A5").Select

If Range("A6").Value <> "" Then
  Selection.End(xlDown).Select
End If
If ActiveCell.Offset(0, 3).Value = "" And ActiveCell.Offset(0, 12).Value = "" And ActiveCell.Offset(0, 1).Value <> "" Then
  If ActiveCell.Offset(0, 7).Value = "0" Then
    UserForm_FinalizarEst.Show
  Else
    UserForm_FinalizarExerce.Show
  End If
Else
  If MsgBox("Você não tem nenhum estudo para finalizar!", vbCritical) = vbOK Then
  End If
End If

Sheets("ESTUDOS").Protect Password:="ENDM10707045"
End Sub

Sub ExcluirPLN()
  Range("PLANNER").ClearContents
End Sub

Sub AbrirAddTarefas()
  UserForm_Tarefas.Show
End Sub

Sub ConfigsUser()
  UserForm_Config.Show
End Sub

