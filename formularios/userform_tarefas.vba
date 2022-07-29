Private Sub btn_AddTar_Click()

Application.ScreenUpdating = False

If tb_DataDaTarefa <> "" And tb_TarefaRealizar <> "" Then
Range("A3").Select
If Range("a4").Value <> "" Then
Selection.End(xlDown).Select
End If
Dim DataCERTA As Date
On Error Resume Next
DataCERTA = tb_DataDaTarefa.Value
ActiveCell.Offset(1, 0).Value = Date$
ActiveCell.Offset(1, 1).Value = DataCERTA
ActiveCell.Offset(1, 2).Value = tb_TarefaRealizar
ActiveCell.Offset(1, 3).Value = "NÃO"
Unload UserForm_Tarefas
ElseIf tb_DataDaTarefa = "" Or tb_TarefaRealizar = "" Then
MsgBox ("ERRO" & Chr(13) & "Você não preencheu todos os espaços!")
ElseIf VBA.Format(tb_DataDaTarefa) <> "dd/mm/yyyy" Then
MsgBox ("ERRO" & Chr(13) & "Você não escreveu a data de maneira correta!" & Chr(13) & "EXEMPLO: 31/12/1900")
End If

Application.ScreenUpdating = True

End Sub

Private Sub btn_CancelarTar_Click()
Unload UserForm_Tarefas
End Sub

Private Sub tb_DataDaTarefa_Change()
tb_DataDaTarefa.MaxLength = 10
If Len(tb_DataDaTarefa.Text) = 2 Then
tb_DataDaTarefa = tb_DataDaTarefa + "/"
End If
If Len(tb_DataDaTarefa.Text) = 5 Then
tb_DataDaTarefa = tb_DataDaTarefa + "/"
End If

End Sub


Private Sub tb_DataDaTarefa_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If Len(tb_DataDaTarefa.Text) = 6 And KeyCode = 8 Then
tb_DataDaTarefa.Text = ""
End If
If Len(tb_DataDaTarefa.Text) = 4 And KeyCode = 8 Then
tb_DataDaTarefa.Text = ""
End If
End Sub

Private Sub tb_TarefaRealizar_Change()
tb_DataDaTarefa.MaxLength = 50
End Sub