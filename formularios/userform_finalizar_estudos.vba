Private Sub btn_FinalizarEst_Click()
Application.ScreenUpdating = False
Sheets("ESTUDOS").Unprotect Password:="ENDM10707045"

Dim PageFF As Double
PageFF = tb_UltimaPageEst

Range("A4").Select
If Range("A5").Select <> "" Then
Selection.End(xlDown).Select
End If
ActiveCell.Offset(0, 3).Value = Time$
ActiveCell.Offset(0, 12).Value = cb_NIVELDEDIFICULDADEEST
ActiveCell.Offset(0, 14).Value = PageFF
ConteudoR = ActiveCell.Offset(0, 1).Value
Dim Data As Double
Data = ActiveCell.Offset(0, 6).Value


'-------------------- PARTE DE REVISÕES ------------------------------------------------------
If cb1_revisoes = True Then
Sheets("TAREFAS").Unprotect Password:="ENDM10707045"
Sheets("TAREFAS").Select
Range("A3").Select
'-----------------------------------
If Range("A4").Value <> "" Then
Selection.End(xlDown).Select
End If
'------PRIMEIRA REVISÃO--------
ActiveCell.Offset(1, 0).Value = Date$
Dim PrimeiraRevisao As Double
Dim NovaData1 As Date
PrimeiraRevisao = Sheets("CONFIGURAÇÃO").Range("C15").Value
NovaData1 = VBA.Format(Data + PrimeiraRevisao, "dd/mm/yyyy")
ActiveCell.Offset(1, 1).Value = NovaData1
ActiveCell.Offset(1, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(1, 3).Value = "NÃO"
'------PRIMEIRA REVISÃO--------



'------SEGUNDA REVISÃO--------

Dim SegundaRevisao As Double
SegundaRevisao = Sheets("CONFIGURAÇÃO").Range("C16").Value
If SegundaRevisao <> "0" Then
Dim NovaData2 As Date
ActiveCell.Offset(2, 0).Value = Date$
NovaData2 = VBA.Format(Data + SegundaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(2, 1).Value = NovaData2
ActiveCell.Offset(2, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(2, 3).Value = "NÃO"
End If
'------SEGUNDA REVISÃO--------




'------TERCEIRA REVISÃO--------
Dim TerceiraRevisao As Double
TerceiraRevisao = Sheets("CONFIGURAÇÃO").Range("C17").Value
If TerceiraRevisao <> "0" Then
Dim NovaData3 As Date
ActiveCell.Offset(3, 0).Value = Date$
NovaData3 = VBA.Format(Data + TerceiraRevisao, "dd/mm/yyyy")
ActiveCell.Offset(3, 1).Value = NovaData3
ActiveCell.Offset(3, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(3, 3).Value = "NÃO"
End If
'------TERCEIRA REVISÃO--------


'------QUARTA REVISÃO--------
Dim QuartaRevisao As Double
QuartaRevisao = Sheets("CONFIGURAÇÃO").Range("C18").Value
If QuartaRevisao <> "0" Then
Dim NovaData4 As Date
ActiveCell.Offset(4, 0).Value = Date$
NovaData4 = VBA.Format(Data + QuartaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(4, 1).Value = NovaData4
ActiveCell.Offset(4, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(4, 3).Value = "NÃO"
End If
'------QUARTA REVISÃO--------


'------QUINTA REVISÃO--------
Dim QuintaRevisao As Double
QuintaRevisao = Sheets("CONFIGURAÇÃO").Range("C19").Value
If QuintaRevisao <> "0" Then
Dim NovaData5 As Date
ActiveCell.Offset(5, 0).Value = Date$
NovaData5 = VBA.Format(Data + QuintaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(5, 1).Value = NovaData5
ActiveCell.Offset(5, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(5, 3).Value = "NÃO"
End If
'------QUINTA REVISÃO--------


'------SEXTA REVISÃO--------
Dim SextaRevisao As Double
SextaRevisao = Sheets("CONFIGURAÇÃO").Range("C20").Value
If SextaRevisao <> "0" Then
Dim NovaData6 As Date
ActiveCell.Offset(6, 0).Value = Date$
NovaData6 = VBA.Format(Data + SextaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(6, 1).Value = NovaData6
ActiveCell.Offset(6, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(6, 3).Value = "NÃO"
End If
'------SEXTA REVISÃO--------


'------SETIMA REVISÃO--------
Dim SetimaRevisao As Double
SetimaRevisao = Sheets("CONFIGURAÇÃO").Range("C21").Value
If SetimaRevisao <> "0" Then
Dim NovaData7 As Date
ActiveCell.Offset(7, 0).Value = Date$
NovaData7 = VBA.Format(Data + SetimaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(7, 1).Value = NovaData7
ActiveCell.Offset(7, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(7, 3).Value = "NÃO"
End If
'------SETIMA REVISÃO--------


'------OITAVA REVISÃO--------
Dim OitavaRevisao As Double
OitavaRevisao = Sheets("CONFIGURAÇÃO").Range("C22").Value
If OitavaRevisao <> "0" Then
Dim NovaData8 As Date
ActiveCell.Offset(8, 0).Value = Date$
NovaData8 = VBA.Format(Data + OitavaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(8, 1).Value = NovaData8
ActiveCell.Offset(8, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(8, 3).Value = "NÃO"
End If
'------OITAVA REVISÃO--------


'------NONA REVISÃO--------
Dim NonaRevisao As Double
NonaRevisao = Sheets("CONFIGURAÇÃO").Range("C23").Value
If NonaRevisao <> "0" Then
Dim NovaData9 As Date
ActiveCell.Offset(9, 0).Value = Date$
NovaData9 = VBA.Format(Data + NonaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(9, 1).Value = NovaData9
ActiveCell.Offset(9, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(9, 3).Value = "NÃO"
End If
'------NONA REVISÃO--------


'------DECIMA REVISÃO--------
Dim DecimaRevisao As Double
DecimaRevisao = Sheets("CONFIGURAÇÃO").Range("C24").Value
If DecimaRevisao <> "0" Then
Dim NovaData10 As Date
ActiveCell.Offset(10, 0).Value = Date$
NovaData10 = VBA.Format(Data + DecimaRevisao, "dd/mm/yyyy")
ActiveCell.Offset(10, 1).Value = NovaData10
ActiveCell.Offset(10, 2).Value = "Revisão de " & ConteudoR
ActiveCell.Offset(10, 3).Value = "NÃO"
End If
'------DECIMA REVISÃO--------

ElseIf cb_NIVELDEDIFICULDADEEST = "" Then
MsgBox ("ERRO" & Chr(13) & "Todos os campos não foram preenchidos!")
End If
Sheets("ESTUDOS").Protect Password:="ENDM10707045"
Sheets("TAREFAS").Protect Password:="ENDM10707045"
Unload UserForm_FinalizarEst

Application.ScreenUpdating = False
End Sub

Private Sub CommandButton1_Click()

Unload UserForm_FinalizarEst

End Sub


Private Sub UserForm_Click()

End Sub
