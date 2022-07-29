

Private Sub btn_ConfigFechar_Click()
Unload UserForm_ConfigREV
UserForm_Config.Show
End Sub

Private Sub btn_SalvarConfig_Click()

Application.ScreenUpdating = False
Sheets("CONFIGURAÇÃO").Unprotect Password:="ENDM10707045"

Dim PrimeiraRevisao As Double
Dim SegundaRevisao As Double
Dim TerceiraRevisao As Double
Dim QuartaRevisao As Double
Dim QuintaRevisao As Double
Dim SextaRevisao As Double
Dim SetimaRevisao As Double
Dim OitavaRevisao As Double
Dim NonaRevisao As Double
Dim DecimaRevisao As Double
PrimeiraRevisao = tb_revisao1
SegundaRevisao = tb_revisao2
TerceiraRevisao = tb_revisao3
QuartaRevisao = tb_revisao4
QuintaRevisao = tb_revisao5
SextaRevisao = tb_revisao6
SetimaRevisao = tb_revisao7
OitavaRevisao = tb_revisao8
NonaRevisao = tb_revisao9
DecimaRevisao = tb_revisao10
Sheets("Configuração").Range("C15").Value = PrimeiraRevisao
Sheets("Configuração").Range("C16").Value = SegundaRevisao
Sheets("Configuração").Range("C17").Value = TerceiraRevisao
Sheets("Configuração").Range("C18").Value = QuartaRevisao
Sheets("Configuração").Range("C19").Value = QuintaRevisao
Sheets("Configuração").Range("C20").Value = SextaRevisao
Sheets("Configuração").Range("C21").Value = SetimaRevisao
Sheets("Configuração").Range("C22").Value = OitavaRevisao
Sheets("Configuração").Range("C23").Value = NonaRevisao
Sheets("Configuração").Range("C24").Value = DecimaRevisao
MsgBox ("Configurações realizadas com sucesso!" & Chr(13) & "Primeira revisão: " & PrimeiraRevisao & " dia(s) depois." & Chr(13) & "Segunda revisão: " & SegundaRevisao & " dia(s) depois." & Chr(13) & "Terceira revisão: " & TerceiraRevisao & " dia(s) depois." & Chr(13) & "Quarta Revisão: " & QuartaRevisao & " dia(s) depois." & Chr(13) & "Quinta Revisão: " & QuintaRevisao & " dia(s) depois." & Chr(13) & "Sexta Revisão: " & SextaRevisao & " dia(s) depois." & Chr(13) & "Setima Revisão: " & SetimaRevisao & " dia(s) depois." & Chr(13) & "Oitava Revisão: " & OitavaRevisao & " dia(s) depois." & Chr(13) & "Nona Revisão: " & NonaRevisao & " dia(s) depois." & Chr(13) & "Decima Revisão: " & DecimaRevisao & " dia(s) depois.")
Unload UserForm_ConfigREV
UserForm_Config.Show

Sheets("CONFIGURAÇÃO").Protect Password:="ENDM10707045"
Application.ScreenUpdating = True

End Sub

Private Sub tb_revisao1_Change()
Me.tb_revisao1.MaxLength = 4
End Sub
Private Sub tb_revisao2_Change()
Me.tb_revisao2.MaxLength = 4
End Sub
Private Sub tb_revisao3_Change()
Me.tb_revisao3.MaxLength = 4
End Sub
Private Sub tb_revisao4_Change()
Me.tb_revisao4.MaxLength = 4
End Sub
Private Sub tb_revisao5_Change()
Me.tb_revisao5.MaxLength = 4
End Sub
Private Sub tb_revisao6_Change()
Me.tb_revisao6.MaxLength = 4
End Sub
Private Sub tb_revisao7_Change()
Me.tb_revisao7.MaxLength = 4
End Sub
Private Sub tb_revisao8_Change()
Me.tb_revisao8.MaxLength = 4
End Sub
Private Sub tb_revisao9_Change()
Me.tb_revisao9.MaxLength = 4
End Sub
Private Sub tb_revisao10_Change()
Me.tb_revisao10.MaxLength = 4
End Sub

Private Sub UserForm_Click()

End Sub
