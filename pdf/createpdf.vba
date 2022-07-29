Sub MACROPDF()

Application.ScreenUpdating = False
Sheets("RELATORIO PDF").Unprotect Password:="ENDM10707045"
Sheets("CONFIGURAÇÃO").Unprotect Password:="ENDM10707045"

Sheets("CONFIGURAÇÃO").Range("C53").Value = Time$
SemanaPDF = Sheets("CONFIGURAÇÃO").Range("AR14").Value
DiaPDF = Sheets("CONFIGURAÇÃO").Range("AR15").Value

fase = "Semana " & SemanaPDF & " Dia " & DiaPDF
localnome = ThisWorkbook.Path & "/" & fase
    Sheets("RELATORIO PDF").Range("A1:W49").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        localnome, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
                             
Sheets("CONFIGURAÇÃO").Protect Password:="ENDM10707045"
Sheets("RELATORIO PDF").Protect Password:="ENDM10707045"
Sheets("ESTUDOS").Select
Application.ScreenUpdating = True
   
        
End Sub
