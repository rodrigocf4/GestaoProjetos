Sub Statexec()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Worksheets("PROJETOS").Activate

Statorc

Dim faixa As Range

'Defino a faixa de data de entrada do projeto como referncia para o offset das colunas
'A entrada _ o fator primordial para dar o input do projeto, sendo o preenchimento dela necessörio para o status N o Iniciado

Set faixa = Range("F8", Range("F8").End(xlDown))

For Each Data In faixa

    If (Data.Offset(0, 9).Value = "CONCLUIDO") Or (Data.Offset(0, 9).Value = "CONCLUIDO COM ATRASO") Then

'Para a condicao nao Iniciado, a fase de orcamento deve estar concluida alimentar a data de inicio da execucao
'mesmo que se passe uma estimativa antes de iniciar
   
    If Data.Offset(0, 11).Value = "" Then
    
        Data.Offset(0, 15).Value = "NAO INICIADO"
    
    Else
        
        If (Data.Offset(0, 11).Value <> "") And ((Data.Offset(0, 12).Value >= Date) _
            Or (Data.Offset(0, 13).Value >= Date)) And (Data.Offset(0, 14).Value = "") Then
                
            Data.Offset(0, 15).Value = "EM ANDAMENTO"
           
        ElseIf (Data.Offset(0, 14).Value = "") And (Data.Offset(0, 11).Value <> "") And ((Data.Offset(0, 12).Value < Date) _
                 Or (Data.Offset(0, 13) < Date)) Then
            
            Data.Offset(0, 15).Value = "ATRASADO"
            
            Else
            
            If Data.Offset(0, 14) <> "" Then
            
                Select Case Data.Offset(0, 13).Value
            
                    Case ""
                        
                        Data.Offset(0, 15).Value = "CONCLUIDO"
            
                    Case Is <> ""
                        
                         Data.Offset(0, 15).Value = "CONCLUIDO COM ATRASO"
            
            End Select
            End If
                        
    End If
End If
Else

Data.Offset(0, 15).Value = ""
Data.Offset(0, 15).Interior.ColorIndex = xlNone

End If

Next Data

formatalin
atualizaplan

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub