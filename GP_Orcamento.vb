Sub Statorc()

Dim faixa As Range

'Defino a faixa de data de entrada do projeto como referncia para o offset das colunass
'A entrada _ o fator primordial para dar o input do projeto, sendo o preenchimento dela necessörio para o status N o Iniciado

Set faixa = Range("F8", Range("F8").End(xlDown))

For Each Data In faixa

    If Data.Offset(0, 3).Value <> "" Then

'Para a condi o N o Iniciado, somente a data da entrada do projeto deve ser preenchida, sem preenchimento da data de inêcioo
'mesmo que se passe uma estimativa antes de iniciar
   
    If Data.Offset(0, 4).Value = "" Then
    
        Data.Offset(0, 9).Value = "NAO INICIADO"
    
    Else
        
        If (Data.Offset(0, 3).Value <> "") And (Data.Offset(0, 4).Value <> "") And ((Data.Offset(0, 5).Value >= Date) _
            Or (Data.Offset(0, 6).Value >= Date)) And (Data.Offset(0, 7).Value = "") Then
                
            Data.Offset(0, 9).Value = "EM ANDAMENTO"
           
        ElseIf (Data.Offset(0, 7).Value = "") And (Data.Offset(0, 4).Value <> "") And ((Data.Offset(0, 5).Value < Date) _
                 And (Data.Offset(0, 6) < Date) Or (Data.Offset(0, 6) = "")) Then
            
            Data.Offset(0, 9).Value = "ATRASADO"
            
            Else
            
            If Data.Offset(0, 7) <> "" Then
            
                Select Case Data.Offset(0, 6).Value
            
                    Case ""
                        
                        Data.Offset(0, 9).Value = "CONCLUIDO"
            
                    Case Is <> ""
                        
                         Data.Offset(0, 9).Value = "CONCLUIDO COM ATRASO"
            
            End Select
            End If
                        
    End If
End If
Else

Data.Offset(0, 9).Value = ""
Data.Offset(0, 9).Interior.ColorIndex = xlNone

End If

Next Data

End Sub