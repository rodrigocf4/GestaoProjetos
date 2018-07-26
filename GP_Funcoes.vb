Sub autofiltro()

Worksheets("PROJETOS").Activate

' Defino aqui a faixa que corresponde aos respons‡veis de projeto

Dim resp As Range
Dim x As Integer
Dim y As Range

If Worksheets("PROJETOS").AutoFilterMode Then
    
Worksheets("PROJETOS").UsedRange.AutoFilter

End If

x = Worksheets("PROJETOS").UsedRange.Rows.Count
Set y = Worksheets("TABELAS AUX ORC").Range("A9")
Set resp = Range("E8", "E" & x)
Worksheets("PROJETOS").UsedRange.Offset(6, 0).AutoFilter
resp.AutoFilter 5, y

End Sub

Sub atualizaplan()

Worksheets("TABELAS AUX ORC").Calculate
Worksheets("TABELAS AUX EXEC").Calculate
Worksheets("DASHBOARD").Calculate

End Sub
