Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    Dim xStr As String
    
    xStr = "PROJETOS"
    If Not Intersect(Target, Range("B15")) Is Nothing Then
        Sheets(xStr).Activate
        
Worksheets("PROJETOS").Activate

' Defino aqui a faixa que corresponde aos respons‡veis de projeto

Dim resp1 As Range
Dim urg As Range
Dim a As Integer
Dim b As Range
Dim c As Variant

c = "SIM"


a = Worksheets("PROJETOS").UsedRange.Rows.Count
Set b = Worksheets("TABELAS AUX ORC").Range("A9")
Set urg = Worksheets("PROJETOS").Range("H8", "H" & a)
Set resp1 = Worksheets("PROJETOS").Range("E8", "E" & a)

resp1.AutoFilter 5, b
urg.AutoFilter 8, c

    End If

End Sub
