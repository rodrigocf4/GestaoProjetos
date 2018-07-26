Sub formatalin()

Worksheets("PROJETOS").UsedRange.Borders.ColorIndex = 2

For Each lin In Worksheets("PROJETOS").UsedRange.Rows
        
    If (lin.Row Mod 2 = 0) And (lin.Row - 8 >= 0) Then
        
        lin.Interior.ColorIndex = 34
       
    ElseIf (lin.Row Mod 2) <> 0 And (lin.Row - 8 >= 0) Then
    
        lin.Interior.ColorIndex = 19
    
    End If

Next lin


urgencia
status
statusexec
cancelcongel

End Sub

Sub limpavazia()

Dim IndexRow As Integer
Dim x As Integer
x = Worksheets("PROJETOS").UsedRange.Rows.Count


For IndexRow = x To 8 Step -1
    
    If Cells(IndexRow, "F").Value = "" Then
    
        If Cells(IndexRow, "I").Value = "" Then
    
         Cells(IndexRow, "F").EntireRow.Delete
        
        End If
    End If
    
Next IndexRow

End Sub

Sub urgencia()

Dim rng As Range
Dim x As Integer
Dim y As Integer
x = Worksheets("PROJETOS").UsedRange.Rows.Count
y = 8

Set rng = Range("H8", "H" & x)

For Each lin In rng

If lin = "SIM" Then

Range("A" & y, "V" & y).Interior.ColorIndex = 45
lin.Interior.ColorIndex = 3

End If

y = y + 1

Next lin

End Sub

Sub cancelcongel()

Dim rngstatus As Range
Dim x As Integer
Dim y As Integer
x = Worksheets("PROJETOS").UsedRange.Rows.Count
y = 8

Set rngstatus = Range("P8", "P" & x)

For Each linha In rngstatus
 
    Select Case linha
    
    Case "CANCELADO"
   
    Range("A" & y, "V" & y).Interior.ColorIndex = 15
    
    Case "CONGELADO"
    
    Range("A" & y, "V" & y).Interior.ColorIndex = 17
    
    End Select
    
y = y + 1

Next linha

End Sub

Sub status()

Dim statlin As Range
Dim x As Integer
Dim y As Integer
x = Worksheets("PROJETOS").UsedRange.Rows.Count
y = 8

Set statlin = Range("O8", "O" & x)

For Each linha In statlin
 
    Select Case linha
    
    Case "CONCLUIDO"
   
     linha.Font.Bold = True
     linha.Font.ColorIndex = 1
     linha.Interior.ColorIndex = 10
    
    Case "CONCLUIDO COM ATRASO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 1
    linha.Interior.ColorIndex = 35
                        
    Case "ATRASADO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 2
    linha.Interior.ColorIndex = 3
    
    Case "EM ANDAMENTO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 1
    linha.Interior.ColorIndex = 6
    
    Case "NAO INICIADO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 2
    linha.Interior.ColorIndex = 21
    
    End Select
    
y = y + 1

Next linha

End Sub

Sub statusexec()

Dim statlinexec As Range
Dim x As Integer
Dim y As Integer
x = Worksheets("PROJETOS").UsedRange.Rows.Count
y = 8

Set statlinexec = Range("U8", "U" & x)

For Each linha In statlinexec
 
    Select Case linha
    
    Case "CONCLUIDO"
   
     linha.Font.Bold = True
     linha.Font.ColorIndex = 1
     linha.Interior.ColorIndex = 10
    
    Case "CONCLUIDO COM ATRASO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 1
    linha.Interior.ColorIndex = 35
                        
    Case "ATRASADO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 2
    linha.Interior.ColorIndex = 3
    
    Case "EM ANDAMENTO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 1
    linha.Interior.ColorIndex = 6
    
    Case "NAO INICIADO"
    
    linha.Font.Bold = True
    linha.Font.ColorIndex = 2
    linha.Interior.ColorIndex = 21
    
    End Select
    
y = y + 1

Next linha

End Sub
