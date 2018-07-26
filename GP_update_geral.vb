Sub Worksheet_Change(ByVal Target As Range)

Dim x As Integer

x = Worksheets("PROJETOS").UsedRange.Rows.Count

If Intersect(Target, Me.Range("H8", "M" & x)) Is Nothing Then GoTo canccong
       
  Application.EnableEvents = False
  Application.EnableEvents = True

Statexec

canccong:

If Intersect(Target, Me.Range("P8", "P" & x)) Is Nothing Then GoTo statusexec
       
  Application.EnableEvents = False
  Application.EnableEvents = True

Statexec


statusexec:

If Intersect(Target, Me.Range("Q8", "T" & x)) Is Nothing Then Exit Sub
       
  Application.EnableEvents = False
  Application.EnableEvents = True

Statexec

End Sub
