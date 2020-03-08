Option Explicit

Dim MeterColumn As Integer


MeterColumn=? 'Colonne du numéro de compteur = ?


'Éliminer les números des compteurs qui se truvent plus d'un fois dans el rapport MVRS de S 56 
Sub RemoveDuplicatesByMeter()
    Columns(MeterColumn).RemoveDuplicates Columns:=Array(1)
End Sub


Sub Filter()

Dim LastRow As Long

With Sheets("MVRS")
        LastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
       .AutoFilterMode = False
        With .Range("A2:Z" & LastRow)
             .AutoFilter Field:=1, Criteria1:=Array("April", "August"), Operator:=xlFilterValues
             .AutoFilter Field:=2, Criteria1:="<>"
             ActiveSheet.AutoFilter.Range.Copy
             Sheets("Chart").Select
             Range("A7").Select
            Sheets("Chart").Paste
         End With
End With
End Sub


Sub DepBddS56()
Call RemoveDuplicatesByMeter()
Call Filter()
End Sub
