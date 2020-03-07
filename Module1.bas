Option Explicit

Dim MeterColumn As Integer


MeterColumn=? 'Colonne du numéro de compteur = ?


'Éliminer les números des compteurs qui se truvent plus d'un fois dans el rapport MVRS de S 56 
Sub RemoveDuplicatesByMeter()
    Columns(MeterColumn).RemoveDuplicates Columns:=Array(1)
End Sub




Call RemoveDuplicatesByMeter()
