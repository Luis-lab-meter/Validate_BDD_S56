Option Explicit

Dim MeterColumn As Integer


MeterColumn=? 'Colonne du numéro de compteur = ?


'OK Éliminer les números des compteurs qui se truvent plus d'un fois dans el rapport MVRS de S56 
Sub RemoveDuplicateMeter()
    With Sheets("MVRS")
        Dim LastRow As Long
        LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        Range("A1:C" & LastRow).Select
        ActiveSheet.Range("A1:C" & LastRow).RemoveDuplicates Columns:=2, Header:=xlYes
    End With
End Sub
'End OK

' https://excelmacromastery.com/
Sub StringVLookup()
    
    Dim sFruit As String
    sFruit = "Plum"
    
    Dim sRes As Variant
    sRes = Application.VLookup( _
                       sFruit, shData.Range("A2:B7"), 2, False)
    
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

'Vlookup, find las row and extend range
Sub MakeFormulas()
Dim SourceLastRow As Long
Dim OutputLastRow As Long
Dim sourceSheet As Worksheet
Dim outputSheet As Worksheet

'What are the names of our worksheets?
Set sourceSheet = Worksheets("Sheet1")
Set outputSheet = Worksheets("Sheet2")

'Determine last row of source
With sourceSheet
    SourceLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
With outputSheet
    'Determine last row in col P
    OutputLastRow = .Cells(.Rows.Count, "P").End(xlUp).Row
    'Apply our formula
    .Range("Q2:Q" & OutputLastRow).Formula = _
        "=VLOOKUP(A2,'" & sourceSheet.Name & "'!$A$2:$B$" & SourceLastRow & ",2,0)"
End With
End Sub



Sub DepBddS56()
Call RemoveDuplicatesByMeter()
Call Filter()
End Sub
