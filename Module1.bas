

'OK Éliminer les números des compteurs qui se truvent plus d'un fois dans el rapport MVRS de S56 
Option Explicit
Dim LastRowMVRS As Long
Dim LastRowHEBDO As Long


'Remove double Meter Number
Sub RemoveDuplicateMeter()
    With Sheets("MVRS")
        LastRowMVRS = .Cells(.Rows.Count, "F").End(xlUp).Row
        Range("A1:V" & LastRowMVRS).RemoveDuplicates Columns:=6, Header:=xlYes
    End With
End Sub


' https://excelmacromastery.com/
Sub StringVLookup()
  Dim i As Integer
With Sheets("HEBDO")
    LastRowHEBDO = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
    Dim ValFinded As Variant
    Dim MeterNum As Variant
     
    For i = 2 to LastRowHEBDO
      MeterNum = Sheets("MVRS").Range("F" & i).Value
      ValFinded = Application.VLookup(MeterNum, Worksheets("HEBDO").Range("A1:FS" & LastRowHEBDO), 2, False)
      Sheets("MVRS").Range("M" & i)=ValFinded
    Next i

End Sub
'End OK


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
