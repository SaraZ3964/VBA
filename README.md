# VBA-Homework
Sub Stock()

Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate  

Dim Ticker As String
Dim Total As Double
Total = 0

Dim YearlyChange As Double
YearlyChange = 0
Dim PercentChange As Double
Dim OP As Double
Dim CP As Double

Dim j As Integer
j = 2

last = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total stock volumn"
ws.Range("L1").Value = "Percent Change"
ws.Range("K1").Value = "Yearly Change"

For i = 2 To last
   If Cells(i, 1).Value <> Cells(i + 1, 1) Then
        Ticker = Cells(i, 1).Value
        Total = Cells(i, 7).Value + Total
    
    OP = Cells(i, 3).Value
    CP = Cells(i, 6).Value
    
    YearlyChange = CP - OP
    
    If (OP <> 0 And CP <> 0) Then
      PercentChange = YearlyChange / OP
       ws.Range("L" & j).Value = PercentChange
    ws.Range("L" & j).NumberFormat = "0.00%"
    ElseIf (OP = 0 And CP <> 0) Then
            PercentChange = 1
    Else:
        PercentChange = 0
    End If
    
    ws.Range("K" & j).Value = YearlyChange
    ws.Range("I" & j).Value = Ticker
    ws.Range("J" & j).Value = Total
    j = j + 1
    Total = 0
    OP = Cells(i + 1, 3).Value
    
    Else:
        Total = Cells(i, 7).Value + Total
    
End If

   If ws.Range("K" & j).Value > 0 Then
        ws.Range("K" & j).Interior.ColorIndex = 4
    
    Else:
        ws.Range("K" & j).Interior.ColorIndex = 3
End If

ws.Range("N2").Value = "Greatest%increase"
ws.Range("N3").Value = "Greatest%Decrease"
ws.Range("N4").Value = "Greatest total volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"


Dim Maxpercent As Double
Dim Minpercent As Double
Dim Maxtotal As Double


Maxpercent = Application.WorksheetFunction.Max(ws.Range("l:l"))
Cells(2, 16).Value = Maxpercent
Cells(2, 16).NumberFormat = "0.00%"

Minpercent = Application.WorksheetFunction.Min(ws.Range("l:l"))
Cells(3, 16).Value = Minpercent
Cells(3, 16).NumberFormat = "0.00%"

Maxtotal = Application.WorksheetFunction.Max(ws.Range("j:j"))
Cells(4, 16).Value = Maxtotal


If Cells(i, 12).Value = Maxpercent Then
    Cells(2, 15).Value = Cells(i, 9).Value

ElseIf Cells(i, 12).Value = Minpercent Then
               Cells(3, 15).Value = Cells(i, 9).Value

ElseIf Cells(i, 10).Value = Maxtotal Then
               Cells(3, 15).Value = Cells(i, 9).Value
End If

Next i
Next ws

End Sub
