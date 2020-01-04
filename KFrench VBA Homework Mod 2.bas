Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
   'Sub StockValues():

'Define Variables
'Define worksheets since it has to loop through all 7 worksheets
'Use last row function
'Reset values to zero once the loop ends
Dim Ticker As String
Dim OpenP As Double
Dim CloseP As Double
Dim YlyChange As Double
Dim YDate As String

Dim PercentChange As Double
Dim SVolume As Double

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

Dim SumTabRow As Integer
SumTabRow = 2

YlyChange = 0

PercentChange = 0
OpenP = Cells(2, 3).Value

SVolume = 0

'70926 are the number of rows
'lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To 70926

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Ticker Name
Ticker = Cells(i, 1).Value
CloseP = Cells(i, 6).Value
      
'If YDate = Cells(i, 2).Value = Right(B1, 4) Then
'OpenP = Cells(i, 3).Value
      
' Add to the Total

YlyChange = CloseP - OpenP

      ' Print the Ticker in the Summary Table
      Range("I" & SumTabRow).Value = Ticker

      ' Print the YlyChange to the Summary Table
      Range("J" & SumTabRow).Value = YlyChange

PercentChange = ((CloseP - OpenP) / OpenP) * 100



'Print the PercentChange to the Summary Table
Range("K" & SumTabRow).Value = PercentChange



SVolume = SVolume + Cells(i, 7).Value


'Print the Stock Volume to the Summary Table
Range("L" & SumTabRow).Value = SVolume


SumTabRow = SumTabRow + 1
YlyChange = 0
PercentChange = 0

Else

      ' Add to the Brand Total
      SVolume = SVolume + Cells(i, 7).Value

End If



    If Cells(i, 10).Value > 0 Then Cells(i, 10).Interior.ColorIndex = 4
    If Cells(i, 10).Value < 0 Then Cells(i, 10).Interior.ColorIndex = 3
     If Cells(i, 11).Value > 0 Then Cells(i, 11).Interior.ColorIndex = 4
    If Cells(i, 11).Value < 0 Then Cells(i, 11).Interior.ColorIndex = 3
Next i



'End Sub
 


End Sub
