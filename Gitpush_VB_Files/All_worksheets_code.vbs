Sub Button1_Click()
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

Dim mycell As Range
Dim myrange1, myrange2 As Range
Dim ticker As String
Dim myrange1col As Integer

'set the range and find the corresponding value

Set myrange1 = Worksheets("A").Range("i2:j289")
GreatestTotalvolume = Application.WorksheetFunction.Max(myrange1)
Range("P4").Value = GreatestTotalvolume '
ticker = "i" & Application.WorksheetFunction.Match(Range("P4").Value, Range("j2:j289"), 0) + 1
Range("O4").Value = Range(ticker).Value


Set myrange2 = Worksheets("A").Range("l2:l289")
GreatestperIncrease = Application.WorksheetFunction.Max(myrange2)
Range("P2").Value = GreatestperIncrease
ticker = "i" & Application.WorksheetFunction.Match(Range("P2").Value, Range("l2:l289"), 0) + 1
Range("O2").Value = Range(ticker).Value


Set myrange3 = Worksheets("A").Range("l2:l289")
Greatestperdecrease = Application.WorksheetFunction.Min(myrange3)
Range("P3").Value = Greatestperdecrease
ticker = "i" & Application.WorksheetFunction.Match(Range("P3").Value, Range("l2:l289"), 0) + 1
Range("O3").Value = Range(ticker).Value




End Sub
Sub Hardsolution()
'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

'set the variables
Dim worksheetname As String
Dim ticker As String
Dim lastrow As Long
Dim column As Integer
Dim volume As Double
Dim Total_stock_volume As Long
Dim yearly_change As Double
Dim openticker As Double
Dim percentagechange As Double
Dim aopen As Integer
'yearly_change = 0



For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


worksheetsname = ws.Name
ws.Range("I1").Value = "Ticker"
ws.Range("j1").Value = "Totalvolume"
ws.Range("k1").Value = "Yearlychange"
ws.Range("l1").Value = "Percentagechange"
'MsgBox worksheetsname

volume = 0

column = 1
rownumber = 2
  'lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through rows and columns
For i = 2 To lastrow

ticker = ws.Cells(i, 1).Value

' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i, column).Value <> ws.Cells(i + 1, column).Value Then
    closeticker = ws.Cells(i, 6).Value
    'openticker = Cells(i + 1, 3).Value
  ' MsgBox (closeticker)
    'MsgBox ("firstopenticker" + Str(firstopenticker))
yearly_change = firstopenticker - closeticker
'MsgBox ("yearly change" + Str(firstopenticker) + Str(closeticker))
percentagechange = yearly_change / firstopenticker * 100


ws.Cells(rownumber, 9).Value = ticker
ws.Cells(rownumber, 10).Value = volume
ws.Cells(rownumber, 11).Value = yearly_change
ws.Cells(rownumber, 12).Value = percentagechange
If (percentagechange > 0) Then
        ws.Cells(rownumber, 11).Font.Color = 4
            ws.Cells(rownumber, 11).Interior.ColorIndex = 4
  Else
         ws.Cells(rownumber, 11).Font.Color = 3
          ws.Cells(rownumber, 11).Interior.ColorIndex = 3
       ' Range("k2:k298").interoir.Color = vbRed


    End If
   
   
    
    rownumber = rownumber + 1
    volume = 0
    aopen = 0
    
    Else
    aopen = aopen + 1
    
     volume = ws.Cells(i + 1, 7).Value + volume
     If aopen = 1 Then
     firstopenticker = ws.Cells(i, 3).Value
     
     End If
     
     End If
     
Next i

  

Next ws
End Sub
