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