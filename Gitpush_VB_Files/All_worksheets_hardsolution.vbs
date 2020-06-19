'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 


Sub Greatestpercentvalues()

Dim mycell As Range
Dim myrange1, myrange2 As Range
Dim ticker As String
Dim myrange1col As Integer
Dim lastrow As Long
Dim Greatest_Total_volume As Long

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


worksheetsname = ws.Name

'MsgBox worksheetsname
ws.Range("N4").Value = "Greatest ToTal volume"
ws.Range("N3").Value = "Greatest%decrease"
ws.Range("N2").Value = "Greatest%increase"
'set the range and find the corresponding value

Set myrange1 = ws.Range("i2:j289")
GreatestTotalvolume = ws.Application.WorksheetFunction.Max(myrange1)
ws.Range("P4").Value = GreatestTotalvolume
'GreatestTotalvolume = ws.Range("N4").Value
'MsgBox (GreatestTotalvolume)
ticker = "i" & WorksheetFunction.Match(ws.Range("P4").Value, ws.Range("j2:j289"), 0) + 1
ws.Range("O4").Value = ws.Range(ticker).Value


Set myrange2 = ws.Range("l2:l289")
GreatestperIncrease = ws.Application.WorksheetFunction.Max(myrange2)
ws.Range("P2").Value = GreatestperIncrease
ticker = "i" & Application.WorksheetFunction.Match(ws.Range("P2").Value, ws.Range("l2:l289"), 0) + 1
ws.Range("O2").Value = ws.Range(ticker).Value


Set myrange3 = ws.Range("l2:l289")
Greatestperdecrease = ws.Application.WorksheetFunction.Min(myrange3)
ws.Range("P3").Value = Greatestperdecrease
ticker = "i" & Application.WorksheetFunction.Match(ws.Range("P3").Value, ws.Range("l2:l289"), 0) + 1
ws.Range("O3").Value = ws.Range(ticker).Value

Next ws


End Sub
