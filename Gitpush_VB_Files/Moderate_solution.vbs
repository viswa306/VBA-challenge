Sub newsheet()
'code for moderate solution
'Create a script that will loop through all the stocks for one year and output the following information.

  ' The ticker symbol.

 ' * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

 ' * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

 ' * The total stock volume of the stock.

'* You should also have conditional formatting that will highlight positive change in green and negative change in red.

'* The result should look as follows.

'set the variables
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
volume = 0

column = 1
rownumber = 2
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through rows and columns
For i = 2 To lastrow

ticker = Cells(i, 1).Value

' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i, column).Value <> Cells(i + 1, column).Value Then
                closeticker = Cells(i, 6).Value
 
                yearly_change = firstopenticker - closeticker
'MsgBox ("yearly change" + Str(firstopenticker) + Str(closeticker))
              percentagechange = (yearly_change / firstopenticker * 100)
              'percentagechange = Round(percentagechange, 2)
 'MsgBox (percentagechange)

            Cells(rownumber, 9).Value = ticker
            Cells(rownumber, 10).Value = volume
            Cells(rownumber, 11).Value = yearly_change
            
          Cells(rownumber, 12).Value = percentagechange
          'rownumber.Value = Round(Cells(rownumber.Value, 12).Value, 2)
           
             
            
 If (percentagechange > 0) Then
        Cells(rownumber, 11).Font.Color = 4
            Cells(rownumber, 11).Interior.ColorIndex = 4
  Else
          Cells(rownumber, 11).Font.Color = 3
            Cells(rownumber, 11).Interior.ColorIndex = 3
       ' Range("k2:k298").interoir.Color = vbRed


    End If
    rownumber = rownumber + 1
    volume = 0
    aopen = 0
    
    Else
        aopen = aopen + 1
    
        volume = Cells(i + 1, 7).Value + volume
        
        
 If aopen = 1 Then
     firstopenticker = Cells(i, 3).Value
     
     End If
     
     End If
     
Next i

End Sub
