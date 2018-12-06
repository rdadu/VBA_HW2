Sub StockPrices()
'Homework 2 - Pull from a series of worksheets in a workbook that loops thru each sheet by ticker and year
'for a given ticker: sum volume and determine the the low price and high price in a given year

  ' Set an initial variables
  Dim StockTicker As String
  Dim StockYear As Integer
  Dim StockVolume As Double
  Dim OpenPrice As Double
  Dim ClosePrice As Double
  Dim PriceChange As Double
  Dim ws As Worksheet
 
  ' Keep track each stock in the summary worksheet
  
  Dim SumSheet As Worksheet
  Set SumSheet = Worksheets("Summary")
  
  Dim SumRow As Integer
  SumRow = 2
  
' Need a Loop to go thru all sheets in workbook
  Dim i As Long
  Dim Lastrow As Long
  
  For Each ws In Worksheets
      Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      MsgBox (ws.Name)
    
      ' Loop through all Tickers minus header row.
      If ws.Name <> ("Summary") Then
          For i = 2 To Lastrow
        
           OpenPrice = IIf(i = 2, ws.Cells(i, 3).Value, OpenPrice)
            ' Check if we are still within the same ticker, if it is not then set values for summary..
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              ' Set the Stock name
              StockTicker = ws.Cells(i, 1).Value
              StockYear = Left(ws.Cells(i, 2).Value, 4)
            
              ' Add to the Volume Total
              StockVolume = StockVolume + ws.Cells(i, 7).Value
              ClosePrice = ws.Cells(i, 6).Value
        
              ' Print the Stock, Volume and Prices in the Summary Table >> move to a separate worksheet
              SumSheet.Cells(SumRow, 1).Value = StockTicker
              SumSheet.Cells(SumRow, 2).Value = StockYear
              SumSheet.Cells(SumRow, 3).Value = OpenPrice
              SumSheet.Cells(SumRow, 4).Value = ClosePrice
              If OpenPrice <> 0 Then
                SumSheet.Cells(SumRow, 5).Value = Abs((OpenPrice - ClosePrice) / OpenPrice)
              Else
                SumSheet.Cells(SumRow, 5).Value = 100
              End If
              
              If OpenPrice < ClosePrice Then
                SumSheet.Cells(SumRow, 5).Interior.ColorIndex = 4
              Else
                SumSheet.Cells(SumRow, 5).Interior.ColorIndex = 3
              End If
              
              SumSheet.Cells(SumRow, 6).Value = StockVolume
        
              ' Add one to the summary table row
              SumRow = SumRow + 1
              StockVolume = 0
              OpenPrice = ws.Cells(i + 1, 3).Value
           Else
            ' If the cell immediately following a row is the same ticker...
            ' Add to the Total
              StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            End If
          Next i
    End If
 Next ws

End Sub
