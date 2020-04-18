VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockmarket()

Dim ticker As String
Dim numberTicker As Integer
Dim lastrow As Long
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As Double

Dim greatestPercentIncrease As Double
Dim greastpercentIncreaseTicker As String
Dim greatestPercentdecrease As Double
Dim greatestpercentdecreaseTicker As String
Dim greatestStockVol As Double
Dim greatestStockVolTicker As String

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

numberTicker = 0
ticker = ""
yearlyChange = 0
openPrice = 0
percentChange = 0
totalVolume = 0

For i = 2 To lastrow

    ticker = ws.Cells(i, 1).Value
    
    If openPrice = 0 Then
        openPrice = ws.Cells(i, 3).Value
    End If
    
    totalVolume = totalVolume + ws.Cells(i, 7).Value
    
    If ws.Cells(i + 1, 1).Value <> ticker Then
        numberTicker = numberTicker + 1
        ws.Cells(numberTicker + 1, 9) = ticker
        
        closePrice = ws.Cells(i, 6).Value
        
        yearChange = closePrice - openPrice
        
        ws.Cells(numberTicker + 1, 10).Value = yearChange
        
        If yearChange > 0 Then
            ws.Cells(numberTicker + 1, 10).Interior.ColorIndex = 4
        ElseIf yearChange < 0 Then
            ws.Cells(numberTicker + 1, 10).Interior.ColorIndex = 3
        End If
        
        If openPrice = 0 Then
            percentChange = 0
        Else
            percentChange = yearChange / openPrice
        End If
        
        
        ws.Cells(numberTicker + 1, 11).Value = Format(percentChange, "Percent")
        
        openPrice = 0
        ws.Cells(numberTicker + 1, 12).Value = totalVolume
        totalVolume = 0
    End If
        
    Next i
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    greatestPercentIncrease = ws.Cells(2, 11).Value
    greatestpercentIncreaseTicker = ws.Cells(2, 9).Value
    greatestPercentdecrease = ws.Cells(2, 11).Value
    greatestpercentdecreaseTicker = ws.Cells(2, 9).Value
    greatestStockVol = ws.Cells(2, 12).Value
    greatestStockVolTicker = ws.Cells(2, 9).Value
    
    For i = 2 To lastrow
    
        If ws.Cells(i, 11).Value > greatestPercentIncrease Then
            greatestPercentIncrease = ws.Cells(i, 11).Value
            greatestpercentIncreaseTicker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 11).Value < greatestPercentdecrease Then
            greatestPercentdecrease = ws.Cells(i, 11).Value
            greatestpercentdecreaseTicker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value > greatestStockVol Then
            greatestStockVol = ws.Cells(i, 12).Value
            greatestStockVolTicker = ws.Cells(i, 9).Value
        End If
        
      Next i
      
    ws.Range("P2").Value = Format(greatestpercentIncreaseTicker, "Percent")
    ws.Range("Q2").Value = Format(greatestPercentIncrease, "Percent")
    ws.Range("P3").Value = Format(greatestpercentdecreaseTicker, "Percent")
    ws.Range("Q3").Value = Format(greatestPercentdecrease, "Percent")
    ws.Range("P4").Value = greatestStockVolTicker
    ws.Range("Q4").Value = greatestStockVol
            
    
 
    Next ws
    


End Sub


