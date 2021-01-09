Sub stocks1()

    'set Ticker variable
    Dim Ticker As String
    
    'set Open Price Variable
    Dim OpenPrice As Double
    
    'set Close Price variable
    Dim ClosePrice As Double
    
    'set Maximum Ticker variable
    Dim MaxTicker As String
    
    'set Minimum Ticker variable
    Dim MinTicker As String
    
    'set Maximum Percentage variable
    Dim MaxPercent As Double
    
    'set Minimum Percentage variable
    Dim MinPercent As Double
    
    'set Volume variable
    Dim Volume As Double
    
    'set Annual Change variable
    Dim YearlyChange As Double
    
    'set Annual Percentage Change variable
    Dim PercentChange As Double
    
    'set Total Volume variable
    Dim TotalVolume As Double
    
    'set Maximum Volume Ticker Name variable
    Dim MaxVolTicker As String
    
    'set Maximum Ticker value variable
    Dim MaxVolume As Double
    
    'set Summary Table variable
    Dim SummaryTable As Long
    
    'set Last row variable
    Dim Lastrow As Long
    
    'dim worksheet and workbook with respect to their lox
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    'set loop on all worksheets
        For Each ws In wb.Sheets
    
    'set headers rows from (i-l)
        ws.Cells(1, 9).Value = ("Ticker_Symbol")
        ws.Cells(1, 10).Value = ("Yearly_Change")
        ws.Cells(1, 11).Value = ("Percent_Change")
        ws.Cells(1, 12).Value = ("Total_Volume")
        ws.Cells(1, 16).Value = ("Ticker")
        ws.Cells(1, 17).Value = ("Value")
        ws.Cells(2, 15).Value = ("Greatest % increase")
        ws.Cells(3, 15).Value = ("Greatest % decrease")
        ws.Cells(4, 15).Value = ("Greatest total volume")
        
    'Assignment 
    Ticker = " "
    Volume = 0
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    MaxTicker = " "
    MinTicker = " "
    MaxPercent = 0
    MinPercent = 0
    MaxVolTicker =  " "
    MaxVolume = 0
    SummaryTable = 2
    
    'Loop through cells to last row
   LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  
        'set the begining value of stock
        OpenPrice = ws.Cells(2, 3).Value
    
        'loop through the ticker symbol
        For i = 2 To Lastrow


    'set the ticker name
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

               
                'initiate the ticker startup point
                Ticker = ws.Cells(i, 1).Value
                
                'calculate the year end price and price change for year
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice

                'set condition for a 0 opening value
                If OpenPrice <> 0 Then
                PercentChange = (YearlyChange / OpenPrice) * 100
            
                End If
                    
            
                'add the total volume to the ticker name
                Volume = Volume + ws.Cells(i, 7).Value
                
                'print ticker name in the table
                ws.Range("I" & SummaryTable).Value = Ticker
                
                'print yearly change in the table
                ws.Range("J" & SummaryTable).Value = YearlyChange

                'color red for negative, green for positive yearly change

                If (YearlyChange <= 0) Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                ElseIf (YearlyChange > 0) Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                End If
                
                'put the percent change in the table
                ws.Range("K" & SummaryTable).Value = (CStr(PercentChange) & "%")

                'put the total volume in the table
                ws.Range("L" & SummaryTable).Value = TotalVolume
                
                'add 1 to the summary table row count
                SummaryTable = SummaryTable + 1
                
              'get next beginning price
              OpenPrice = ws.Cells(i + 1, 3).Value
              
              'do calculations
              
              If (PercentChange > MaxPercent) Then
                    MaxPercent = PercentChange
                    MaxTicker = Ticker

                ElseIf (PercentChange < MinPercent) Then
                    MinPercent = PercentChange
                    MinTicker = Ticker
                    
                End If

                If (TotalVolume > MaxVolume) Then
                    MaxVolume = TotalVolume
                    MaxVolTicker = Ticker

                End If                  
              

              'reset values
              PercentChange = 0
              TotalVolume = 0
              
         
          Else
          
              TotalVolume = TotalVolume + ws.Cells(i, 7).Value
              
          End If
            
        Next i
                'print values in assigned cells
                ws.Range("Q2").Value = (CStr(MaxPercent) & "%")
                ws.Range("Q3").Value = (CStr(MinPercent) & "%")
                ws.Range("P2").Value = MaxTicker
                ws.Range("P3").Value = MinTicker
                ws.Range("Q4").Value = MaxVolume
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"    
               'Auto Resize Columns 
                ws.Columns("A:Q").AutoFit
            
                
        Next ws
    
End Sub
