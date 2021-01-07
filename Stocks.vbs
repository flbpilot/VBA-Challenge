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
    
    'dim worksheet and workbook 
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    'set columns from (i-l)
    For Each ws In wb.Sheets
        ws.Cells(1, 9).Value = ("Ticker Symbol")
        ws.Cells(1, 10).Value = ("Yearly Change")
        ws.Cells(1, 11).Value = ("Percent Change")
        ws.Cells(1, 12).Value = ("Total Volume")
        ws.Cells(1, 16).Value = ("Ticker")
        ws.Cells(1, 17).Value = ("Value")
        ws.Cells(2, 15).Value = ("Greatest % increase")
        ws.Cells(3, 15).Value = ("Greatest % decrease")
        ws.Cells(4, 15).Value = ("Greatest total volume")
        

        
    'set variables for calculations
    Ticker = 0
    Volume = 0
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    MaxTicker = 0
    MinTicker = 0
    MaxPercent = 0
    MinPercent = 0
    MaxVolTicker = 0
    MaxVolume = 0
    
    'set location for variables
    SummaryTable = 2 
    
    'Loop through cells to find last row
   LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
  
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
                
            
                'add the total volume to the ticker name
                Volume = Volume + ws.Cells(i, 7).Value
                
                'put the ticker name in the table
                ws.Range("I" & SummaryTable).Value = Ticker
                
                'put the yearly change in the table
                ws.Range("J" & SummaryTable).Value = YearlyChange
    
                
                'put the percent change in the table
                ws.Range("K" & SummaryTable).Value = (CStr(PercentChange) & "%")

                'put the total volume in the table
                ws.Range("L" & SummaryTable).Value = TotalVolume
                
                'add 1 to the summary table row count
                SummaryTable = SummaryTable + 1
                
              'get next beginning price
              OpenPrice = ws.Cells(i + 1, 3).Value
                
              

              'reset values
              PercentChange = 0
              TotalVolume = 0
              
         
          Else
          
              TotalVolume = TotalVolume + ws.Cells(i, 7).Value
              
          End If
            
        Next i
                
                
        Next ws
    
End Sub
