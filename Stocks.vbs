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
    
    'set columns from (i-l)
        For Each ws In wb.Sheets
        ws.Cells(1, 9).Value = "Ticker_Symbol"
        ws.Cells(1, 10).Value = "YearlyChange"
        ws.Cells(1, 11).Value = "PercentChange"
        ws.Cells(1, 12).Value = "TotalVolume"
        
     
    'Loop through cells to find last row'
   LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
