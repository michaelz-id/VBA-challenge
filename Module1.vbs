Attribute VB_Name = "Module1"
Sub Ticker_sort()

'declare variables

Dim Ticker As String
Dim i As Double
Dim LastRow As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Yrchange As Double
Dim SumTab As Integer
Dim TotalStockVol As Double
Dim ws As Worksheet
Dim LRowValue As Long
Dim TRowValue As Long

'Turn off screen updating to save time and stop the flicker
Application.ScreenUpdating = False

'Loop to go through all worksheets
For Each ws In Worksheets

    'MsgBox ws.Name

    'Set value for LastRow to get extent of column
    LastRow = ws.Cells(Cells.Rows.Count, 1).End(xlUp).Row
           
    'Set counter - TotalStockVol to inital zero value
    TotalStockVol = 0
    
    'set summary table
    SumTab = 1
    
    'Set up Summary Table Headings and formatting Headings / columns
    
    'Summary Headings
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 9).EntireColumn.AutoFit
    
    'Yearly Change heading
     ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 10).EntireColumn.AutoFit
    
    'Percent Change heading
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 11).EntireColumn.AutoFit
    
    'Total Stock Volume heading
     ws.Cells(1, 12).Value = "Total Stock Volume"
     ws.Cells(1, 12).EntireColumn.AutoFit
    
    'Bold Font
     ws.Range("I1:L1").Font.Bold = True
    
    'Format Columns for for apropriate presentation of information
     ws.Range("J2:J" & LastRow).NumberFormat = "#,##0.00"
    
    'Loop through rows
    'Open price for beginning of the year and closing price for the end.
    'Calculate Total Change and Percentage change and total volume per ticker
    
    For i = 2 To LastRow
 
    
        'Test Cell against the value of Cell below to see if value changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
               'Record Ticker Name
                Ticker = ws.Cells(i, 1).Value
                                
                'Find OpenPrice for each value
                TRowValue = ws.Range("A:A").Find(What:=Ticker, SearchDirection:=xlNext, LookAt:=xlWhole).Row
                'MsgBox TRowValue
                
                'OpenPrice
                OpenPrice = ws.Cells(TRowValue, 3).Value
                'MsgBox OpenPrice
                
                'Find ClosePrice
                LRowValue = ws.Range("A:A").Find(What:=Ticker, SearchDirection:=xlPrevious, LookAt:=xlWhole).Row
                'MsgBox LRowValue
                
                'ClosePrice
                ClosePrice = ws.Cells(LRowValue, 6).Value
                'MsgBox ClosePrice
                                
                'Add total Stock Volume
                TotalStockVol = ws.Cells(i, 7).Value + TotalStockVol
                               
                'Put Ticker name in summary Table
                 ws.Cells(SumTab + 1, 9) = Ticker
                
                'Calculate yearly change
                 Yrchange = (ClosePrice - OpenPrice)

                'MsgBox Yrchange
                
                'Put Yearly Change in summary table
                 ws.Cells(SumTab + 1, 10) = Yrchange
                
                'If YrChange >= 0 then Green for positive else red for negative
                If Yrchange >= 0 Then
                
                    ws.Cells(SumTab + 1, 10).Interior.ColorIndex = 4
                    
                Else
                
                     ws.Cells(SumTab + 1, 10).Interior.ColorIndex = 3
                
                End If
                    
                'Calculate Percentage annual change (yearly change/opening value)*100
                 ws.Cells(SumTab + 1, 11) = FormatPercent(Yrchange / OpenPrice, 2)
                
                'Put total stock volume in summary table
                 ws.Cells(SumTab + 1, 12) = TotalStockVol
                            
                'Summary Table row moved down 1 to record next entry
                SumTab = SumTab + 1
                
                'Reset counts for opening and closing prices
                TotalStockVol = 0
                
            Else
                'count stock volume
                TotalStockVol = ws.Cells(i, 7).Value + TotalStockVol
            
            End If
     
     
    Next i
    
Next ws

'call upon BonusPart to calculate bonus part of challenge.
BonusPart

'Turn Screenupdating back on
Application.ScreenUpdating = True

MsgBox "All done"

End Sub

Sub BonusPart()

'declare variables

Dim i As Integer
Dim LastRow As Integer
Dim MaxStockValue As Double
Dim MinPer As Double
Dim MaxPer As Double
Dim ws As Worksheet
Dim MaxTicker As Integer
Dim MinTicker As Integer
Dim SVTick As Integer


'Turn off screen updating to save time and stop the flicker
Application.ScreenUpdating = False

For Each ws In Worksheets


    'Set value for LastRow to search for full extent of the column
    LastRow = ws.Cells(Cells.Rows.Count, 9).End(xlUp).Row

    'Set up Row headings Increase
    ws.Cells(2, 15).Value = "Greatest Percentage Increase"
    ws.Cells(2, 15).EntireColumn.AutoFit
    
    'Set up Row headings decrease
    ws.Cells(3, 15).Value = "Greatest Percentage Decrease"
    ws.Cells(3, 15).EntireColumn.AutoFit

    'Set up Row headings total volume
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 15).EntireColumn.AutoFit

    'Set titles for column headers
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 16).EntireColumn.AutoFit
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 17).EntireColumn.AutoFit

    'Format text for table
    ws.Range("O2:O5").Font.Bold = True
    ws.Range("P1:Q1").Font.Bold = True
 
    'get max percentage value
    MaxPer = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    
    'MsgBox MaxPer
    
    'show max percentage
    ws.Cells(2, 17) = FormatPercent(MaxPer)
    
    'get min percentage
    MinPer = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    
    'MsgBox MinPer
    
    'show min percentage
    ws.Cells(3, 17) = FormatPercent(MinPer)
    
    'get max
    MaxStockValue = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    
    'MsgBox MaxStockValue
    
    'return Max stock volume value
    ws.Cells(4, 17) = MaxStockValue
    
    'Ensure that columns are wide enough to display information
    ws.Cells(1, 16).EntireColumn.AutoFit
    ws.Cells(1, 17).EntireColumn.AutoFit
    
    'return max percentage value formatted as a percentage
    ws.Cells(2, 17) = FormatPercent(MaxPer)
    
    
    'retreive ticker for max percentage
    MaxTicker = Application.Match(ws.Range("Q2"), ws.Range("k2:k" & LastRow), 0) + 1
    
    'MsgBox MaxTicker
    ws.Cells(2, 16).Value = ws.Cells(MaxTicker, 9)
    
    'retreive ticker for minimum percentage
    MinTicker = Application.Match(ws.Range("Q3"), ws.Range("k2:k" & LastRow), 0) + 1
    
    'MsgBox MinTicker
    ws.Cells(3, 16).Value = ws.Cells(MinTicker, 9)
    
    'retrieve ticker for max stock value
    SVTick = Application.Match(ws.Range("Q4"), ws.Range("L2:L" & LastRow), 0) + 1
    
    'MsgBox Max Ticker Volumne
    ws.Cells(4, 16).Value = ws.Cells(SVTick, 9)

Next ws

'Turn Screenupdating back on
Application.ScreenUpdating = True

End Sub


