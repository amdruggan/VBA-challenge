Attribute VB_Name = "Module2"

Sub GenerateOutputTablesForMultipleSheets():
    'This subroutine's sole purpose is to call the other subroutines so that I don't need to loop through the entirety of the other subroutines inside of them.
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        MakeTickerSummationTable ws.Name
        MsgBox ("Completed: " + ws.Name)
        
    Next ws

End Sub

Sub MakeTickerSummationTable(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Range, openColumn As Range, closeColumn As Range, volumeColumn As Range
    Dim outputTable As Range
    Dim Ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim i As Long
    
    
    
    'Set the worksheet to work with
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    'Set the range for the columns in the table we are inputting with
    With ws
        'Finding the bottom of the columns
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        'Storing the column for the ticker in tickerColumn as a range, then doing the same for open, close, and volume
        Set tickerColumn = .Range("A2:A" & lastRow)
        Set openColumn = .Range("C2:C" & lastRow)
        Set closeColumn = .Range("F2:F" & lastRow)
        Set volumeColumn = .Range("G2:G" & lastRow)
    End With
    
    'Outputting to a table
    'This could be 4 separate Range("__").value statements, but doing it with an Array was something cool i picked up on (:
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percentage Change", "Total Volume")
    
    'Needed to set outputTable area as part of the worksheet that I then build from
    Set outputTable = ws.Range("I2")
    
    'Loop through the data to calculate yearly changes and other values
    'Starting at i = 2 and then going until the previously designated last row of the table.
    For i = 2 To lastRow
        'This makes it so that if the sheet is bigger than the number of rows with tickers we don't loop through blank rows:
        If IsEmpty(tickerColumn(i)) Then Exit For
        
        'Check to see if this row of the loop is for the current ticker or if we have encountered a new ticker
        If tickerColumn(i) <> Ticker Then
            'When we find that new tickerwe update the opening price and the ticker name, using the previously defined column variables.
            Ticker = tickerColumn(i)
            openingPrice = openColumn(i)
        End If
        
        'I couldn't figure out why this wasn't working for the first ticker, so I just hard coded a solution to that issue here.
        If i = 2 Then
            Ticker = tickerColumn(2)
            openingPrice = openColumn(2)
        End If
        
        
        'Always update closing price and total volume
        'Closing price could should technically be updated right before you make the yearlyChange or percentageChange calculations, but that wasn't functioning correctly and this was.
        closingPrice = closeColumn(i)
        totalVolume = totalVolume + volumeColumn(i)
        
        'First we check to see if the current entry is the last one or if the next ticker is different
        'In either case we are going to need to update the yearly change since we now have an Open and close price
        If i = lastRow Or tickerColumn(i + 1) <> Ticker Then
        
            'Calculate yearly change and percentage change
            'Technically could be a case where openingPrice = 0 so we need to create an exception here, but I didn't find that in the data.
            yearlyChange = closingPrice - openingPrice
            percentageChange = yearlyChange / openingPrice
            
            'Output the results in the output table
            'I learned to do this with a range dimension instead of Cells()/Value, as it makes it easier to iterate through and the Offset(r,c) method makes this easy
            outputTable.Value = Ticker
            outputTable.Offset(0, 1).Value = yearlyChange
            outputTable.Offset(0, 2).Value = Format(percentageChange, "0.00%")
            outputTable.Offset(0, 3).Value = totalVolume
            
            'Apply formatting to yearly change column, depending on if it if is negative or positive
            If yearlyChange > 0 Then
                outputTable.Offset(0, 1).Interior.ColorIndex = 4
                'Change color to green (color index 4). This occurs when % change over the year is > 0.
            ElseIf yearlyChange < 0 Then
                outputTable.Offset(0, 1).Interior.ColorIndex = 3
                'Change color to red (color index 3). This occurs when % change over the year is < 0.
            End If
            
            'Move to the next row in the output table, offsetting downwards by 1 row
            Set outputTable = outputTable.Offset(1, 0)
            
            'Since at this point we have placed our total volume in the table and are looking at a new ticker, we reset the totalVolume dim to start counting for the next ticker
            totalVolume = 0
        End If
    Next i
    
    'Create the min and max percent table and highest vol table for this sheet
    MinMaxPercentMaxVol ws.Name
    
End Sub

'method takes argument of current sheet
Sub MinMaxPercentMaxVol(ByVal sheetName As String)
    Dim ws As Worksheet
    'using the range dimension here as the previous output table so that we know what data we are using our excel functions on
    Dim outputTable As Range
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim maxIncreaseValue As Double
    Dim maxDecreaseValue As Double
    Dim maxVolumeValue As Double
    
    'doing this for ease of legibility
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    'Set the output table range, same as done for the previous per ticker summation/% table
    Set outputTable = ws.Range("I2:L" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row)
    
    'Find the ticker with the greatest % increase, greatest % decrease, and greatest total volume using the excel functions through Application.WorksheetFunction
    'Functions used are Max, Min
    maxIncreaseValue = Application.WorksheetFunction.Max(outputTable.Columns(3))
    maxIncreaseTicker = outputTable.Cells(Application.WorksheetFunction.Match(maxIncreaseValue, outputTable.Columns(3), 0), 1).Value
    
    maxDecreaseValue = Application.WorksheetFunction.Min(outputTable.Columns(3))
    maxDecreaseTicker = outputTable.Cells(Application.WorksheetFunction.Match(maxDecreaseValue, outputTable.Columns(3), 0), 1).Value
    
    maxVolumeValue = Application.WorksheetFunction.Max(outputTable.Columns(4))
    maxVolumeTicker = outputTable.Cells(Application.WorksheetFunction.Match(maxVolumeValue, outputTable.Columns(4), 0), 1).Value
    
    'Set the headers for the table we are creating
    
    'Column Headers
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    'Left side row headers
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'final code is used to output the values for "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume"
    'Format the values to be either percentage or general (the format functions works with arguments Format(Value, FormatType) to be bassed through cells(_,_).value()
    
    'This is Greatest % increase
    ws.Cells(2, 15).Value = maxIncreaseTicker
    ws.Cells(2, 16).Value = Format(maxIncreaseValue, "0.00%")
    
    'This is Greatest % decrease
    ws.Cells(3, 15).Value = maxDecreaseTicker
    ws.Cells(3, 16).Value = Format(maxDecreaseValue, "0.00%")
    
    'This is total volume
    ws.Cells(4, 15).Value = maxVolumeTicker
    ws.Cells(4, 16).Value = maxVolumeValue
    ws.Cells(4, 16).NumberFormat = "General"
    
End Sub













