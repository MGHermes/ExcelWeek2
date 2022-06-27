Attribute VB_Name = "Module11"
'Task 1: Record each of the unique ticker symbols
'Tast 2: Record sum of G column/col 7 values (volume)
'Task 3: Record diff of first and last price
'Task 4: Record percent diff of first and last price

Sub WallStreet()
        
    'create variables for the information to be recorded
        'create a variable for current ticker
        ticker = Range("A2")
        'create a variable for opening price
        openingPrice = Range("C2")
        'create a variable for closing prince
        closingPrice = Range("F2")
        'create a variable for the price difference
        Dim priceDifference As Double
        'ceate a variable for the percent change
        Dim percentChange As Double
        'create a variable for the totalVolume
        totalVolume = 0
        
    'create user input box
    yearValue = InputBox("Enter a year: 2018, 2019, or 2020")
    Worksheets(yearValue).Activate
    
    'create/ format the summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    'create a variable for the summary table row
    tickerRow = 2
    
    'create a variable for last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'create a loop through the rows
    For Row = 2 To lastRow
    
        'add the volume of the current row
        totalVolume = totalVolume + Cells(Row, 7)
                
        'check to see if the next ticker is different
        If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
                
            'if the next ticker is different...
            'collect the important information about this ticker group
                'record the current ticker
                ticker = Cells(Row, 1).Value
                'record the closing price
                closingPrice = Cells(Row, 6).Value
                'calculate the difference between the opening price and the closing price
                priceDifference = closingPrice - openingPrice
                'calculate the percent change between the opening price and the closing price
                If (openingPrice <> 0) Then
                    percentChange = (closingPrice - openingPrice) / openingPrice
                    Else
                        Cells(tickerRow, 11).Value = "ERROR"
                        MsgBox ("Division by zero on row # " + Row)
                End If
                    
            'print this ticker group information in the summary table
                'print the current ticker
                Cells(tickerRow, 9).Value = ticker
                'print the difference in price
                Cells(tickerRow, 10).Value = priceDifference
                'print the percent change in price
                Cells(tickerRow, 11).Value = percentChange
                'print the total volume
                Cells(tickerRow, 12).Value = totalVolume
                        
            'change the style of the information in the summary table
                'change the style of the percent change to percent
                Cells(tickerRow, 11).NumberFormat = "0.00%"
                'if the yearly change is negative, color the cell red
                If priceDifference < 0 Then
                    Cells(tickerRow, 10).Interior.ColorIndex = 3
                    Cells(tickerRow, 10).Font.ColorIndex = 30
                'If the yearly change is positive, color the cell green
                ElseIf priceDifference > 0 Then
                    Cells(tickerRow, 10).Interior.ColorIndex = 4
                    Cells(tickerRow, 10).Font.ColorIndex = 52
                Else
                    Cells(tickerRow, 10).Interior.ColorIndex = 27
                End If
                    
            'If the next ticker on the list is not an empty cell...
            'reset the opening price to the next opening price on the list
            openingPrice = Cells(Row + 1, 3)
            'reset the totalVolume
            totalVolume = 0
            'add 1 to the ticker row
            tickerRow = tickerRow + 1
                
        End If
    
    'end the loop through the rows
    Next Row
    
    'determine the ticker with the largest % increase, % decrease, and total volume
        
        'determine the last row
        lastRow = Cells(Rows.Count, 9).End(xlUp).Row
        'define a variable for largest percent increase
        bigIncrease = Range("K2").Value 'column 11
        bigIncreaseName = Range("I2").Value 'column 9
        'define a variable for largest percent decrease
        bigDecrease = Range("K2").Value 'column 11
        bigDecreaseName = Range("I2").Value 'column 9
        'define a variable for largest total volume
        bigVolume = Range("L2").Value 'column 12
        bigVolumeName = Range("I2").Value 'column 9
        
        'search the rows in the summary table
        For Row = 2 To lastRow
        
            If Cells(Row, 11) > bigIncrease Then
                bigIncrease = Cells(Row, 11).Value
                bigIncreaseName = Cells(Row, 9).Value
            End If
            
            If Cells(Row, 11) < bigDecrease Then
                bigDecrease = Cells(Row, 11).Value
                bigDecreaseName = Cells(Row, 9).Value
            End If
            
            If Cells(Row, 12) > bigVolume Then
                bigVolume = 0
                bigVolume = Cells(Row, 12).Value
                bigVolumeName = Cells(Row, 9).Value
            End If
        
        Next Row
        
        'create the bonus summary table
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        'Print the name of the ticker with the greatest % increase
        Range("O2").Value = bigIncreaseName
        'Print the name of the ticker with the greatest % decrease
        Range("O3").Value = bigDecreaseName
        'Print the name of the ticker with the greatest total volume
        Range("O4").Value = bigVolumeName
        'Print the value of the ticker with the greatest % increase
        Range("P2").Value = bigIncrease
        Range("P2").NumberFormat = "0.00%"
        'Print the value of the ticker with the greatest % decrease
        Range("P3").Value = bigDecrease
        Range("P3").NumberFormat = "0.00%"
        'Print the value of the ticker with the greatest total volume
        Range("P4").Value = bigVolume
    
End Sub
