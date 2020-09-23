Attribute VB_Name = "Module1"
Option Explicit

Sub stonks():

    'Variables
    Dim i, j, k As Integer
    Dim cType As String
    Dim openTot, closeTot, vol, dailyChange, percentChange As Double
    Dim ws As Worksheet
    
    'Loop through each worksheet in workbook
    For Each ws In Worksheets
    
    'Assign values to variables to begin with, at the start of each worksheet loop
        k = 2
        openTot = 0
        closeTot = 0
        vol = 0
        dailyChange = 0
        percentChange = 0

        'Assign column header values for stock list
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"

        'Pre format column to be percentages
        ws.Columns("K:K").NumberFormat = "0.00%"

        'Loop through every row in the given data sets
        For i = 2 To ws.Range("A1").CurrentRegion.Rows.Count
            
            'Conditional to test if the stock name is equal to the stock name in the row directly beneath it
            'If the value does not equal the value beneath it
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                
                'Run another conditional to test if stock name is the first time it appears in the list AND if that value equals 0
                'If true
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) And ws.Cells(i, 3).Value <> 0 Then
                     'Record open value in variable
                     openTot = ws.Cells(i, 3).Value
                'If first appearnce of stock equals 0, keep running until first first open value that does not equal 0
                ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1) And ws.Cells(i - 1, 3).Value = 0 Then
                     openTot = ws.Cells(i, 3).Value
                End If
                
                'Assign variable to last close value in stock set
                closeTot = ws.Cells(i, 6).Value
                'Assign variable for the difference between the stocks open and closed values
                dailyChange = ws.Cells(i, 6) - ws.Cells(i, 3)
                'Add onto volumn as long as the stock is the same
                vol = vol + ws.Cells(i, 7).Value
                
                'Assign ticker value & yearly change to incremented row
                ws.Cells(k, 9) = ws.Cells(i, 1).Value
                ws.Cells(k, 10) = closeTot - openTot
                
                
                If openTot = 0 Then
                     percentChange = 0
                Else
                     percentChange = ws.Cells(k, 10) / openTot
                End If

                ws.Cells(k, 11) = percentChange
               ws.Cells(k, 12) = vol
                
                'Conditional to test if the cell should be colored green or red
                If ws.Cells(k, 10) > 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 3
                End If
                
                'Increment k value to place new ticker beneath the previous in output list
                k = k + 1
                'Reset values for next ticker in loop
                openTot = 0
                closeTot = 0
                vol = 0
                dailyChange = 0
                percentChange = 0
                
            'If ticker value beneath current ticker does equal the ticker
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) And ws.Cells(i, 1).Value <> 0 Then
                    openTot = ws.Cells(i, 3).Value
                ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1) And ws.Cells(i - 1, 3).Value = 0 Then
                    openTot = ws.Cells(i, 3).Value
                End If
                closeTot = ws.Cells(i, 6).Value
                dailyChange = ws.Cells(i, 6) - ws.Cells(i, 3)
                If openTot = 0 Then
                   percentChange = 0
                Else
                   percentChange = ws.Cells(k, 10) / openTot
                End If
                vol = vol + ws.Cells(i, 7).Value

            End If
            
        Next i
           
           'Assign header values for stock stats
           ws.Range("N2").Value = "Greatest % Increase"
           ws.Range("N3").Value = "Greatest % Decrease"
           ws.Range("N4").Value = "Greatest Volume Increase"
           ws.Range("O1").Value = "Ticker"
           ws.Range("P1").Value = "Value"
    
           'Get max value of percent change & volume and min of percent change
           ws.Cells(2, 16) = WorksheetFunction.Max(ws.Range("K:K"))
           ws.Cells(3, 16) = WorksheetFunction.Min(ws.Range("K:K"))
           ws.Cells(4, 16) = WorksheetFunction.Max(ws.Range("L:L"))
           
           'Pre format percentages for min and max percentages
           ws.Range("P2:P3").NumberFormat = "0.00%"
    
           'Loop through created data set to find ticker for the min and maxes
           For j = 2 To ws.Range("I1").CurrentRegion.Rows.Count
                
              'Conditional to test where the value is in new list,take ticker value & place in required destination
              If ws.Cells(j, 11).Value = ws.Range("P2") Then
                  ws.Range("O2").Value = ws.Cells(j, 11).Offset(0, -2).Value
              ElseIf ws.Cells(j, 11).Value = ws.Range("P3") Then
                  ws.Range("O3").Value = ws.Cells(j, 11).Offset(0, -2).Value
              ElseIf ws.Cells(j, 12).Value = ws.Range("P4") Then
                  ws.Range("O4").Value = ws.Cells(j, 12).Offset(0, -3).Value
              End If
        
           Next j
           
           ws.UsedRange.Columns.AutoFit
           
        Next ws

                
        

End Sub

