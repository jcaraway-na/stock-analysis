Attribute VB_Name = "ClassCodeRefactor_VBA_Script"
Sub AllStocksAnalysisRefactored()
Dim i As Integer
Dim j As Integer
Dim totalVolume As Double
Dim tickers(12) As String
Dim startTime As Single
Dim endTime  As Single
Dim totalSheets As Integer
Dim worksheetExist As Boolean


    
    '.............................UX Improvements.............................................
    '.........................................................................................
    'if user clicks "cancel" on input box control, then
    'macro exit is handled
    yearValue = Trim(Str(Application.InputBox("What year would you like to run the analysis on?", , , , , , , 1)))
        
    If yearValue = False Then
    
        Exit Sub
        
    Else

        'references sheet detector for totalSheets count.
        'loops through worksheets until s = totalSheets.
        'checks to see if enterd worksheet exists.
        For s = 1 To totalSheets
            If ThisWorkbook.Worksheets(s).Name = yearValue Then
                worksheetExist = True
                Exit For
            Else
                worksheetExist = False
            End If
            
        Next s
        
    End If
    '.........................................................................................
    '.........................................................................................
    'if worksheetExist == True then run stock analyzer
    If worksheetExist = True Then
        startTime = Timer
        
        
        '...................................Formatting............................................
        '.........................................................................................
        'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
        
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        '.........................................................................................
        '.........................................................................................
    
        'Initialize array of all tickers
        
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
        'Activate data worksheet
        Worksheets(yearValue).Activate
        
        'Get the number of rows to loop over
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        '1a) Create a ticker Index
        Dim tickerIndex As Integer
        
        tickerIndex = 0
        
        '1b) Create three output arrays
        Dim tickerVolumes(12) As Double
        Dim tickerStartingPrices(12) As Double
        Dim tickerEndingPrices(12) As Double
        
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
                
                ''2b) Loop over all the rows in the spreadsheet.
                For j = 2 To rowCount
            
                        '3a) Increase volume for current ticker
                        If Cells(j, 1).Value = tickers(i) Then
                            
                            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                
                        End If
                        
                        '3b) Check if the current row is the first row with the selected tickerIndex.
                        'If  Then
                        If (Cells(j - 1, 1).Value <> tickers(i) And Cells(j, 1).Value = tickers(i)) Then
                        
                            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                
                        End If
                            
                        'End If
                        
                        '3c) check if the current row is the last row with the selected ticker
                         'If the next row’s ticker doesn’t match, increase the tickerIndex.
                        'If  Then
                        If Cells(j + 1, 1).Value <> tickers(i) And Cells(j, 1).Value = tickers(i) Then
                            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                            '3d Increase the tickerIndex.
                            tickerIndex = tickerIndex + 1
                            
                        End If
                    
                Next j
                
            Next i
        
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For j = 0 To 11
            
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + j, 1).Value = tickers(j)
            Cells(4 + j, 2).Value = tickerVolumes(j)
            Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1
            
        Next j
        
        'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.Bold = True
        Range("A3:C3").Font.Color = RGB(255, 255, 255)
        Range("A3:C3").Interior.Color = RGB(0, 0, 0)
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        Range("B4:B" & rowCount).NumberFormat = "#,##0"
        Range("C4:C" & rowCount).NumberFormat = "0.0%"
        
        Columns("B").AutoFit
    
        dataRowStart = 4
        dataRowEnd = 15
    
        For i = dataRowStart To dataRowEnd
            
            If Cells(i, 3) > 0 Then
                
                Cells(i, 3).Interior.Color = vbGreen
                
            Else
            
                Cells(i, 3).Interior.Color = vbRed
                
            End If
            
        Next i
     
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
    Else
        'if worksheet <> exist then thow message and restart sub
        MsgBox ("Worksheet " + yearValue + " does not exist. Please try again")
        Call AllStocksAnalysisRefactored
            
    End If

End Sub

