Attribute VB_Name = "Module11"
Sub Stocks():
    
    On Error Resume Next

    'Define the ticker value
    Dim Ticker As String
    'Define the index for rows when printing results
    Dim j As Integer
    'Define Stock Volume to allow for large totals
    Dim StockVolume As LongLong
    'Define The opening value of the year to allow decimals
    Dim YearOpen As Double
    'Define the closing value of the year to allow decimals
    Dim YearClose As Double
    
    'set variables for the challenge section of the code
    Dim MaxChange As Double
    Dim MinChange As Double
    Dim MaxVolume As LongLong
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim VolTicker As String

For Each ws In Worksheets
    
        'set the initial value for Ticker
        Ticker = ws.Range("A2")
        
        'set the printing index to start in the second row
        j = 2
        
        'set the opening values for Variables
        YearOpen = ws.Range("C2")
        YearClose = 0
        StockVolume = 0
    
        'Find the last non-blank row in column A(1)
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'label the output columns
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        'MsgBox (LastRow)
        For i = 2 To LastRow:
        
            'check to see if the next symbol is the same
            If Ticker = ws.Cells(i, 1) Then
                
                'add up the stock volume
                StockVolume = StockVolume + ws.Cells(i, 7)
    
                       
            Else
                'Print out the  ticker symbol in column i
                ws.Cells(j, 9) = Ticker
                
                
                'Print out the total stock volume
                ws.Cells(j, 12) = StockVolume
                
                
                'Set the new YearClose Value
                YearClose = ws.Cells(i - 1, 6)
                
                
                'Calculate and print the Yearly Change
                ws.Cells(j, 10).Value = YearClose - YearOpen
                
                'Format the cell color based on whether the value is positive or negative
                If ws.Cells(j, 10) > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                
                End If
                    
                
                'Calculate and print the % change for the year
                ws.Cells(j, 11).NumberFormat = "0.0%"
                
                If ws.Cells(j, 10) = 0 Then
                    ws.Cells(j, 11) = 0
                    
                Else
                    ws.Cells(j, 11) = (YearClose / YearOpen) - 1
                    
                End If
                
                'change the row where the value will be printed
                j = j + 1
                
                'Update the saved ticker symbol
                Ticker = ws.Cells(i, 1)
                
                'Reset StockVolume to Zero
                StockVolume = 0
                
                'Update YearOpen to the new value
                YearOpen = ws.Cells(i, 3)
    
            End If
        Next i
    
    'Challenge section of code
    
        'Add the column headers
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
        
        'Add the Row headers
        ws.Range("N2") = "Greatest % Increase"
        ws.Range("N3") = "Greatest % Decrease"
        ws.Range("N4") = "Greatest Total Volume"
        
        'Reset variable values
        MaxChange = 0
        MinChange = 0
        MaxVolume = 0
        MinTicker = ""
        MaxTicker = ""
        VolTicker = ""
        
        'Find the last non-blank row in column I(9)
        NewLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For k = 2 To NewLastRow
        
            If MaxChange < ws.Cells(k, 11) Then
                MaxChange = ws.Cells(k, 11)
                MaxTicker = ws.Cells(k, 9)
            End If
            
            If MinChange > ws.Cells(k, 11) Then
                MinChange = ws.Cells(k, 11)
                MinTicker = ws.Cells(k, 9)
            End If
            
            If MaxVolume < ws.Cells(k, 12) Then
                MaxVolume = ws.Cells(k, 12)
                VolTicker = ws.Cells(k, 9)
            End If
        
        Next k
        'configure cells and print out values
        ws.Range("P2").NumberFormat = "0.0%"
        ws.Range("P2") = MaxChange
        ws.Range("O2") = MaxTicker
            
        ws.Range("P3").NumberFormat = "0.0%"
        ws.Range("P3") = MinChange
        ws.Range("O3") = MinTicker
            
        ws.Range("P4") = MaxVolume
        ws.Range("O4") = VolTicker
        
        'Auto-Fit the cells to the contents
        ws.Columns("A:P").EntireColumn.AutoFit
        
Next ws
    
End Sub
