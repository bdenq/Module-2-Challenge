Attribute VB_Name = "Module1"
Sub StockSummary()
    
    
'Creation of Column Headers
    'Ticker
        Cells(1, "I").Value = "Ticker"
    'Yearly Change
        Cells(1, "J").Value = "Yearly Change"
    'Percent Change
        Cells(1, "K").Value = "Percent Change"
    'Total Stock Value
        Cells(1, "L").Value = "Total Stock Volume"
    'Autofitting the Columns
        Columns("I:L").AutoFit
    
'Creation of Bonus Section Labels
    'Ticker
        Cells(1, "P").Value = "Ticker"
    'Value
        Cells(1, "Q").Value = "Value"
    'Greatest % Increase
        Cells(2, "O").Value = "Greatest % Increase"
    'Greatest % Decrease
        Cells(3, "O").Value = "Greatest % Decrease"
    'Greatest Total Volumne
        Cells(4, "O").Value = "Greatest Total Volume"
    'Autofitting the Columns
        Columns("O:Q").AutoFit
    
    
'Creation of all variables
    'create a variable to get the last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            'Cells(Rows.Count,1) will start at the very last row
            'End(xlUp) will find a filled row starting from the last row and xlookup to the
            'filled row
            'Row returns the number of the row based on the selected Cell
    'Ticker name Variable
        Dim Ticker As String
    'Yearly Change (1st year opening - last year closing) Variable
        Dim Opening_Value As Double
        Opening_Value = Cells(2, "C").Value
            'need to get opening value of first ticker to start
        Dim Closing_Value As Double
        Dim Yearly_Change As Double
    'Percentage Change Variable
        Dim Percentage_Change As Double
    'Total Stock Value Variable
        Dim Total_Value As Double
        Total_Value = 0
    'Create a variable to print all variables on the excel sheet
        Dim Summary_Row As Integer
        Summary_Row = 2

'Finding unique Ticker names
    For i = 2 To lastrow
        
    'When two Ticker names are not the same
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
        'Ticker Code
            'stores the Ticker name into variable
                Ticker = Cells(i, 1).Value
            'puts the ticker name on the ticker column
                Range("I" & Summary_Row).Value = Ticker
        
        'Total_Value Code
            'Adding the volume (value) of the ticker with the previous tickers
                Total_Value = Total_Value + Cells(i, 7).Value
            'puts the total value of the ticker on the value column
                Range("L" & Summary_Row).Value = Total_Value
            
        'Yearly_Change Code
            'get closing price of last row
                Closing_Value = Cells(i, "F").Value
            'Calculate Yearly Change
                Yearly_Change = Opening_Value - Closing_Value
            'puts Yearly Change on the Yearly Change column
                Range("J" & Summary_Row).Value = Yearly_Change
            'Yearly_Change Interior Color code
                'Creating Conditionals for each color
                'Green color for positives
                    If Range("J" & Summary_Row).Value > 0 Then
                        Range("J" & Summary_Row).Interior.ColorIndex = 4
                'Red color for negatives
                    ElseIf Range("J" & Summary_Row).Value < 0 Then
                        Range("J" & Summary_Row).Interior.ColorIndex = 3
                'Yellow color for no changes
                    ElseIf Range("J" & Summary_Row).Value = 0 Then
                        Range("J" & Summary_Row).Interior.ColorIndex = 6
                    End If
            
        'Percentage_Change Code
            'Calculating the Percentage_Change and formatted to be rounded
                Percentage_Change = Round(((Closing_Value / Opening_Value) - 1) * 100, 2)
            'Putting previous result in the Percentage Change Column and adding a % sign
                Range("K" & Summary_Row).Value = Percentage_Change & "%"
            
        'Reseting for next ticker
            'Creating Opening value for next ticker
                Opening_Value = Cells(i + 1, "C").Value
            'adds a one to the Summary_Row for the next row
                Summary_Row = Summary_Row + 1
            'Reset the Total_value counter for the next row
                Total_Value = 0
        
    'when the ticker names are the same
        Else
            
        'Adding the volume (value) of the ticker with the previous tickers
            Total_Value = Total_Value + Cells(i, 7).Value
            
        End If
    Next i

'Bonus Section
    'Greatest % Increase
        'Create Variable
            Dim Greatest_Increase As Double
        'Finding Greatest % Increase in K Column and Formatting
            Greatest_Increase = WorksheetFunction.Max(Range("K:K")) * 100
            Cells(2, "Q").Value = Greatest_Increase & "%"
        'Ticker for Greatest % Increase and putting it in table
            For i = 2 To lastrow
                If Cells(i, "K").Value = Cells(2, "Q").Value Then
                    Cells(2, "P").Value = Cells(i, "I").Value
                End If
            Next i
    'Greatest % Decrease
        'Create Variable
            Dim Greatest_Decrease As Double
        'Finding Greatest % Decrease in K Column and Formatting
            Greatest_Decrease = WorksheetFunction.Min(Range("K:K")) * 100
            Cells(3, "Q").Value = Greatest_Decrease & "%"
        'Ticker for Greatest % Decrease and putting it in table
            For i = 2 To lastrow
                If Cells(i, "K").Value = Cells(3, "Q").Value Then
                    Cells(3, "P").Value = Cells(i, "I").Value
                End If
            Next i
    'Greatest Total Volume
        'Create Variable and set initial value
            Dim Greatest_Total As Variant
            Greatest_Total = Cells(2, "L").Value
        'Finding Greatest Total Volume
            For i = 2 To lastrow
                If Cells(i + 1, "L").Value > Greatest_Total Then
                    Greatest_Total = Cells(i + 1, "L").Value
                End If
            Next i
            'putting the Greatest Total Volume in the table
            Cells(4, "Q").Value = Greatest_Total
        'Ticker for Greatest % Increase and putting it in table
            For i = 2 To lastrow
                If Cells(i, "L").Value = Cells(4, "Q").Value Then
                    Cells(4, "P").Value = Cells(i, "I").Value
                End If
            Next i
    'AutoFit
        Columns("O:Q").AutoFit
    
    'SCRAP Greatest Total Volume
        'Create Variable
            'Dim Greatest_Total As Variant
        'Finding Greatest % Increase in K Column and Formatting
            'Greatest_Total = Format(WorksheetFunction.Max(Range("L:L")), "Scientific")
            'Cells(4, "Q").Value = Greatest_Total
        'Ticker for Greatest % Increase and putting it in table
            'For i = 2 To lastrow
                'If Cells(i, "L").Value = Cells(4, "Q").Value Then
                    'Cells(4, "P").Value = Cells(i, "I").Value
                'End If
            'Next i


End Sub
