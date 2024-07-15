Attribute VB_Name = "VBACharts"
Sub StockAnalysis()

'Create script that will loop through all of the stock per quarter (for all four worksheets) and return stock statistics'

'Loop this script through all of the four worksheets'
    For Each ws In Worksheets

        'Name all column headers for output in the Stock Statistics Summary Table'
        ws.Range("I1").Value = "Stock Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Stock Statistics Summary Table"
        
        'Name all spaces for output in the Stock Statistics Summary Table'
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Stock Volume"
        ws.Range("P1").Value = "Stock Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Format Stock Statistics Summary Table to auto fit for expected large results'
        ws.Columns("I:Q").AutoFit
        

        'Define initial variables'
        
        Dim stock_ticker As String
       
       'Define the variable to hold total stock volume for each stock ticker'
       
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        'Define variable and value of Stock Statistics Summary Table'
        
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
         
       'Define variables for quarter open, quarter close, and quarter change prices'
       
        Dim Quarter_Open As Double
        Dim Quarter_Close As Double
        Dim Quarter_Change As Double
        
        Dim Stock_Amount As Long
        Stock_Amount = 2
        
        'Define percentage change and last row'
        
        Dim Percentage_Change As Double
        
        Dim Last_Row As Long
        Dim Last_Row_Value As Long
       

        'Find the last row through all worksheets'
        
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through last row in all of the four worksheets'
        
        For i = 2 To Last_Row

'Find and return total stock volume that it is associated with'

            'Calculate total stock volume for each stock ticker'
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            'Set a conditional to find the information within the same stock ticker'
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            stock_ticker = ws.Cells(i, 1).Value
            
                'Print output in Stock Statistics Summary Table'
                
                ws.Range("I" & Summary_Table_Row).Value = stock_ticker
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'Reset total stock volume to zero so the loop can continue'
                Total_Stock_Volume = 0


    'Find quarterly change in each stock ticker using quarter open, quarter close, and quarter change'
    
                'Set values for quarter open, quarter close and quarter change'
                Quarter_Open = ws.Range("C" & Stock_Amount)
                Quarter_Close = ws.Range("F" & i)
                Quarter_Change = Quarter_Close - Quarter_Open
                
                'Print quarter change in Stock Statistics Summary Table'
                ws.Range("J" & Summary_Table_Row).Value = Quarter_Change

                'Find percentage change for each stock ticker'
                If Quarter_Open = 0 Then
                    Percentage_Change = 0
                Else
                    Quarter_Open = ws.Range("C" & Stock_Amount)
                    Percentage_Change = Quarter_Change / Quarter_Open
                End If
                
'Formatting changes'
                
                'Formatting percentge change to make value more presentatble. Use Number Format to change the number type'
                
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Table_Row).Value = Percentage_Change

               'Use conditional formatting to color code positive results green and negative results red'
               
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
                'Reset summary table row so looping can occur'
                Summary_Table_Row = Summary_Table_Row + 1
                Stock_Amount = i + 1
                End If
            Next i

'Assignment Solutions'

'Create solutions for each quarter. Find the "Greatest % increase", "Greatest % decrease" and "Greatest Total Stock Volume'

            'Redefine the last row lastrow'
            Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            'Greatest % increase Conditional'
            For i = 2 To Last_Row
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If
            'Greatest % decrease Conditional'
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If
            'Greatest Total Stock Volume Conditional'
                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
            
        'Format results of greatest increase and decrease so that it presents as a percentage'
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
            
       

    Next ws

End Sub

Sub InputChartNewSheets()

Dim chrt As chart
Set chrt = Charts.Add
With chrt
    .SetSourceData Source:=Sheets("Q1").Range("O2:O4", "Q2:Q4")
    .ChartType = xlLine
End With

End Sub


