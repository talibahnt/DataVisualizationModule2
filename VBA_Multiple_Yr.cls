VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis():

'Loop this script through all of the four worksheets'
    For Each ws In Worksheets

'Create a script that will loop through all of the stock per quarter (for all four worksheets) and return stock statisitics'
    Dim Stock_Ticker As String

'Define a variable to hold th total stock volume for each stock ticker'
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

'Define a variable and value of stock statistics' summary table'
    Dim Summary_of_Stock_Table_Row As Long
    Summary_of_Stock_Table_Row = 2

'Define variables for quarter open, quarter close, and quarter changes in price'
    Dim Quarter_Open As Double
    Dim Quarter_Close As Double
    Dim Quarter_Change As Double

    Dim Stock_Amount As Long
    Stock_Amount = 2

'Define percentate change and the last row'

    Dim Last_Row As Long
    Dim Last_Row_Value As Long

'Loop through the last row in all worksheets. Create a loop.'

    For i = 2 To Last_Row

'Find and return a the total volume  that it is associated with'

'Calculate the total volume for each stock ticker'
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

'Set a conditional to find if the information  is within the same stock ticker'
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    Stock_Ticker = ws.Cells(i, 1).Value

'Print output in corresponding spaces in summary table'

    ws.Range("I" & Summary_of_Stock_Table_Row).Value = Stock_Ticker
    ws.Range("L" & Summary_of_Stock_Table_Row).Value = Total_Stock_Volume

'Reset total volume for each ticker'
    Total_Stock_Volume = 0

'Find yearly change in each stock ticker'
    Quarter_Open = ws.Range("C" & Stock_Amount)
    Quarter_Close = ws.Range("F" & i)
    Quarter_Change = Quarter_Close - Quarter_Open

'Print output in the Stock Statistics' Table
    ws.Range("J" & Summary_of_Stock_Table_Row).Value = Quarter_Change

'Find the percentage change for each stock ticker. Create an if-else conditional.'
If Quarter_Open = 0 Then
    Percentage_Change = 0
        Else
    Quarter_Open = ws.Range("C" & Stock_Amount)
    Percentage_Change = Quarter_Change / Quarter_Open
End If

'Formatting changes to make them more presentable. Number Format to change the type of number.'
    ws.Range("K" & Summary_of_Stock_Table_Row).NumberFormat = "%"
    ws.Range("K" & Summary_of_Stock_Table_Row).Value = Percentage_Change

'Use conditional formatting to color code positive results green and negative results red'

If ws.Range("J" & Summary_of_Stock_Table_Row).Value >= 0 Then
    ws.Range("J" & Summary_of_Stock_Table_Row).InteriorColorIndex = 4
Else
    ws.Range("J" & Summary_of_Stock_Table_Row).InteriorColorIndex = 3
End If

'Reset Stock Statistics' Summary Table Row so that looping can continue'
    Summary_of_Stock_Table_Row = Summary_of_Stock_Table_Row + 1
    Stock_Amount = i + 1
 End If
Next i

'Assignment Solutions'


'Redefine the last row variable'
Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Create solutions for each quarter. Find the "Greatest % Increase," "Greatest % Decrease", and the "Greatest Total Stock Volume"
For i = 2 To Last_Row
    If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
       ws.Range("Q2").Value = ws.Range("K" & i).Value
       ws.Range("P2").Value = ws.Range("I" & i).Value
    End If
    
'Create a for loop to make a conditional for the greatest % decrease'
    If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Range("K" & i).Value
    ws.Range("P3").Value = ws.Range("I" & i).Value
    
'Create a for loop to make a condtional for the greatest % increase'
    If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
    ws.Range("Q4").Value = ws.Range("K" & i).Value
    ws.Range("P4").Value = ws.Range("I" & i).Value
    End If

Next i
'Format the results of the greatest increase and decrease to  percentages using Number Format'
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

    Next ws



End Sub
